import type { OutlookGenConfig } from "../../../types/dataDump.js";
import { addMailAttachment, createCalendarEvent, createContact, createMailItem } from "../../../graph/dataDump/outlookCreation.js";
import { runThrottled } from "../../rateLimiter.js";
import { createSeededSource, forkSeed, pick, randomInt } from "../seededRandom.js";
import { buildEmailBody, buildEventBody, randomCompanyName, randomEmailSubject, randomEventTitle, randomJobTitle, randomPersonName, randomPhoneNumber } from "../businessContent.js";
import { generateFileContent } from "../fileContent.js";
import { getOrCreateContainer } from "../containers.js";
import { BatchAccumulator, nextBatchIndexFor } from "../batching.js";
import { applyTaskProgress, bumpSubcount, setTaskCheckpoint } from "../progress.js";
import { resolveDateRange, randomTimestampInRange } from "../dateRange.js";
import type { RunnerContext, RunnerOutcome } from "./oneDriveRunner.js";

const ATTACHMENT_TYPES = ["docx", "xlsx", "pdf", "png"] as const;

/**
 * Generates real Outlook mail items (spec §16's CRITICAL SAFETY REQUIREMENT: created directly inside
 * a mailbox folder — see graph/dataDump/outlookCreation.ts — never sent, never reaches an external
 * address), plus, independently (spec §21/§22), real calendar events and contacts per selected user.
 * Each of the three sub-generators is optional per-user (count 0 skips it) and accounted separately.
 */
export async function runOutlookWorkload(ctx: RunnerContext, users: { id: string; upn: string; displayName: string }[], config: OutlookGenConfig): Promise<RunnerOutcome> {
  let hadFailures = false;
  let globalEmailIndex = 0;
  let batchIndex = await nextBatchIndexFor(ctx.taskId);
  const dateRangeWindow = resolveDateRange(ctx.dateRange);

  for (const user of users) {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";

    const mailboxContainer = await getOrCreateContainer(ctx.taskId, "mailbox_folder", `${ctx.namingPrefix} - Outlook - ${user.upn}`, null, async () => ({ graphId: user.id }));

    if (config.emailsPerUser > 0) {
      const signal = await generateEmailsForUser(ctx, user, mailboxContainer.id, config, dateRangeWindow, batchIndex, () => globalEmailIndex++);
      if (signal.stop) return signal.stop;
      batchIndex = signal.nextBatchIndex;
      if (signal.hadFailures) hadFailures = true;
    }

    if (config.calendarEventCount > 0) {
      const result = await generateCalendarEventsForUser(ctx, user, users, config, dateRangeWindow);
      if (result === "cancelled" || result === "paused") return result;
      if (result === "hadFailures") hadFailures = true;
    }

    if (config.contactCount > 0) {
      const result = await generateContactsForUser(ctx, user, config);
      if (result === "cancelled" || result === "paused") return result;
      if (result === "hadFailures") hadFailures = true;
    }
  }

  return hadFailures ? "completed_with_errors" : "completed";
}

async function generateEmailsForUser(
  ctx: RunnerContext,
  user: { id: string; upn: string; displayName: string },
  mailboxContainerId: string,
  config: OutlookGenConfig,
  dateRangeWindow: { startMs: number; endMs: number },
  startBatchIndex: number,
  nextEmailIndex: () => number
): Promise<{ stop?: "cancelled" | "paused"; nextBatchIndex: number; hadFailures: boolean }> {
  const emailIndices = Array.from({ length: config.emailsPerUser }, () => nextEmailIndex());
  const settled: { ok: boolean; name: string; graphId: string | null; sizeBytes: number }[] = [];

  await runThrottled(
    emailIndices,
    async (emailIndex) => {
      const emailSeed = createSeededSource(forkSeed(ctx.operationSeed, `outlook:email:${emailIndex}`));
      const subject = `${ctx.namingPrefix}: ${randomEmailSubject(emailSeed)}`;
      const body = buildEmailBody(emailSeed, subject);
      const folder = emailIndex % 2 === 0 ? "sentitems" : "inbox";
      const created = await createMailItem(ctx.client, user.id, folder, {
        subject,
        body,
        fromUpn: folder === "sentitems" ? user.upn : "demo.sender@example.com",
        toUpn: folder === "sentitems" ? "demo.recipient@example.com" : user.upn,
      });

      let attachmentBytes = 0;
      if (emailSeed.rng() < config.attachmentsPerEmail) {
        const attachmentType = pick(emailSeed.rng, ATTACHMENT_TYPES);
        const generated = await generateFileContent(ctx.operationSeed, emailIndex, attachmentType, config.namingStyle);
        await addMailAttachment(ctx.client, user.id, created.id, generated.fileName, generated.buffer, generated.mimeType);
        attachmentBytes = generated.buffer.length;
      }

      return { id: created.id, sizeBytes: attachmentBytes };
    },
    {
      label: "DataDump-Outlook",
      batchSize: 10,
      isCancelled: ctx.isCancelled,
      onItemSettled: (_emailIndex, result) => {
        if (result.ok) settled.push({ ok: true, name: `email-${result.value.id}`, graphId: result.value.id, sizeBytes: result.value.sizeBytes });
        else settled.push({ ok: false, name: "email-failed", graphId: null, sizeBytes: 0 });
      },
    }
  );

  const accumulator = new BatchAccumulator(ctx.taskId, mailboxContainerId, startBatchIndex);
  let createdCount = 0;
  let failedCount = 0;
  let sizeBytesSum = 0;
  for (const item of settled) {
    await accumulator.add({ name: item.name, graphId: item.graphId, sizeBytes: item.sizeBytes, status: item.ok ? "created" : "failed" });
    if (item.ok) {
      createdCount++;
      sizeBytesSum += item.sizeBytes;
    } else {
      failedCount++;
    }
  }
  await accumulator.flush();
  await bumpSubcount(ctx.taskId, "email", { requested: settled.length, created: createdCount, failed: failedCount });
  await applyTaskProgress(ctx.taskId, ctx.operationId, { created: createdCount, failed: failedCount, skipped: 0, sizeBytes: sizeBytesSum });
  await setTaskCheckpoint(ctx.taskId, { phase: "emails", nextBatchIndex: accumulator.currentBatchIndex, lastUserId: user.id });

  return { nextBatchIndex: accumulator.currentBatchIndex, hadFailures: failedCount > 0 };
}

async function generateCalendarEventsForUser(
  ctx: RunnerContext,
  user: { id: string; upn: string; displayName: string },
  allUsers: { id: string; upn: string; displayName: string }[],
  config: OutlookGenConfig,
  dateRangeWindow: { startMs: number; endMs: number }
): Promise<"cancelled" | "paused" | "ok" | "hadFailures"> {
  if (await ctx.isCancelled()) return "cancelled";
  if (await ctx.isPaused()) return "paused";

  let created = 0;
  let failed = 0;
  for (let i = 0; i < config.calendarEventCount; i++) {
    const seed = createSeededSource(forkSeed(ctx.operationSeed, `outlook:event:${user.id}:${i}`));
    const title = randomEventTitle(seed);
    const startMs = dateRangeWindow.startMs + seed.rng() * (dateRangeWindow.endMs - dateRangeWindow.startMs);
    const start = new Date(startMs);
    const end = new Date(startMs + (30 + randomInt(seed.rng, 0, 3) * 30) * 60_000);
    const attendeeUpns = config.includeAttendees ? allUsers.filter((u) => u.id !== user.id).slice(0, 2).map((u) => u.upn) : undefined;
    try {
      await createCalendarEvent(ctx.client, user.id, {
        subject: `${ctx.namingPrefix}: ${title}`,
        body: buildEventBody(seed, title),
        startIso: start.toISOString(),
        endIso: end.toISOString(),
        attendeeUpns,
        recurrence: config.includeRecurringEvents && i % 5 === 0 ? { type: "weekly", interval: 1, endDateIso: new Date(startMs + 60 * 24 * 60 * 60_000).toISOString() } : undefined,
      });
      created++;
    } catch {
      failed++;
    }
  }
  await bumpSubcount(ctx.taskId, "calendar_event", { requested: config.calendarEventCount, created, failed });
  await applyTaskProgress(ctx.taskId, ctx.operationId, { created, failed, skipped: 0, sizeBytes: 0 });
  return failed > 0 ? "hadFailures" : "ok";
}

async function generateContactsForUser(ctx: RunnerContext, user: { id: string; upn: string; displayName: string }, config: OutlookGenConfig): Promise<"cancelled" | "paused" | "ok" | "hadFailures"> {
  if (await ctx.isCancelled()) return "cancelled";
  if (await ctx.isPaused()) return "paused";

  let created = 0;
  let failed = 0;
  for (let i = 0; i < config.contactCount; i++) {
    const seed = createSeededSource(forkSeed(ctx.operationSeed, `outlook:contact:${user.id}:${i}`));
    const name = randomPersonName(seed);
    const company = randomCompanyName(seed);
    try {
      await createContact(ctx.client, user.id, {
        givenName: name.first,
        surname: name.last,
        emailAddress: `${name.first}.${name.last}@${company.toLowerCase().replace(/[^a-z0-9]/g, "")}.example.com`,
        companyName: company,
        jobTitle: randomJobTitle(seed),
        businessPhone: randomPhoneNumber(seed),
      });
      created++;
    } catch {
      failed++;
    }
  }
  await bumpSubcount(ctx.taskId, "contact", { requested: config.contactCount, created, failed });
  await applyTaskProgress(ctx.taskId, ctx.operationId, { created, failed, skipped: 0, sizeBytes: 0 });
  return failed > 0 ? "hadFailures" : "ok";
}
