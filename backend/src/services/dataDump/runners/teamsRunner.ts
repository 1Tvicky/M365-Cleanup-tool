import type { ChannelConfig, TeamsGenConfig } from "../../../types/dataDump.js";
import { addChannelMember, addTeamMember, createChannel, createTeam } from "../../../graph/dataDump/teamsCreation.js";
import { createSeededSource, forkSeed } from "../seededRandom.js";
import { randomDepartment, channelNamesForTeam } from "../businessContent.js";
import { getOrCreateContainer } from "../containers.js";
import { applyTaskProgress, bumpSubcount, setTaskCheckpoint } from "../progress.js";
import type { RunnerContext, RunnerOutcome } from "./oneDriveRunner.js";

interface TargetTeam {
  containerId: string;
  graphId: string;
  isNew: boolean;
}

/**
 * Generates real Teams + channels + memberships (existing tenant users only). Supports both target
 * shapes from the Select Teams / Create New Team step (spec §12): `config.selectedTeamIds` adds
 * configured channels/members to EXISTING teams, `config.newTeam` creates one brand-new team first —
 * both may be set in the same run. `config.channels` gives explicit per-channel name/members/
 * messages/replies; when absent, falls back to the legacy count-based teamCount/channelsPerTeam
 * quick-start path so profile-driven generation (no manual selection) still works unchanged.
 *
 * Channel messages and replies are counted as "skipped" (spec §27's accounting) — see
 * graph/dataDump/teamsCreation.ts's postChannelMessageUnsupported for why they are never attempted
 * as live Graph calls.
 */
export async function runTeamsWorkload(
  ctx: RunnerContext,
  availableUsers: { id: string; upn: string; displayName: string }[],
  config: TeamsGenConfig,
  selectedTeams: { id: string; displayName: string }[] = []
): Promise<RunnerOutcome> {
  let hadFailures = false;
  const usingExplicitSelection = selectedTeams.length > 0 || !!config.newTeam;
  const userById = new Map(availableUsers.map((u) => [u.id, u]));

  const targets: TargetTeam[] = [];

  // Existing teams selected from the Select Teams step — the container already IS the real Graph
  // team, so it's recorded (if not already) rather than created.
  for (const team of selectedTeams) {
    const container = await getOrCreateContainer(ctx.taskId, "team", team.displayName, null, async () => ({ graphId: team.id }));
    targets.push({ containerId: container.id, graphId: team.id, isNew: false });
  }

  // A brand-new team, created once.
  if (config.newTeam) {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";
    const owner = config.newTeam.memberUserIds.map((id) => userById.get(id)).find(Boolean) ?? availableUsers[0];
    await bumpSubcount(ctx.taskId, "team", { requested: 1 });
    if (!owner) {
      hadFailures = true;
      await bumpSubcount(ctx.taskId, "team", { failed: 1 });
      await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
    } else {
      try {
        const container = await getOrCreateContainer(ctx.taskId, "team", config.newTeam.displayName, null, async () => {
          const team = await createTeam(ctx.client, config.newTeam!.displayName, config.newTeam!.displayName.replace(/[^a-zA-Z0-9]/g, "").slice(0, 60) || "ddteam", owner.id, {
            visibility: config.newTeam!.visibility,
            description: config.newTeam!.description,
          });
          return { graphId: team.groupId };
        });
        targets.push({ containerId: container.id, graphId: container.graphId, isNew: true });
        await bumpSubcount(ctx.taskId, "team", { created: 1 });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 1, failed: 0, skipped: 0, sizeBytes: 0 });

        for (const memberId of config.newTeam.memberUserIds) {
          if (memberId === owner.id) continue;
          await bumpSubcount(ctx.taskId, "user", { requested: 1 });
          try {
            await addTeamMember(ctx.client, container.graphId, memberId);
            await bumpSubcount(ctx.taskId, "user", { created: 1 });
          } catch {
            await bumpSubcount(ctx.taskId, "user", { skipped: 1 }); // most common cause: already a member (e.g. a resumed run) — not a real failure
          }
        }
      } catch (err) {
        hadFailures = true;
        await bumpSubcount(ctx.taskId, "team", { failed: 1 });
        const message = err instanceof Error ? err.message : String(err);
        console.error(`[data-dump] Teams: new team ${config.newTeam.displayName} failed`, message);
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
      }
    }
  }

  // Legacy quick-start path: no manual selection at all (profile-driven Basic/Migration/Enterprise/
  // Performance without a Select Teams / Create New Team step) — creates config.teamCount brand-new
  // teams the same way earlier versions of this runner always did, so that path keeps working
  // unchanged. Once any explicit selection exists, this branch never runs (targets are exactly what
  // the operator picked, no more, no less).
  if (!usingExplicitSelection) {
    for (let teamIndex = 0; teamIndex < config.teamCount; teamIndex++) {
      if (await ctx.isCancelled()) return "cancelled";
      if (await ctx.isPaused()) return "paused";

      const teamSeed = createSeededSource(forkSeed(ctx.operationSeed, `teams:${teamIndex}`));
      const department = randomDepartment(teamSeed);
      const teamName = `${ctx.namingPrefix} - ${department} ${teamIndex + 1}`;
      const owner = availableUsers[teamIndex % availableUsers.length];
      await bumpSubcount(ctx.taskId, "team", { requested: 1 });
      if (!owner) {
        hadFailures = true;
        await bumpSubcount(ctx.taskId, "team", { failed: 1 });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
        continue;
      }

      try {
        const container = await getOrCreateContainer(ctx.taskId, "team", teamName, null, async () => {
          const team = await createTeam(ctx.client, teamName, `${ctx.namingPrefix.replace(/[^a-zA-Z0-9]/g, "")}${teamIndex}`, owner.id);
          return { graphId: team.groupId };
        });
        targets.push({ containerId: container.id, graphId: container.graphId, isNew: true });
        await bumpSubcount(ctx.taskId, "team", { created: 1 });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 1, failed: 0, skipped: 0, sizeBytes: 0 });

        const memberCount = Math.min(config.membersPerTeam, availableUsers.length);
        for (let m = 0; m < memberCount; m++) {
          const member = availableUsers[(teamIndex + m) % availableUsers.length]!;
          if (member.id === owner.id) continue;
          try {
            await addTeamMember(ctx.client, container.graphId, member.id);
          } catch {
            // A member already present (e.g. a resumed run) 409s — not a real failure; best-effort for a demo/test roster.
          }
        }
      } catch (err) {
        hadFailures = true;
        await bumpSubcount(ctx.taskId, "team", { failed: 1 });
        const message = err instanceof Error ? err.message : String(err);
        console.error(`[data-dump] Teams: team ${teamName} failed`, message);
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
      }
    }
  }

  // Explicit per-channel configs apply to every target team the same way; the legacy quick-start
  // path (no explicit selection/channels) synthesizes count-based channels per team instead.
  for (const target of targets) {
    if (await ctx.isCancelled()) return "cancelled";
    if (await ctx.isPaused()) return "paused";

    const channelPlans: ChannelConfig[] =
      config.channels && config.channels.length > 0
        ? config.channels
        : usingExplicitSelection
          ? []
          : channelNamesForTeam(randomDepartment(createSeededSource(forkSeed(ctx.operationSeed, `teams:${target.graphId}`))), config.channelsPerTeam).map((name) => ({
              name,
              messages: config.messagesPerChannel,
              replies: config.repliesPerMessage,
            }));

    for (const plan of channelPlans) {
      if (await ctx.isCancelled()) return "cancelled";
      if (await ctx.isPaused()) return "paused";

      await bumpSubcount(ctx.taskId, "channel", { requested: 1 });
      try {
        const isPrivate = !!plan.memberUserIds && plan.memberUserIds.length > 0;
        const channelContainer = await getOrCreateContainer(ctx.taskId, "channel", plan.name, target.containerId, async () => {
          if (plan.name === "General" && !isPrivate) return { graphId: "general" }; // every team already has a default General channel — not re-created
          const channel = await createChannel(ctx.client, target.graphId, plan.name, {
            membershipType: isPrivate ? "private" : "standard",
            ownerUserId: isPrivate ? plan.memberUserIds![0] : undefined,
          });
          return { graphId: channel.id };
        });
        await bumpSubcount(ctx.taskId, "channel", { created: 1 });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 1, failed: 0, skipped: 0, sizeBytes: 0 });

        if (isPrivate) {
          for (const memberId of plan.memberUserIds!.slice(1)) {
            await bumpSubcount(ctx.taskId, "channel_member", { requested: 1 });
            try {
              await addChannelMember(ctx.client, target.graphId, channelContainer.graphId, memberId);
              await bumpSubcount(ctx.taskId, "channel_member", { created: 1 });
            } catch {
              await bumpSubcount(ctx.taskId, "channel_member", { skipped: 1 });
            }
          }
        }

        await bumpSubcount(ctx.taskId, "message", { requested: plan.messages, skipped: plan.messages });
        await bumpSubcount(ctx.taskId, "reply", { requested: plan.messages * plan.replies, skipped: plan.messages * plan.replies });
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 0, skipped: plan.messages + plan.messages * plan.replies, sizeBytes: 0 });
      } catch (err) {
        hadFailures = true;
        await bumpSubcount(ctx.taskId, "channel", { failed: 1 });
        const message = err instanceof Error ? err.message : String(err);
        console.error(`[data-dump] Teams: channel ${plan.name} on team ${target.graphId} failed`, message);
        await applyTaskProgress(ctx.taskId, ctx.operationId, { created: 0, failed: 1, skipped: 0, sizeBytes: 0 });
      }
    }

    await setTaskCheckpoint(ctx.taskId, { phase: "teams", lastTeamGraphId: target.graphId });
  }

  return hadFailures ? "completed_with_errors" : "completed";
}
