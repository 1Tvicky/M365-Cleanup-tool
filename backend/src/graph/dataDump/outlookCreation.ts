import type { Client } from "@microsoft/microsoft-graph-client";

/**
 * Outlook Data Dump writes — uses the already-granted Mail.ReadWrite application permission
 * (docs/azure-ad-app-registration.md), the same permission the existing Cleaning module already
 * uses to enumerate/delete mail. No new Azure AD consent is required.
 *
 * CRITICAL SAFETY REQUIREMENT (spec §16): generated mail must never be sent externally. Every
 * message here is created DIRECTLY inside a mail folder via `POST /users/{id}/mailFolders/{folder}
 * /messages` — this creates a message resource in place (the same mechanism real mailbox-migration
 * tools use to inject historical mail) and never calls /sendMail or any endpoint that could dispatch
 * a real SMTP delivery. There is no code path in this file capable of sending mail.
 */

export interface CreatedMailItem {
  id: string;
}

export async function createMailItem(
  client: Client,
  userId: string,
  folder: "inbox" | "sentitems",
  message: {
    subject: string;
    body: string;
    fromUpn: string;
    toUpn: string;
    isRead?: boolean;
  }
): Promise<CreatedMailItem> {
  const created: any = await client.api(`/users/${userId}/mailFolders/${folder}/messages`).post({
    subject: message.subject,
    body: { contentType: "text", content: message.body },
    from: { emailAddress: { address: message.fromUpn } },
    toRecipients: [{ emailAddress: { address: message.toUpn } }],
    isRead: message.isRead ?? true,
    isDraft: false,
  });
  return { id: created.id };
}

/** Attaches a small file (base64 content) to an already-created message. Graph's fileAttachment endpoint accepts up to 3 MiB inline this way — larger attachments would need the same upload-session pattern as OneDrive/SharePoint, not implemented here since spec's Outlook attachment sizes are demo-scale, not large-file-migration-scale. */
export async function addMailAttachment(client: Client, userId: string, messageId: string, fileName: string, content: Buffer, mimeType: string): Promise<void> {
  await client.api(`/users/${userId}/messages/${messageId}/attachments`).post({
    "@odata.type": "#microsoft.graph.fileAttachment",
    name: fileName,
    contentType: mimeType,
    contentBytes: content.toString("base64"),
  });
}

export interface CreatedEvent {
  id: string;
}

/**
 * Creates a real calendar event with a genuine historical start/end (spec §21) — unlike Outlook
 * mail's receivedDateTime (only settable on the unsupported beta endpoint, see outlookRunner.ts's
 * comment), a calendar event's start/end are ordinary, client-supplied, v1.0-supported properties
 * regardless of whether the date is in the past — this is a real capability, not a workaround.
 * `attendeeUpns` are existing tenant users only (spec: never create external identities); an event
 * with no attendees is still a valid, real event on the organizer's own calendar.
 */
export async function createCalendarEvent(
  client: Client,
  userId: string,
  event: {
    subject: string;
    body: string;
    startIso: string;
    endIso: string;
    attendeeUpns?: string[];
    recurrence?: { type: "daily" | "weekly"; interval: number; endDateIso: string };
  }
): Promise<CreatedEvent> {
  const created: any = await client.api(`/users/${userId}/events`).post({
    subject: event.subject,
    body: { contentType: "text", content: event.body },
    start: { dateTime: event.startIso, timeZone: "UTC" },
    end: { dateTime: event.endIso, timeZone: "UTC" },
    isReminderOn: false,
    attendees: (event.attendeeUpns ?? []).map((upn) => ({ emailAddress: { address: upn }, type: "required" })),
    ...(event.recurrence
      ? {
          recurrence: {
            pattern: { type: event.recurrence.type, interval: event.recurrence.interval },
            range: { type: "endDate", startDate: event.startIso.slice(0, 10), endDate: event.recurrence.endDateIso.slice(0, 10) },
          },
        }
      : {}),
  });
  return { id: created.id };
}

export interface CreatedContact {
  id: string;
}

/** Creates a real contact directly in the mailbox's default Contacts folder — Contacts.ReadWrite is already granted (docs/azure-ad-app-registration.md, added for Outlook Contacts cleanup), no new consent needed. */
export async function createContact(
  client: Client,
  userId: string,
  contact: { givenName: string; surname: string; emailAddress: string; companyName?: string; jobTitle?: string; businessPhone?: string }
): Promise<CreatedContact> {
  const created: any = await client.api(`/users/${userId}/contacts`).post({
    givenName: contact.givenName,
    surname: contact.surname,
    emailAddresses: [{ address: contact.emailAddress, name: `${contact.givenName} ${contact.surname}` }],
    companyName: contact.companyName,
    jobTitle: contact.jobTitle,
    businessPhones: contact.businessPhone ? [contact.businessPhone] : [],
  });
  return { id: created.id };
}
