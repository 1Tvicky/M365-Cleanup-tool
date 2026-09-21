import { useEffect, useState } from "react";
import type { FileTypeKey, OneDriveGenConfig, OutlookGenConfig, SharePointGenConfig } from "../../api/dataDump";
import { useDebouncedValue } from "../cleaning/DiscoveryTable";

/**
 * Per-workload numeric configuration — folders/files/size for OneDrive & SharePoint (spec §4/§10),
 * emails/attachments/calendar/contacts for Outlook (spec §20/§21/§22). No field here carries a
 * hardcoded maximum (spec §23) — every NumberField accepts any non-negative value.
 *
 * Every field write goes through the PARENT's functional setState (never a plain-object merge built
 * from a locally-captured config snapshot) — several fields can debounce-commit within the same
 * tick, and a plain-object write here previously lost sibling fields under exactly that race
 * (fixed after catching it live against a real generation run — see git history on this file).
 */
export function OneDriveConfigureStep({ onChange }: { onChange: (updater: (prev: OneDriveGenConfig) => OneDriveGenConfig) => void }) {
  return (
    <div className="space-y-4">
      <ConfigGrid>
        <NumberField label="Total folders" hint="Root folders created directly under OneDrive." onChange={(n) => onChange((p) => ({ ...p, rootFolders: n }))} />
        <NumberField label="Sub folders" hint="Children under every folder, at every level." onChange={(n) => onChange((p) => ({ ...p, subFoldersPerFolder: n }))} />
        <NumberField label="Nested levels" hint="1 = root folders only, no sub folders." onChange={(n) => onChange((p) => ({ ...p, maxFolderDepth: n }))} />
        <NumberField label="Files per folder" hint="Created in every folder — root, sub, and nested." onChange={(n) => onChange((p) => ({ ...p, filesPerFolder: n }))} />
        <NumberField label="Target total size (MB)" onChange={(n) => onChange((p) => ({ ...p, targetTotalSizeBytes: n * 1024 * 1024 }))} />
        <NumberField label="Average file size (KB)" onChange={(n) => onChange((p) => ({ ...p, minFileSizeBytes: Math.max(1, n * 512), maxFileSizeBytes: Math.max(1, n * 1536) }))} />
      </ConfigGrid>
      <FileTypeField onChange={(distribution) => onChange((p) => setFileTypeDistribution(p, distribution))} />
    </div>
  );
}

export function SharePointConfigureStep({ onChange }: { onChange: (updater: (prev: SharePointGenConfig) => SharePointGenConfig) => void }) {
  return (
    <div className="space-y-4">
      <ConfigGrid>
        <NumberField label="Document libraries" onChange={(n) => onChange((p) => ({ ...p, librariesPerSite: n }))} />
        <NumberField label="Total folders" hint="Root folders created directly under each library." onChange={(n) => onChange((p) => ({ ...p, foldersPerLibrary: n }))} />
        <NumberField label="Sub folders" hint="Children under every folder, at every level." onChange={(n) => onChange((p) => ({ ...p, subFoldersPerFolder: n }))} />
        <NumberField label="Nested levels" hint="1 = root folders only, no sub folders." onChange={(n) => onChange((p) => ({ ...p, maxFolderDepth: n }))} />
        <NumberField label="Files per folder" hint="Created in every folder — root, sub, and nested." onChange={(n) => onChange((p) => ({ ...p, filesPerFolder: n }))} />
        <NumberField label="Target total size (MB)" onChange={(n) => onChange((p) => ({ ...p, targetTotalSizeBytes: n * 1024 * 1024 }))} />
      </ConfigGrid>
      <FileTypeField onChange={(distribution) => onChange((p) => setFileTypeDistribution(p, distribution))} />
    </div>
  );
}

/** Sets or clears `fileTypeDistribution` on a config object without ever sending `{}` — an empty distribution isn't "use the profile default," it's "every weight is zero," which pickWeighted treats as uniform over every FileTypeKey (see seededRandom.ts), silently reintroducing txt/zip/jpg the profile normally excludes. Omitting the key entirely is what actually falls back to the profile's own mix. */
function setFileTypeDistribution<T extends { fileTypeDistribution?: Partial<Record<FileTypeKey, number>> }>(prev: T, distribution: Partial<Record<FileTypeKey, number>> | undefined): T {
  const next = { ...prev };
  if (distribution) next.fileTypeDistribution = distribution;
  else delete next.fileTypeDistribution;
  return next;
}

const FILE_TYPE_OPTIONS: { key: FileTypeKey; label: string }[] = [
  { key: "docx", label: "Word (.docx)" },
  { key: "xlsx", label: "Excel (.xlsx)" },
  { key: "pptx", label: "PowerPoint (.pptx)" },
  { key: "pdf", label: "PDF (.pdf)" },
  { key: "txt", label: "Text (.txt)" },
  { key: "csv", label: "CSV (.csv)" },
  { key: "png", label: "Image (.png)" },
  { key: "zip", label: "Zip (.zip)" },
];

/**
 * Lets the operator restrict generation to specific file extensions instead of the profile's default
 * mixed distribution. Owns its own local selection (same uncontrolled-field convention as NumberField
 * below) — checking any box sends an equal-weight distribution over just the checked types; clearing
 * every box sends `undefined`, which reverts to the profile's default mix.
 */
function FileTypeField({ onChange }: { onChange: (distribution: Partial<Record<FileTypeKey, number>> | undefined) => void }) {
  const [selected, setSelected] = useState<Set<FileTypeKey>>(new Set());

  function toggle(key: FileTypeKey) {
    setSelected((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      if (next.size === 0) {
        onChange(undefined);
      } else {
        const distribution: Partial<Record<FileTypeKey, number>> = {};
        for (const k of next) distribution[k] = 1;
        onChange(distribution);
      }
      return next;
    });
  }

  return (
    <div>
      <label className="text-xs text-slate-500">File Types</label>
      <div className="mt-1.5 flex flex-wrap gap-x-4 gap-y-1.5">
        {FILE_TYPE_OPTIONS.map((opt) => (
          <label key={opt.key} className="flex items-center gap-1.5 text-xs text-slate-600">
            <input type="checkbox" checked={selected.has(opt.key)} onChange={() => toggle(opt.key)} />
            {opt.label}
          </label>
        ))}
      </div>
      <p className="mt-1 text-xs text-slate-400">
        {selected.size === 0 ? "No selection uses a mixed set of file types (profile default)." : "Only the checked file types will be generated."}
      </p>
    </div>
  );
}

export function OutlookConfigureStep({ onChange }: { onChange: (updater: (prev: OutlookGenConfig) => OutlookGenConfig) => void }) {
  return (
    <div className="space-y-5">
      <div>
        <h4 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Emails</h4>
        <ConfigGrid>
          <NumberField label="Emails per user" onChange={(n) => onChange((p) => ({ ...p, emailsPerUser: n }))} />
          <NumberField label="Attachments (%)" onChange={(n) => onChange((p) => ({ ...p, attachmentsPerEmail: Math.min(1, Math.max(0, n / 100)) }))} />
        </ConfigGrid>
      </div>
      <div>
        <h4 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Calendar</h4>
        <ConfigGrid>
          <NumberField label="Events per user" onChange={(n) => onChange((p) => ({ ...p, calendarEventCount: n }))} />
        </ConfigGrid>
        <div className="mt-2 flex gap-4 text-sm text-slate-600">
          <label className="flex items-center gap-1.5">
            <input type="checkbox" onChange={(e) => onChange((p) => ({ ...p, includeRecurringEvents: e.target.checked }))} />
            Include recurring events
          </label>
          <label className="flex items-center gap-1.5">
            <input type="checkbox" defaultChecked onChange={(e) => onChange((p) => ({ ...p, includeAttendees: e.target.checked }))} />
            Include attendees
          </label>
        </div>
      </div>
      <div>
        <h4 className="mb-2 text-xs font-semibold uppercase tracking-wide text-slate-500">Contacts</h4>
        <ConfigGrid>
          <NumberField label="Contacts per user" onChange={(n) => onChange((p) => ({ ...p, contactCount: n }))} />
        </ConfigGrid>
      </div>
    </div>
  );
}

function ConfigGrid({ children }: { children: React.ReactNode }) {
  return <div className="grid grid-cols-2 gap-4 sm:grid-cols-3">{children}</div>;
}

function NumberField({ label, hint, onChange }: { label: string; hint?: string; onChange: (n: number) => void }) {
  const [raw, setRaw] = useState("");
  const debounced = useDebouncedValue(raw, 300);
  useEffect(() => {
    if (debounced !== "") onChange(Number(debounced));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [debounced]);
  return (
    <label className="text-xs text-slate-500">
      {label}
      <input
        type="number"
        min={0}
        value={raw}
        onChange={(e) => setRaw(e.target.value)}
        placeholder="profile default"
        className="mt-1 block w-full rounded-md border border-slate-200 px-2 py-1.5 text-sm focus:border-[#1b2fc4] focus:outline-none"
      />
      {hint && <span className="mt-1 block text-[11px] font-normal text-slate-400">{hint}</span>}
    </label>
  );
}
