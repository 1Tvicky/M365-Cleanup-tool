import { useState } from "react";
import type { AvailableResourceRow } from "../../api/clouds";
import { listTenantUsers, type ChannelConfig, type NewTeamConfig } from "../../api/dataDump";
import { ResourceSelectionStep, connectionResourceFetcher } from "./ResourceSelectionStep";

/**
 * Teams' Select Resources step — existing teams (spec §13) and/or a brand-new team (spec §14/§15),
 * with explicit tenant-user membership and per-channel configuration (spec §16/§17), all in the same
 * Cleanup-style resource-selection shape as the other workloads.
 */
export function TeamsStep({
  tenantId,
  teamsConnectionId,
  selectedTeams,
  onToggleTeam,
  onToggleAllTeams,
  newTeam,
  onNewTeamChange,
  channels,
  onChannelsChange,
}: {
  tenantId: string;
  teamsConnectionId: string | null;
  selectedTeams: Map<string, AvailableResourceRow>;
  onToggleTeam: (id: string, row: AvailableResourceRow) => void;
  onToggleAllTeams: (rows: AvailableResourceRow[]) => void;
  newTeam: NewTeamConfig | null;
  onNewTeamChange: (team: NewTeamConfig | null) => void;
  channels: ChannelConfig[];
  onChannelsChange: (channels: ChannelConfig[]) => void;
}) {
  const [showNewTeamForm, setShowNewTeamForm] = useState(!!newTeam);
  const [memberDisplayNames, setMemberDisplayNames] = useState<Map<string, string>>(new Map());

  const memberFetcher = (opts: { search?: string; page?: number; pageSize?: number }) => listTenantUsers(tenantId, opts);

  function toggleMember(id: string, row: AvailableResourceRow) {
    const current = newTeam ?? emptyTeam();
    const next = new Set(current.memberUserIds);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setMemberDisplayNames((prev) => new Map(prev).set(id, row.displayName));
    onNewTeamChange({ ...current, memberUserIds: [...next] });
  }

  return (
    <div className="space-y-6">
      {teamsConnectionId ? (
        <ResourceSelectionStep
          fetcher={connectionResourceFetcher(teamsConnectionId)}
          title="Teams"
          subtitle="Select existing Teams to generate channels/members into."
          searchPlaceholder="Search Teams…"
          secondaryLabel="Details"
          selected={selectedTeams}
          onToggle={onToggleTeam}
          onToggleAll={onToggleAllTeams}
          emptyMessage="No Teams found."
        />
      ) : (
        <p className="rounded-xl border border-slate-200 bg-white px-4 py-6 text-center text-sm text-slate-500">
          Connect Microsoft Teams via Add Clouds to select existing Teams, or create a new one below.
        </p>
      )}

      <div className="rounded-xl border border-slate-200 bg-white p-5">
        <label className="flex cursor-pointer items-center gap-2 text-sm font-medium text-slate-700">
          <input
            type="checkbox"
            checked={showNewTeamForm}
            onChange={(e) => {
              setShowNewTeamForm(e.target.checked);
              if (!e.target.checked) onNewTeamChange(null);
            }}
          />
          Create New Team
        </label>
        {showNewTeamForm && (
          <div className="mt-4 space-y-4">
            <Field label="Team Name">
              <input
                value={newTeam?.displayName ?? ""}
                onChange={(e) => onNewTeamChange({ ...emptyTeam(newTeam), displayName: e.target.value })}
                placeholder="Migration Demo Team"
                className="w-full rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-[#1b2fc4] focus:outline-none"
              />
            </Field>
            <Field label="Description">
              <input
                value={newTeam?.description ?? ""}
                onChange={(e) => onNewTeamChange({ ...emptyTeam(newTeam), description: e.target.value })}
                placeholder="Synthetic migration test data"
                className="w-full rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-[#1b2fc4] focus:outline-none"
              />
            </Field>
            <Field label="Privacy">
              <div className="space-y-1.5">
                <label className="flex items-center gap-2 text-sm text-slate-700">
                  <input type="radio" checked={(newTeam?.visibility ?? "private") === "private"} onChange={() => onNewTeamChange({ ...emptyTeam(newTeam), visibility: "private" })} />
                  Private
                </label>
                <label className="flex items-center gap-2 text-sm text-slate-700">
                  <input type="radio" checked={newTeam?.visibility === "public"} onChange={() => onNewTeamChange({ ...emptyTeam(newTeam), visibility: "public" })} />
                  Public
                </label>
              </div>
            </Field>
            <Field label={`Members (${newTeam?.memberUserIds.length ?? 0} selected)`}>
              <ResourceSelectionStep
                fetcher={memberFetcher}
                title="Select Users"
                subtitle=""
                searchPlaceholder="Search users…"
                secondaryLabel="Email"
                selected={new Map((newTeam?.memberUserIds ?? []).map((id) => [id, { id, displayName: memberDisplayNames.get(id) ?? id }]))}
                onToggle={toggleMember}
                onToggleAll={(rows) => rows.forEach((r) => toggleMember(r.id, r))}
                emptyMessage="No users found."
              />
            </Field>
          </div>
        )}
      </div>

      <ChannelsEditor channels={channels} onChange={onChannelsChange} candidateMembers={memberDisplayNames} />
    </div>
  );
}

function emptyTeam(current?: NewTeamConfig | null): NewTeamConfig {
  return current ?? { displayName: "", visibility: "private", memberUserIds: [] };
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div className="text-xs font-medium text-slate-500">
      {label}
      <div className="mt-1">{children}</div>
    </div>
  );
}

/** Dynamic list of per-channel configs — name/members/messages/replies (spec §16), add/remove rows freely, no fixed maximum. */
function ChannelsEditor({
  channels,
  onChange,
  candidateMembers,
}: {
  channels: ChannelConfig[];
  onChange: (channels: ChannelConfig[]) => void;
  candidateMembers: Map<string, string>;
}) {
  function updateChannel(index: number, patch: Partial<ChannelConfig>) {
    const next = [...channels];
    next[index] = { ...next[index]!, ...patch };
    onChange(next);
  }

  function toggleChannelMember(index: number, userId: string) {
    const current = new Set(channels[index]!.memberUserIds ?? []);
    if (current.has(userId)) current.delete(userId);
    else current.add(userId);
    updateChannel(index, { memberUserIds: [...current] });
  }

  return (
    <div className="rounded-xl border border-slate-200 bg-white p-5">
      <div className="mb-3 flex items-center justify-between">
        <h3 className="text-sm font-semibold text-slate-800">Channels</h3>
        <button
          onClick={() => onChange([...channels, { name: `Channel ${channels.length + 1}`, messages: 0, replies: 0 }])}
          className="rounded-md border border-slate-300 px-3 py-1.5 text-xs font-medium text-slate-700 hover:bg-slate-50"
        >
          + Add Channel
        </button>
      </div>
      {channels.length === 0 && <p className="text-sm text-slate-400">No channels configured yet — click "Add Channel".</p>}
      <div className="space-y-4">
        {channels.map((channel, i) => (
          <div key={i} className="rounded-lg border border-slate-100 bg-slate-50 p-4">
            <div className="mb-3 flex items-center justify-between">
              <input
                value={channel.name}
                onChange={(e) => updateChannel(i, { name: e.target.value })}
                className="rounded-md border border-slate-200 px-3 py-1.5 text-sm font-medium focus:border-[#1b2fc4] focus:outline-none"
              />
              <button onClick={() => onChange(channels.filter((_, j) => j !== i))} className="text-xs text-rose-600 hover:underline">
                Remove
              </button>
            </div>
            {candidateMembers.size > 0 && (
              <div className="mb-3">
                <div className="mb-1 text-xs text-slate-500">Members (leave empty = every team member)</div>
                <div className="flex flex-wrap gap-2">
                  {[...candidateMembers.entries()].map(([id, name]) => (
                    <label key={id} className="flex items-center gap-1 rounded-full border border-slate-200 bg-white px-2 py-1 text-xs text-slate-600">
                      <input type="checkbox" checked={(channel.memberUserIds ?? []).includes(id)} onChange={() => toggleChannelMember(i, id)} />
                      {name}
                    </label>
                  ))}
                </div>
              </div>
            )}
            <div className="grid grid-cols-2 gap-3">
              <label className="text-xs text-slate-500">
                Messages
                <input
                  type="number"
                  min={0}
                  value={channel.messages}
                  onChange={(e) => updateChannel(i, { messages: Number(e.target.value) || 0 })}
                  className="mt-1 block w-full rounded-md border border-slate-200 px-2 py-1 text-sm focus:border-[#1b2fc4] focus:outline-none"
                />
              </label>
              <label className="text-xs text-slate-500">
                Replies (per message)
                <input
                  type="number"
                  min={0}
                  value={channel.replies}
                  onChange={(e) => updateChannel(i, { replies: Number(e.target.value) || 0 })}
                  className="mt-1 block w-full rounded-md border border-slate-200 px-2 py-1 text-sm focus:border-[#1b2fc4] focus:outline-none"
                />
              </label>
            </div>
          </div>
        ))}
      </div>
      {channels.some((c) => c.messages > 0) && (
        <p className="mt-3 text-xs text-amber-700">
          ⚠ Messages/replies are configured for planning purposes only — Microsoft Graph doesn't support posting channel messages under application permissions outside a migration-mode
          team import. Channels themselves ARE created for real.
        </p>
      )}
    </div>
  );
}
