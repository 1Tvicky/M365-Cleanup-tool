import { useMemo, useState } from "react";
import type { CleaningChannelRow, CountStatus } from "../../api/cleaning";

/** Never a bare "0" when the real count isn't known yet — always one of these friendly phrases instead. */
function messageCountLabel(row: CleaningChannelRow): string {
  if (row.countStatus === "completed" && row.messageCount !== null) return `${row.messageCount.toLocaleString()} messages`;
  if (row.countStatus === "calculating") return "Calculating message count…";
  if (row.countStatus === "failed") return "Unable to calculate";
  return "Waiting to be calculated";
}

function statusColor(status: CountStatus): string {
  if (status === "completed") return "text-slate-700";
  if (status === "failed") return "text-rose-500";
  return "text-slate-400 italic";
}

/** Rolls up a team's channel statuses without conflating "still working on it" with "gave up" — the doc is explicit these must read differently, never both as a generic "Calculating…". */
function teamMessageLabel(channels: CleaningChannelRow[]): { text: string; className: string } {
  if (channels.every((c) => c.countStatus === "completed")) {
    const total = channels.reduce((sum, c) => sum + (c.messageCount ?? 0), 0);
    return { text: `${total.toLocaleString()} messages`, className: "text-slate-600" };
  }
  if (channels.some((c) => c.countStatus === "calculating")) {
    return { text: "Calculating…", className: "text-slate-400 italic" };
  }
  if (channels.every((c) => c.countStatus === "failed")) {
    return { text: "Unable to calculate", className: "text-rose-500" };
  }
  if (channels.some((c) => c.countStatus === "failed")) {
    return { text: "Partially unable to calculate", className: "text-amber-600" };
  }
  return { text: "Waiting to be calculated", className: "text-slate-400 italic" };
}

export function TeamsChannels({
  channels,
  loading,
  error,
  search,
  onSearchChange,
  selected,
  onToggle,
  onToggleTeam,
}: {
  channels: CleaningChannelRow[];
  loading: boolean;
  error: string | null;
  search: string;
  onSearchChange: (value: string) => void;
  selected: Set<string>;
  onToggle: (channelId: string) => void;
  onToggleTeam: (channelIds: string[]) => void;
}) {
  const [expanded, setExpanded] = useState<Set<string>>(new Set());

  const teams = useMemo(() => {
    const byTeam = new Map<string, { teamName: string; channels: CleaningChannelRow[] }>();
    for (const ch of channels) {
      if (!byTeam.has(ch.teamId)) byTeam.set(ch.teamId, { teamName: ch.teamName, channels: [] });
      byTeam.get(ch.teamId)!.channels.push(ch);
    }
    return [...byTeam.entries()].map(([teamId, v]) => ({ teamId, ...v }));
  }, [channels]);

  function toggleExpand(teamId: string) {
    setExpanded((prev) => {
      const next = new Set(prev);
      next.has(teamId) ? next.delete(teamId) : next.add(teamId);
      return next;
    });
  }

  return (
    <div className="overflow-hidden rounded-xl border border-slate-200 bg-white">
      <div className="border-b border-slate-100 px-4 py-3">
        <input
          type="search"
          value={search}
          onChange={(e) => onSearchChange(e.target.value)}
          placeholder="Search teams or channels…"
          className="w-64 rounded-md border border-slate-200 px-3 py-1.5 text-sm text-slate-700 placeholder:text-slate-400 focus:border-[#1b2fc4] focus:outline-none"
        />
      </div>

      {loading ? (
        <p className="px-4 py-8 text-center text-sm text-slate-500">Discovering teams…</p>
      ) : error ? (
        <p className="px-4 py-8 text-center text-sm text-rose-600">{error}</p>
      ) : teams.length === 0 ? (
        <p className="px-4 py-8 text-center text-sm text-slate-500">No teams found yet.</p>
      ) : (
        <div className="divide-y divide-slate-100">
          {teams.map((team) => {
            const isExpanded = expanded.has(team.teamId);
            const teamStatus = teamMessageLabel(team.channels);
            const channelIds = team.channels.map((c) => c.id);
            const allSelected = channelIds.every((id) => selected.has(id));

            return (
              <div key={team.teamId}>
                <div className="flex items-center gap-3 px-4 py-3">
                  <input type="checkbox" checked={allSelected} onChange={() => onToggleTeam(channelIds)} aria-label={`Select all channels in ${team.teamName}`} />
                  <button onClick={() => toggleExpand(team.teamId)} className="flex flex-1 items-center gap-2 text-left">
                    <span className={`inline-block text-slate-400 transition-transform ${isExpanded ? "rotate-90" : ""}`}>▶</span>
                    <span className="text-sm font-semibold text-slate-800">{team.teamName}</span>
                    <span className="text-xs text-slate-400">{team.channels.length} channel{team.channels.length === 1 ? "" : "s"}</span>
                  </button>
                  <span className={`text-sm ${teamStatus.className}`}>{teamStatus.text}</span>
                </div>
                {isExpanded && (
                  <div className="bg-slate-50/60 pb-2 pl-11 pr-4">
                    {team.channels.map((ch) => (
                      <div key={ch.id} className="flex items-center gap-3 py-1.5">
                        <input type="checkbox" checked={selected.has(ch.id)} onChange={() => onToggle(ch.id)} aria-label={`Select ${ch.channelName}`} />
                        <span className="flex-1 text-sm text-slate-700">{ch.channelName}</span>
                        <span className={`text-xs ${statusColor(ch.countStatus)}`}>{messageCountLabel(ch)}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
