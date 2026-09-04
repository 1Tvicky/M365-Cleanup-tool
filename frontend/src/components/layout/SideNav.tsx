export type Page = "clouds" | "cleaning" | "reports";

const NAV_ITEMS: { id: Page; label: string; icon: string }[] = [
  { id: "clouds", label: "Clouds", icon: "\u{2601}\u{FE0F}" },
  { id: "cleaning", label: "Cleaning", icon: "\u{1F9F9}" },
  { id: "reports", label: "Reports", icon: "\u{1F4CB}" },
];

export function SideNav({ active, onNavigate, onLogout }: { active: Page; onNavigate: (page: Page) => void; onLogout: () => void }) {
  return (
    <aside className="flex h-screen w-20 shrink-0 flex-col bg-[#0b1c46]">
      <div className="flex flex-col items-center gap-1 py-5">
        <div className="flex h-8 w-8 items-center justify-center rounded-md bg-teal-400 text-xs font-bold text-[#0b1c46]">
          CF
        </div>
        <span className="text-[10px] font-medium text-slate-300">CloudFuze</span>
      </div>
      <nav className="mt-4 flex flex-col gap-1 px-2">
        {NAV_ITEMS.map((item) => {
          const isActive = item.id === active;
          return (
            <button
              key={item.id}
              onClick={() => onNavigate(item.id)}
              className={`relative flex flex-col items-center gap-1.5 rounded-md px-1 py-3 text-center transition-colors ${
                isActive ? "bg-white/10 text-white" : "text-slate-400 hover:bg-white/5 hover:text-slate-200"
              }`}
            >
              {isActive && <span className="absolute left-0 top-2 h-6 w-0.5 rounded-r bg-teal-400" />}
              <span className="text-lg" aria-hidden>
                {item.icon}
              </span>
              <span className="text-[11px] font-medium leading-tight">{item.label}</span>
            </button>
          );
        })}
      </nav>
      <div className="mt-auto px-2 pb-2">
        {/* CloudsPage also has its own UserMenu/Log out in its header — this one exists so logging
            out doesn't require navigating to Clouds first; it's reachable from every page. */}
        <button
          onClick={onLogout}
          className="flex w-full flex-col items-center gap-1.5 rounded-md px-1 py-3 text-center text-slate-400 transition-colors hover:bg-white/5 hover:text-slate-200"
        >
          <span className="text-lg" aria-hidden>
            🚪
          </span>
          <span className="text-[11px] font-medium leading-tight">Log out</span>
        </button>
      </div>
      <div className="space-y-1 px-3 pb-4 text-center text-[9px] leading-tight text-slate-500">
        <div>Terms of use</div>
        <div>Privacy policy</div>
      </div>
    </aside>
  );
}
