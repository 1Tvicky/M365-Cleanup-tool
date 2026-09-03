import { useState } from "react";

export function UserMenu({ name, onLogout }: { name: string; onLogout?: () => void }) {
  const [open, setOpen] = useState(false);

  return (
    <div className="relative">
      <button
        onClick={() => setOpen((v) => !v)}
        className="flex items-center gap-2 rounded-md px-2 py-1.5 text-sm text-slate-700 hover:bg-slate-50"
      >
        <span className="flex h-7 w-7 items-center justify-center rounded-full border border-brand-300 text-brand-600">
          <svg viewBox="0 0 24 24" className="h-4 w-4" fill="currentColor">
            <path d="M12 12a4.5 4.5 0 1 0 0-9 4.5 4.5 0 0 0 0 9Zm0 2c-4.14 0-7.5 2.35-7.5 5.25V21h15v-1.75C19.5 16.35 16.14 14 12 14Z" />
          </svg>
        </span>
        <span className="font-medium">{name}</span>
        <svg viewBox="0 0 20 20" className="h-3.5 w-3.5 text-slate-400" fill="currentColor">
          <path d="M5.5 7.5 10 12l4.5-4.5" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" />
        </svg>
      </button>

      {open && onLogout && (
        <>
          <div className="fixed inset-0 z-10" onClick={() => setOpen(false)} />
          <div className="absolute right-0 z-20 mt-1 w-40 rounded-lg border border-slate-200 bg-white py-1 shadow-lg">
            <button
              onClick={() => {
                setOpen(false);
                onLogout();
              }}
              className="w-full px-3 py-2 text-left text-sm text-slate-700 hover:bg-slate-50"
            >
              Log out
            </button>
          </div>
        </>
      )}
    </div>
  );
}
