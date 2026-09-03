/**
 * Product icons for the M365 workloads this tool connects to. Shapes and exact hex values were
 * sampled pixel-by-pixel from the official icons (OneDrive #0364B8/#0078D4/#28A8EA cloud,
 * SharePoint #036C70/#1A9BA1/#37C6D0 circle cluster with an "S" tile, Teams' current boxless
 * flag+avatar mark) rather than approximated by eye.
 */

export function OneDriveIcon({ className }: { className?: string }) {
  const cloudPath =
    "M13 33c-4.7 0-8.5-3.6-8.5-8.1 0-4.1 3.1-7.5 7.2-8 1-4.6 5.1-8.1 10-8.1 5 0 9.2 3.6 10.1 8.3 4.1.5 7.2 3.9 7.2 8 0 4.3-3.6 7.9-8.1 7.9H13Z";
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden>
      <defs>
        <clipPath id="onedrive-cloud">
          <path d={cloudPath} />
        </clipPath>
      </defs>
      <path d={cloudPath} fill="#0364B8" />
      <path d="M6 22 40 10v26H6Z" fill="#28A8EA" opacity=".55" clipPath="url(#onedrive-cloud)" />
    </svg>
  );
}

export function SharePointIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden>
      <circle cx="30" cy="15" r="9" fill="#1A9BA1" />
      <circle cx="31" cy="29" r="8" fill="#37C6D0" />
      <rect x="6" y="12" width="22" height="22" rx="3" fill="#036C70" />
      <path
        d="M12 27.5c0 1.9 1.8 3.4 4.4 3.4 2.5 0 4.1-1.3 4.1-3.2 0-1.6-.9-2.5-2.9-3l-1.7-.4c-.9-.2-1.3-.5-1.3-1.1 0-.7.7-1.2 1.7-1.2 1.1 0 1.8.6 1.9 1.5h2.6c-.1-2.2-1.9-3.6-4.4-3.6-2.4 0-4.2 1.4-4.2 3.4 0 1.6 1 2.6 2.9 3.1l1.7.4c1 .3 1.4.6 1.4 1.2 0 .7-.7 1.2-1.9 1.2-1.3 0-2.1-.6-2.2-1.6H12Z"
        fill="#fff"
      />
    </svg>
  );
}

/** Current (boxless) Teams mark — a flag/"T" shape overlapping a rounded-square avatar, no bounding tile. */
export function TeamsIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden>
      <circle cx="36" cy="14" r="4.5" fill="#2D3894" />
      <path d="M30 15c0-2.2 1.8-4 4-4h7c2.2 0 4 1.8 4 4v8c0 3.3-2.7 6-6 6s-6-2.7-6-6v-8Z" fill="#6D81FF" />
      <path
        d="M8 12h19v4.2h-7.2V38h-4.6V16.2H8Z"
        fill="#2D3894"
      />
    </svg>
  );
}

/** Four-color M365 "waffle" badge — used to represent a whole connected tenant (all three workloads under one consent grant), not a single workload. */
export function Microsoft365Icon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden>
      <rect x="5" y="5" width="17" height="17" fill="#F25022" />
      <rect x="26" y="5" width="17" height="17" fill="#7FBA00" />
      <rect x="5" y="26" width="17" height="17" fill="#00A4EF" />
      <rect x="26" y="26" width="17" height="17" fill="#FFB900" />
    </svg>
  );
}

export function GoogleIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden>
      <path fill="#4285F4" d="M44 24.5c0-1.6-.1-3.1-.4-4.5H24v9h11.3c-.5 2.6-2 4.9-4.2 6.4v5.3h6.8C41.9 37 44 31.2 44 24.5Z" />
      <path fill="#34A853" d="M24 44c5.7 0 10.5-1.9 14-5.2l-6.8-5.3c-1.9 1.3-4.3 2-7.2 2-5.5 0-10.2-3.7-11.9-8.8H5.1v5.5C8.6 39.4 15.7 44 24 44Z" />
      <path fill="#FBBC05" d="M12.1 26.7c-.4-1.3-.7-2.7-.7-4.2s.2-2.9.7-4.2v-5.5H5.1C3.6 15.7 2.8 19.7 2.8 24s.8 8.3 2.3 11.2l7-5.5Z" />
      <path fill="#EA4335" d="M24 9.5c3.1 0 5.9 1.1 8.1 3.2l6-6C34.5 3.4 29.7 1.5 24 1.5 15.7 1.5 8.6 6.1 5.1 12.8l7 5.5C13.8 13.2 18.5 9.5 24 9.5Z" />
    </svg>
  );
}

/** Classic single-color Office "ribbon" mark — used on the Office 365 SSO button, distinct from the Microsoft365Icon waffle. */
export function OfficeRibbonIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden>
      <path
        d="M14 6h20l6 8-6 8H14l6-8Z"
        fill="#E8590C"
      />
      <path
        d="M14 22h20l6 8-6 8H14l6-8Z"
        fill="#E8590C"
        opacity=".85"
      />
    </svg>
  );
}

/* --- Outline variants for the login screen's blue-panel provider row (line-art on solid color, matching the reference footer treatment). --- */

export function OneDriveOutlineIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden fill="none" stroke="currentColor" strokeWidth="2" strokeLinejoin="round">
      <path d="M15 34c-4.7 0-8.5-3.6-8.5-8.1 0-4.1 3.1-7.5 7.2-8 1-4.6 5.1-8.1 10-8.1 5 0 9.2 3.6 10.1 8.3 4.1.5 7.2 3.9 7.2 8 0 4.3-3.6 7.9-8.1 7.9H15Z" />
    </svg>
  );
}

export function SharePointOutlineIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden fill="none" stroke="currentColor" strokeWidth="2">
      <circle cx="30" cy="15" r="8" />
      <circle cx="31" cy="29" r="7" />
      <rect x="6" y="12" width="21" height="21" rx="3" />
    </svg>
  );
}

export function TeamsOutlineIcon({ className }: { className?: string }) {
  return (
    <svg viewBox="0 0 48 48" className={className} aria-hidden fill="none" stroke="currentColor" strokeWidth="2" strokeLinejoin="round">
      <circle cx="36" cy="14" r="4" />
      <path d="M30 15c0-2.2 1.8-4 4-4h7c2.2 0 4 1.8 4 4v8c0 3.3-2.7 6-6 6s-6-2.7-6-6v-8Z" />
      <path d="M8 12h19v4.2h-7.2V38h-4.6V16.2H8Z" strokeLinecap="round" />
    </svg>
  );
}
