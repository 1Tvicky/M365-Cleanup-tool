import { useState } from "react";

/**
 * Google Workspace's "Connect" action — deliberately a form dialog, not a popup like the M365
 * connect flow. Domain-wide delegation is granted out-of-band by the customer's own Workspace
 * super-admin in their Admin Console (Security → API controls → Domain-wide delegation) before this
 * dialog can succeed; this just verifies that grant took effect. See
 * docs/google-workspace-integration.md for the full flow.
 */
export function GoogleConnectDialog({
  onCancel,
  onSubmit,
  submitting,
  error,
}: {
  onCancel: () => void;
  onSubmit: (domain: string, adminEmail: string) => void;
  submitting: boolean;
  error: string | null;
}) {
  const [domain, setDomain] = useState("");
  const [adminEmail, setAdminEmail] = useState("");

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 px-4" onClick={onCancel}>
      <div className="w-full max-w-md rounded-xl bg-white p-6 shadow-xl" onClick={(e) => e.stopPropagation()}>
        <h3 className="text-base font-semibold text-slate-900">Connect Google Workspace</h3>
        <p className="mt-2 text-sm text-slate-500">
          Before continuing, a Workspace super admin must authorize this app's service account for domain-wide delegation in{" "}
          <span className="font-medium">Admin Console → Security → API controls → Domain-wide delegation</span>, with the required scopes. This
          step can't be completed from here.
        </p>

        <label className="mt-4 block text-sm font-medium text-slate-700">
          Workspace domain
          <input
            value={domain}
            onChange={(e) => setDomain(e.target.value)}
            placeholder="acme.com"
            className="mt-1 block w-full rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        </label>

        <label className="mt-3 block text-sm font-medium text-slate-700">
          Admin email
          <input
            value={adminEmail}
            onChange={(e) => setAdminEmail(e.target.value)}
            placeholder="admin@acme.com"
            className="mt-1 block w-full rounded-md border border-slate-200 px-3 py-2 text-sm focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500"
          />
        </label>

        {error && <p className="mt-3 text-sm text-rose-600">{error}</p>}

        <div className="mt-5 flex justify-end gap-3">
          <button onClick={onCancel} className="rounded-md border border-slate-200 px-4 py-2 text-sm font-semibold text-slate-600 hover:bg-slate-50">
            Cancel
          </button>
          <button
            onClick={() => onSubmit(domain.trim(), adminEmail.trim())}
            disabled={submitting || !domain.trim() || !adminEmail.trim()}
            className="rounded-md bg-[#1a73e8] px-4 py-2 text-sm font-semibold text-white hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50"
          >
            {submitting ? "Verifying…" : "Verify & Connect"}
          </button>
        </div>
      </div>
    </div>
  );
}
