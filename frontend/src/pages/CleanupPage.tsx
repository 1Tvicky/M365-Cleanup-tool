import { useEffect, useRef, useState } from "react";
import { TenantSelector } from "../components/cleanup/TenantSelector";
import { ScopePicker, type CleanupScopeForm } from "../components/cleanup/ScopePicker";
import { PreviewTable } from "../components/cleanup/PreviewTable";
import { ConfirmModal } from "../components/cleanup/ConfirmModal";
import { ProgressTracker } from "../components/cleanup/ProgressTracker";
import { MOCK_TENANTS, MOCK_PREVIEW } from "../api/mockData";
import type { JobStatus, Preview } from "../types";

const EMPTY_FORM: CleanupScopeForm = {
  workloads: new Set(),
  cutoffDate: "",
  removeM365Groups: false,
  searchQuery: "",
};

export function CleanupPage() {
  const [tenantId, setTenantId] = useState<string | null>(MOCK_TENANTS[0]?.id ?? null);
  const [form, setForm] = useState<CleanupScopeForm>(EMPTY_FORM);
  const [preview, setPreview] = useState<Preview | null>(null);
  const [showConfirm, setShowConfirm] = useState(false);
  const [job, setJob] = useState<{ status: JobStatus; progress: { total: number; completed: number; failed: number } } | null>(
    null
  );
  const intervalRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const tenant = MOCK_TENANTS.find((t) => t.id === tenantId) ?? null;

  function handlePreview() {
    // In production: POST /api/v1/tenants/:tenantId/cleanup/preview — dry-run only, no deletes.
    setPreview(MOCK_PREVIEW);
  }

  function handleExecute(_exportManifestOnly: boolean) {
    setShowConfirm(false);
    setJob({ status: "export_in_progress", progress: { total: MOCK_PREVIEW.totals.itemCount, completed: 0, failed: 0 } });
  }

  function handleCancel() {
    setJob((prev) => (prev ? { ...prev, status: "cancelled" } : prev));
  }

  // Simulates job progress client-side so the flow is fully demonstrable without a live backend.
  useEffect(() => {
    if (!job || job.status === "cancelled") return;

    if (job.status === "export_in_progress") {
      const t = setTimeout(() => setJob((prev) => (prev ? { ...prev, status: "queued" } : prev)), 1200);
      return () => clearTimeout(t);
    }
    if (job.status === "queued") {
      const t = setTimeout(() => setJob((prev) => (prev ? { ...prev, status: "running" } : prev)), 800);
      return () => clearTimeout(t);
    }
    if (job.status === "running") {
      intervalRef.current = setInterval(() => {
        setJob((prev) => {
          if (!prev) return prev;
          const next = Math.min(prev.progress.completed + Math.ceil(prev.progress.total / 20), prev.progress.total);
          const done = next >= prev.progress.total;
          return {
            status: done ? "completed" : "running",
            progress: { ...prev.progress, completed: next },
          };
        });
      }, 400);
      return () => {
        if (intervalRef.current) clearInterval(intervalRef.current);
      };
    }
  }, [job?.status]);

  return (
    <div className="mx-auto max-w-4xl px-8 py-8">
      <h1 className="mb-6 text-xl font-semibold text-slate-900">Cleanup</h1>

      <TenantSelector tenants={MOCK_TENANTS} selectedId={tenantId} onSelect={setTenantId} />

      {tenant && (
        <ScopePicker form={form} onChange={setForm} onPreview={handlePreview} />
      )}

      {preview && !job && <PreviewTable preview={preview} onConfirm={() => setShowConfirm(true)} />}

      {job && (
        <ProgressTracker status={job.status} progress={job.progress} onCancel={handleCancel} />
      )}

      {showConfirm && preview && tenant && (
        <ConfirmModal
          preview={preview}
          tenantName={tenant.displayName}
          onClose={() => setShowConfirm(false)}
          onExecute={handleExecute}
        />
      )}
    </div>
  );
}
