"use client";
import { useEffect, useState } from "react";
import { useParams } from "next/navigation";

type Summary = {
  sheet_count: number;
  table_count: number;
  assumption_count: number;
  output_count: number;
  warning_count: number;
  warnings: { address: string; message: string }[];
};

export default function InspectPage() {
  const { jobId } = useParams<{ jobId: string }>();
  const [status, setStatus] = useState("loading");
  const [summary, setSummary] = useState<Summary | null>(null);

  useEffect(() => {
    const poll = setInterval(async () => {
      try {
        const res = await fetch(`/api/result/${jobId}`);
        const data = await res.json();
        setStatus(data.status);
        if (data.status === "done") {
          setSummary(data.summary);
          clearInterval(poll);
        }
        if (data.status === "error") clearInterval(poll);
      } catch {
        clearInterval(poll);
        setStatus("error");
      }
    }, 800);
    return () => clearInterval(poll);
  }, [jobId]);

  if (status !== "done") {
    return (
      <div className="min-h-screen flex items-center justify-center">
        <p className="text-[#737373] text-sm">
          {status === "error" ? "Conversion failed." : "Converting…"}
        </p>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen bg-[#171717]">
      {/* Nav */}
      <nav className="h-10 flex items-center justify-between px-4 border-b border-[#2e2e2e] shrink-0">
        <div className="flex items-center gap-3">
          <span className="text-[13px] font-medium text-[#ededed]">
            osheet<span className="text-[#525252] font-normal">.io</span>
          </span>
          <span className="text-[#525252] text-xs font-mono">/ converted</span>
        </div>
        <div className="flex gap-2">
          <a
            href={`/api/download/${jobId}/osheet`}
            className="text-xs font-medium text-[#ededed] border border-[#2e2e2e] px-3 py-1 rounded hover:bg-[#1f1f1f]"
          >
            Export .osheet
          </a>
          <a
            href={`/api/download/${jobId}/xlsx`}
            className="text-xs font-medium bg-[#ededed] text-[#171717] px-3 py-1 rounded hover:opacity-80"
          >
            ↓ Download .xlsx
          </a>
        </div>
      </nav>

      {/* Stat bar */}
      {summary && (
        <div className="flex items-center gap-2 px-4 py-2 border-b border-[#2e2e2e] shrink-0 bg-[#1f1f1f]">
          {[
            { label: "sheets", count: summary.sheet_count, color: "" },
            { label: "tables", count: summary.table_count, color: "" },
            { label: "assumptions", count: summary.assumption_count, color: "text-[#d4a96a]" },
            { label: "outputs", count: summary.output_count, color: "text-[#6ab07a]" },
            { label: "warnings", count: summary.warning_count, color: "text-[#d46a6a]" },
          ].map((s, i) => (
            <span
              key={i}
              className="flex items-center gap-1 text-xs border border-[#2e2e2e] bg-[#171717] px-2 py-1 rounded"
            >
              <span className={`font-mono font-medium ${s.color || "text-[#ededed]"}`}>{s.count}</span>
              <span className="text-[#737373]">{s.label}</span>
            </span>
          ))}
        </div>
      )}

      {/* Warnings */}
      {summary && summary.warnings.length > 0 && (
        <div className="px-4 py-2 border-b border-[#2e2e2e] bg-[#1f1f1f] shrink-0">
          {summary.warnings.map((w, i) => (
            <p key={i} className="text-xs text-[#d46a6a]">
              <span className="text-[#ededed] font-medium">{w.address}</span> — {w.message}
            </p>
          ))}
        </div>
      )}

      {/* Body */}
      <div className="flex-1 flex items-center justify-center text-[#737373] text-sm">
        <div className="text-center">
          <p className="mb-2 text-[#ededed]">Conversion complete.</p>
          <p className="text-xs text-[#525252]">Download your files above or upload another.</p>
          <a href="/" className="mt-4 inline-block text-xs text-[#525252] underline hover:text-[#737373]">
            ← Upload another file
          </a>
        </div>
      </div>
    </div>
  );
}
