"use client";
import { useState, useCallback } from "react";
import { useRouter } from "next/navigation";

export default function UploadPage() {
  const router = useRouter();
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith(".xlsx")) {
      setError("Only .xlsx files are supported");
      return;
    }
    setLoading(true);
    setError(null);
    const form = new FormData();
    form.append("file", file);
    try {
      const res = await fetch("/api/convert", { method: "POST", body: form });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "Conversion failed");
      router.push(`/inspect/${data.job_id}`);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : "Conversion failed");
      setLoading(false);
    }
  }, [router]);

  return (
    <main className="min-h-screen flex flex-col items-center justify-center px-6">
      <div className="mb-10 text-center">
        <h1 className="text-[32px] font-medium tracking-[-0.8px] text-[#ededed] mb-3">osheet</h1>
        <p className="text-[#737373] text-[15px]">Upload any .xlsx — get back an AI-native workbook.</p>
      </div>

      <div
        className={`w-full max-w-md border-2 border-dashed rounded-lg p-12 text-center cursor-pointer transition-colors ${dragging ? "border-[#ededed] bg-[#1f1f1f]" : "border-[#2e2e2e] hover:border-[#525252]"}`}
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={(e) => { e.preventDefault(); setDragging(false); const f = e.dataTransfer.files[0]; if (f) handleFile(f); }}
        onClick={() => document.getElementById("file-input")?.click()}
      >
        <input id="file-input" type="file" accept=".xlsx" className="hidden"
          onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
        {loading ? (
          <p className="text-[#b8b8b8] text-sm">Converting…</p>
        ) : (
          <>
            <p className="text-[#ededed] text-sm font-medium mb-1">Drop your .xlsx here</p>
            <p className="text-[#525252] text-xs">or click to browse · max 20 MB</p>
          </>
        )}
      </div>

      {error && <p className="mt-4 text-[#d46a6a] text-sm">{error}</p>}
      <p className="mt-8 text-[#525252] text-xs">Files are processed in memory and not stored.</p>
    </main>
  );
}
