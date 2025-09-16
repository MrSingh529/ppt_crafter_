"use client";
import React, { useState } from "react";

export default function UploadCard() {
  const [excel, setExcel] = useState<File | null>(null);
  const [ppt, setPpt] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const onSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setError(null);
    if (!excel) {
      setError("Please select the Excel file.");
      return;
    }
    const form = new FormData();
    form.append("excel", excel);
    if (ppt) form.append("template", ppt); // optional

    setLoading(true);
    try {
      const res = await fetch("/api", { method: "POST", body: form });
      if (!res.ok) {
        const msg = await res.text();
        throw new Error(msg || "Server error");
      }
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = "updated_poc.pptx";
      document.body.appendChild(a); a.click(); a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err: any) {
      setError(err.message || "Something went wrong");
    } finally {
      setLoading(false);
    }
  };

  return (
    <form onSubmit={onSubmit} className="bg-white shadow-soft rounded-2xl p-6">
      <div className="grid gap-4">
        <div>
          <label className="block text-sm font-medium mb-1">Excel file</label>
          <input
            type="file"
            accept=".xls,.xlsx"
            onChange={(e) => setExcel(e.target.files?.[0] || null)}
            className="w-full file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:bg-neutral-900 file:text-white file:text-sm file:cursor-pointer border rounded-xl p-1"
          />
        </div>
        <div>
          <label className="block text-sm font-medium mb-1">PPT template <span className="text-neutral-500">(optional)</span></label>
          <input
            type="file"
            accept=".pptx"
            onChange={(e) => setPpt(e.target.files?.[0] || null)}
            className="w-full file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:bg-neutral-900 file:text-white file:text-sm file:cursor-pointer border rounded-xl p-1"
          />
        </div>
      </div>

      {error && <p className="text-sm text-red-600 mt-3">{error}</p>}

      <button
        type="submit"
        className="mt-5 w-full rounded-2xl px-4 py-3 text-sm font-medium bg-neutral-900 text-white disabled:opacity-60"
        disabled={loading}
      >
        {loading ? "Generating…" : "Generate PPTX"}
      </button>

      <p className="mt-3 text-xs text-neutral-500">
        Upload Excel only to use our default template, or include your own .pptx.
      </p>
    </form>
  );
}
