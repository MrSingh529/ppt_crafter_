import UploadCard from "../components/UploadCard";

export default function Home() {
  return (
    <main className="min-h-screen flex items-center justify-center p-6">
      <div className="max-w-2xl w-full">
        <div className="mb-8 text-center">
          <h1 className="text-3xl font-semibold tracking-tight">IMARC POC - PPT Crafter</h1>
          <p className="text-sm text-neutral-600 mt-2">Upload your Excel & PPT template. We’ll generate a polished PPTX using Python logic.</p>
        </div>
        <UploadCard />
        <footer className="mt-8 text-center text-xs text-neutral-500">POC • Built by Harpinder Singh</footer>
      </div>
    </main>
  );
}