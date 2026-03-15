const formats = [
  { label: "Microsoft PowerPoint (.pptx)", icon: "doc" },
  { label: "Adobe PDF (.pdf)", icon: "download" },
  { label: "Legacy PowerPoint (.ppt)", icon: "doc" },
  { label: "Microsoft Word (.docx)", icon: "doc" },
  { label: "Legacy Word (.doc)", icon: "doc" },
] as const;

function FormatIcon({ type }: { type: "doc" | "download" }) {
  if (type === "download") {
    return (
      <svg
        xmlns="http://www.w3.org/2000/svg"
        fill="none"
        viewBox="0 0 24 24"
        strokeWidth={1.5}
        stroke="currentColor"
        className="w-4 h-4 text-blue-600 shrink-0"
      >
        <path
          strokeLinecap="round"
          strokeLinejoin="round"
          d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m.75 12 3 3m0 0 3-3m-3 3v-6m-1.5-9H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"
        />
      </svg>
    );
  }
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      fill="none"
      viewBox="0 0 24 24"
      strokeWidth={1.5}
      stroke="currentColor"
      className="w-4 h-4 text-blue-600 shrink-0"
    >
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M19.5 14.25v-2.625a3.375 3.375 0 0 0-3.375-3.375h-1.5A1.125 1.125 0 0 1 13.5 7.125v-1.5a3.375 3.375 0 0 0-3.375-3.375H8.25m0 12.75h7.5m-7.5 3H12M10.5 2.25H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 0 0-9-9Z"
      />
    </svg>
  );
}

export default function SupportedFormats() {
  return (
    <div className="mb-4">
      <h3 className="text-base font-bold text-slate-900 mb-3">
        Supported Formats
      </h3>
      <div className="flex gap-2.5 flex-wrap">
        {formats.map((f) => (
          <div
            key={f.label}
            className="flex items-center gap-2 px-4 py-2 bg-slate-100 border border-slate-200 rounded-lg text-sm font-medium text-slate-600"
          >
            <FormatIcon type={f.icon} />
            {f.label}
          </div>
        ))}
      </div>
    </div>
  );
}
