interface Props {
  file: File;
  onRemove: () => void;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / (1024 * 1024)).toFixed(1) + " MB";
}

export default function FileCard({ file, onRemove }: Props) {
  return (
    <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 rounded-xl p-3.5 px-4 text-left">
      <div className="shrink-0">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          fill="none"
          viewBox="0 0 24 24"
          strokeWidth={1.5}
          stroke="currentColor"
          className="w-8 h-8 text-blue-600"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M3.75 3v11.25A2.25 2.25 0 0 0 6 16.5h2.25M3.75 3h-1.5m1.5 0h16.5m0 0h1.5m-1.5 0v11.25A2.25 2.25 0 0 1 18 16.5h-2.25m-7.5 0h7.5m-7.5 0-1 3m8.5-3 1 3m0 0 .5 1.5m-.5-1.5h-9.5m0 0-.5 1.5M9 11.25v1.5M12 9v3.75m3-6v6"
          />
        </svg>
      </div>
      <div className="flex-1 min-w-0">
        <div className="text-sm font-semibold text-slate-900 truncate">
          {file.name}
        </div>
        <div className="text-xs text-slate-400">{formatSize(file.size)}</div>
      </div>
      <button
        onClick={onRemove}
        className="shrink-0 bg-transparent border-none text-slate-400 cursor-pointer p-1 rounded hover:text-red-500 transition-colors"
        title="Remove file"
      >
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="20"
          height="20"
          fill="none"
          viewBox="0 0 24 24"
          strokeWidth={1.5}
          stroke="currentColor"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M6 18 18 6M6 6l12 12"
          />
        </svg>
      </button>
    </div>
  );
}
