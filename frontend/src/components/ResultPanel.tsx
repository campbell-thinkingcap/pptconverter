interface Props {
  filename: string;
  onDownloadAgain: () => void;
  onConvertAnother: () => void;
}

export default function ResultPanel({
  filename,
  onDownloadAgain,
  onConvertAnother,
}: Props) {
  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-8 text-center mb-8">
      <div className="w-14 h-14 bg-green-50 rounded-full flex items-center justify-center mx-auto mb-4">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          fill="none"
          viewBox="0 0 24 24"
          strokeWidth={2}
          stroke="currentColor"
          className="w-7 h-7 text-green-500"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="m4.5 12.75 6 6 9-13.5"
          />
        </svg>
      </div>
      <div className="text-lg font-bold text-green-500 mb-1.5">
        SCORM Package Ready!
      </div>
      <div className="text-sm text-slate-500 mb-5 break-all">{filename}</div>
      <div className="flex gap-2.5 justify-center">
        <button
          onClick={onDownloadAgain}
          className="px-6 py-2.5 bg-green-500 text-white border-none rounded-xl text-sm font-bold cursor-pointer hover:bg-green-600 transition-colors"
        >
          Download Again
        </button>
        <button
          onClick={onConvertAnother}
          className="px-6 py-2.5 bg-white text-slate-500 border border-slate-300 rounded-xl text-sm font-semibold cursor-pointer hover:bg-slate-100 hover:text-slate-700 transition-colors"
        >
          Convert Another
        </button>
      </div>
    </div>
  );
}
