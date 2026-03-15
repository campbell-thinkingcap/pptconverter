interface Props {
  message: string;
  onRetry: () => void;
}

export default function ErrorPanel({ message, onRetry }: Props) {
  return (
    <div className="bg-white border border-red-200 rounded-2xl p-8 text-center mb-8">
      <div className="w-14 h-14 bg-red-50 rounded-full flex items-center justify-center mx-auto mb-4">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          fill="none"
          viewBox="0 0 24 24"
          strokeWidth={2}
          stroke="currentColor"
          className="w-7 h-7 text-red-500"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M12 9v3.75m9-.75a9 9 0 1 1-18 0 9 9 0 0 1 18 0Zm-9 3.75h.008v.008H12v-.008Z"
          />
        </svg>
      </div>
      <div className="text-lg font-bold text-red-500 mb-1.5">
        Conversion Failed
      </div>
      <div className="text-sm text-slate-500 mb-5">{message}</div>
      <button
        onClick={onRetry}
        className="px-6 py-2.5 bg-blue-600 text-white border-none rounded-xl text-sm font-bold cursor-pointer hover:bg-blue-700 transition-colors"
      >
        Try Again
      </button>
    </div>
  );
}
