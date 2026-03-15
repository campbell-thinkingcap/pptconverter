import { useCallback, useRef, useState } from "react";

interface Props {
  onFileSelect: (file: File) => void;
}

export default function UploadZone({ onFileSelect }: Props) {
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setDragOver(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) onFileSelect(file);
    },
    [onFileSelect],
  );

  const handleClick = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) onFileSelect(file);
    },
    [onFileSelect],
  );

  return (
    <div
      onClick={handleClick}
      onDragEnter={handleDragEnter}
      onDragOver={handleDragEnter}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      className={`bg-white border-2 border-dashed rounded-2xl p-12 text-center cursor-pointer transition-colors mb-8 ${
        dragOver
          ? "border-blue-600 bg-blue-50"
          : "border-slate-300 hover:border-slate-400"
      }`}
    >
      <div className="mb-4">
        <svg
          xmlns="http://www.w3.org/2000/svg"
          fill="none"
          viewBox="0 0 24 24"
          strokeWidth={1.5}
          stroke="currentColor"
          className="w-12 h-12 text-blue-600 mx-auto"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            d="M12 16.5V9.75m0 0 3 3m-3-3-3 3M6.75 19.5a4.5 4.5 0 0 1-1.41-8.775 5.25 5.25 0 0 1 10.233-2.33 3 3 0 0 1 3.758 3.848A3.752 3.752 0 0 1 18 19.5H6.75Z"
          />
        </svg>
      </div>
      <div className="text-lg font-semibold text-slate-900 mb-2">
        Drag and drop your files here
      </div>
      <div className="text-sm text-slate-400 mb-6">
        Support for PPT, PPTX, PDF, DOC, and DOCX formats up to 100MB
      </div>
      <div className="flex gap-3 justify-center">
        <button
          type="button"
          className="px-7 py-3 bg-white text-slate-700 border border-slate-300 rounded-xl text-sm font-semibold cursor-pointer hover:bg-slate-100 hover:border-slate-400 transition-colors"
        >
          Select Files
        </button>
        <button
          type="button"
          className="px-7 py-3 bg-blue-600 text-white border-none rounded-xl text-sm font-bold opacity-50 cursor-default"
          disabled
        >
          Convert Now
        </button>
      </div>
      <input
        ref={fileInputRef}
        type="file"
        accept=".ppt,.pptx,.pdf,.doc,.docx"
        className="hidden"
        onChange={handleInputChange}
      />
    </div>
  );
}
