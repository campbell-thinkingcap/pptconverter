interface Props {
  phase: "uploading" | "converting";
  uploadProgress: number;
}

export default function ProgressPanel({ phase, uploadProgress }: Props) {
  const isUploading = phase === "uploading";

  return (
    <div className="bg-white border border-slate-200 rounded-2xl p-8 text-center mb-8">
      <div className="text-base font-semibold text-slate-900 mb-4">
        {isUploading ? "Uploading..." : "Converting..."}
      </div>
      <div className="w-full h-2 bg-slate-200 rounded overflow-hidden mb-3">
        <div
          className={`h-full bg-gradient-to-r from-blue-600 to-blue-400 rounded transition-[width] duration-300 ${
            !isUploading ? "w-full animate-pulse" : ""
          }`}
          style={isUploading ? { width: `${uploadProgress}%` } : undefined}
        />
      </div>
      <div className="text-sm text-slate-400">
        {isUploading
          ? `${uploadProgress}% uploaded`
          : "Rendering slides \u2014 this may take a moment"}
      </div>
    </div>
  );
}
