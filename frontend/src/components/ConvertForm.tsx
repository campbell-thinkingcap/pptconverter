import type { AuthMode } from "../hooks/useAuth";
import FileCard from "./FileCard";

interface Props {
  file: File;
  title: string;
  mode: "paginated" | "continuous";
  theme: "light" | "dark";
  apiKey: string;
  authMode: AuthMode;
  onTitleChange: (v: string) => void;
  onModeChange: (v: "paginated" | "continuous") => void;
  onThemeChange: (v: "light" | "dark") => void;
  onApiKeyChange: (v: string) => void;
  onRemoveFile: () => void;
  onConvert: () => void;
}

export default function ConvertForm({
  file,
  title,
  mode,
  theme,
  apiKey,
  authMode,
  onTitleChange,
  onModeChange,
  onThemeChange,
  onApiKeyChange,
  onRemoveFile,
  onConvert,
}: Props) {
  return (
    <div className="bg-white border border-solid border-slate-200 rounded-2xl p-8 mb-8">
      <FileCard file={file} onRemove={onRemoveFile} />

      <input
        type="text"
        value={title}
        onChange={(e) => onTitleChange(e.target.value)}
        placeholder="Course title (optional — uses file metadata if blank)"
        className="w-full py-2.5 px-3.5 bg-white border border-slate-300 rounded-xl text-slate-900 text-sm font-[inherit] my-3 outline-none focus:border-blue-600 transition-colors placeholder:text-slate-400"
      />

      <div className="flex items-center gap-4 mb-3 text-sm">
        <span className="text-slate-500 font-semibold whitespace-nowrap">
          Display mode
        </span>
        <label className="text-slate-700 cursor-pointer flex items-center gap-1.5 whitespace-nowrap">
          <input
            type="radio"
            name="mode"
            value="paginated"
            checked={mode === "paginated"}
            onChange={() => onModeChange("paginated")}
            className="accent-blue-600"
          />
          Paginated
        </label>
        <label className="text-slate-700 cursor-pointer flex items-center gap-1.5 whitespace-nowrap">
          <input
            type="radio"
            name="mode"
            value="continuous"
            checked={mode === "continuous"}
            onChange={() => onModeChange("continuous")}
            className="accent-blue-600"
          />
          Continuous scroll
        </label>
      </div>

      <div className="flex items-center gap-4 mb-3 text-sm">
        <span className="text-slate-500 font-semibold whitespace-nowrap">
          Theme
        </span>
        <label className="text-slate-700 cursor-pointer flex items-center gap-1.5 whitespace-nowrap">
          <input
            type="radio"
            name="theme"
            value="light"
            checked={theme === "light"}
            onChange={() => onThemeChange("light")}
            className="accent-blue-600"
          />
          Light
        </label>
        <label className="text-slate-700 cursor-pointer flex items-center gap-1.5 whitespace-nowrap">
          <input
            type="radio"
            name="theme"
            value="dark"
            checked={theme === "dark"}
            onChange={() => onThemeChange("dark")}
            className="accent-blue-600"
          />
          Dark
        </label>
      </div>

      {authMode === "apikey" && (
        <input
          type="password"
          value={apiKey}
          onChange={(e) => onApiKeyChange(e.target.value)}
          placeholder="API key (required if server has API_KEY set)"
          className="w-full py-2.5 px-3.5 bg-white border border-slate-300 rounded-xl text-slate-900 text-sm font-[inherit] mb-3 outline-none focus:border-blue-600 transition-colors placeholder:text-slate-400"
        />
      )}

      <button
        onClick={onConvert}
        className="w-full py-3.5 bg-blue-600 text-white border-none rounded-xl text-base font-bold cursor-pointer hover:bg-blue-700 transition-colors mt-1"
      >
        Convert Now
      </button>
    </div>
  );
}
