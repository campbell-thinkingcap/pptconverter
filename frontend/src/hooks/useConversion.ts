import { useCallback, useRef, useState } from "react";
import { convertPresentation, triggerBlobDownload } from "../api/client";

type Phase = "idle" | "file-selected" | "uploading" | "converting" | "result" | "error";

export interface ConversionState {
  phase: Phase;
  selectedFile: File | null;
  uploadProgress: number;
  resultFilename: string;
  errorMessage: string;
  title: string;
  mode: "paginated" | "continuous";
  theme: "light" | "dark";
  apiKey: string;
  setTitle: (v: string) => void;
  setMode: (v: "paginated" | "continuous") => void;
  setTheme: (v: "light" | "dark") => void;
  setApiKey: (v: string) => void;
  handleFileSelect: (file: File) => void;
  handleRemoveFile: () => void;
  handleConvert: (googleIdToken: string) => void;
  handleDownloadAgain: () => void;
  handleReset: () => void;
}

const ALLOWED = [".ppt", ".pptx", ".pdf", ".doc", ".docx"];

function getExtension(name: string): string {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.substring(i).toLowerCase() : "";
}

export function useConversion(): ConversionState {
  const [phase, setPhase] = useState<Phase>("idle");
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [resultFilename, setResultFilename] = useState("");
  const [errorMessage, setErrorMessage] = useState("");

  const [title, setTitle] = useState("");
  const [mode, setMode] = useState<"paginated" | "continuous">("paginated");
  const [theme, setTheme] = useState<"light" | "dark">("light");
  const [apiKey, setApiKey] = useState("");

  const blobUrlRef = useRef<string | null>(null);
  const abortRef = useRef<(() => void) | null>(null);

  const revokeBlobUrl = useCallback(() => {
    if (blobUrlRef.current) {
      URL.revokeObjectURL(blobUrlRef.current);
      blobUrlRef.current = null;
    }
  }, []);

  const handleReset = useCallback(() => {
    abortRef.current?.();
    abortRef.current = null;
    revokeBlobUrl();
    setPhase("idle");
    setSelectedFile(null);
    setUploadProgress(0);
    setResultFilename("");
    setErrorMessage("");
    setTitle("");
  }, [revokeBlobUrl]);

  const handleFileSelect = useCallback((file: File) => {
    const ext = getExtension(file.name);
    if (!ALLOWED.includes(ext)) {
      setErrorMessage("Only .ppt, .pptx, .pdf, .doc, and .docx files are supported.");
      setPhase("error");
      return;
    }
    setSelectedFile(file);
    setPhase("file-selected");
    setErrorMessage("");
  }, []);

  const handleRemoveFile = useCallback(() => {
    setSelectedFile(null);
    setPhase("idle");
    setTitle("");
  }, []);

  const handleConvert = useCallback(
    (googleIdToken: string) => {
      if (!selectedFile) return;

      setPhase("uploading");
      setUploadProgress(0);

      const { promise, abort } = convertPresentation({
        file: selectedFile,
        title,
        mode,
        theme,
        googleIdToken,
        apiKey,
        onUploadProgress: (pct) => {
          setUploadProgress(pct);
          if (pct >= 100) setPhase("converting");
        },
      });

      abortRef.current = abort;

      promise
        .then((result) => {
          blobUrlRef.current = result.blobUrl;
          setResultFilename(result.filename);
          triggerBlobDownload(result.blobUrl, result.filename);
          setPhase("result");
        })
        .catch((err: Error) => {
          setErrorMessage(err.message);
          setPhase("error");
        });
    },
    [selectedFile, title, mode, theme, apiKey],
  );

  const handleDownloadAgain = useCallback(() => {
    if (blobUrlRef.current && resultFilename) {
      triggerBlobDownload(blobUrlRef.current, resultFilename);
    }
  }, [resultFilename]);

  return {
    phase,
    selectedFile,
    uploadProgress,
    resultFilename,
    errorMessage,
    title,
    mode,
    theme,
    apiKey,
    setTitle,
    setMode,
    setTheme,
    setApiKey,
    handleFileSelect,
    handleRemoveFile,
    handleConvert,
    handleDownloadAgain,
    handleReset,
  };
}
