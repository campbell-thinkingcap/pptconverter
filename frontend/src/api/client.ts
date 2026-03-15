import type { ConfigResponse, ConvertOptions, ConvertResult } from "./types";

export async function getConfig(): Promise<ConfigResponse> {
  const res = await fetch("/api/config");
  return res.json() as Promise<ConfigResponse>;
}

function parseFilename(header: string | null): string {
  if (!header) return "scorm_package.zip";
  const match = header.match(/filename="?([^";]+)"?/);
  return match?.[1] ?? "scorm_package.zip";
}

export function convertPresentation(
  opts: ConvertOptions,
): { promise: Promise<ConvertResult>; abort: () => void } {
  const xhr = new XMLHttpRequest();
  const abort = () => xhr.abort();

  const promise = new Promise<ConvertResult>((resolve, reject) => {
    xhr.open("POST", "/api/convert");

    if (opts.googleIdToken) {
      xhr.setRequestHeader("X-Google-Token", opts.googleIdToken);
    } else if (opts.apiKey) {
      xhr.setRequestHeader("X-API-Key", opts.apiKey);
    }

    xhr.responseType = "blob";

    xhr.upload.addEventListener("progress", (e) => {
      if (e.lengthComputable) {
        opts.onUploadProgress(Math.round((e.loaded / e.total) * 100));
      }
    });

    xhr.addEventListener("load", () => {
      if (xhr.status === 200) {
        const filename = parseFilename(
          xhr.getResponseHeader("Content-Disposition"),
        );
        const blobUrl = URL.createObjectURL(xhr.response as Blob);
        resolve({ blobUrl, filename });
      } else {
        // Try to parse JSON error from blob
        const reader = new FileReader();
        reader.onload = () => {
          try {
            const json = JSON.parse(reader.result as string) as {
              detail?: string;
            };
            reject(new Error(json.detail ?? `Conversion failed (HTTP ${xhr.status})`));
          } catch {
            reject(new Error(`Conversion failed (HTTP ${xhr.status})`));
          }
        };
        reader.onerror = () => {
          reject(new Error(`Conversion failed (HTTP ${xhr.status})`));
        };
        reader.readAsText(xhr.response as Blob);
      }
    });

    xhr.addEventListener("error", () => {
      reject(new Error("Network error. Is the server running?"));
    });

    const form = new FormData();
    form.append("file", opts.file);
    if (opts.title) form.append("title", opts.title);
    form.append("mode", opts.mode);
    form.append("theme", opts.theme);

    xhr.send(form);
  });

  return { promise, abort };
}

export function triggerBlobDownload(blobUrl: string, filename: string): void {
  const a = document.createElement("a");
  a.href = blobUrl;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}
