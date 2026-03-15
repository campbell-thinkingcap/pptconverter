export interface ConfigResponse {
  googleClientId: string;
}

export interface ConvertOptions {
  file: File;
  title: string;
  mode: "paginated" | "continuous";
  theme: "light" | "dark";
  googleIdToken: string;
  apiKey: string;
  onUploadProgress: (percent: number) => void;
}

export interface ConvertResult {
  blobUrl: string;
  filename: string;
}
