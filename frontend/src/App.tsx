import { useAuth } from "./hooks/useAuth";
import { useConversion } from "./hooks/useConversion";
import Navbar from "./components/Navbar";
import AuthOverlay from "./components/AuthOverlay";
import UploadZone from "./components/UploadZone";
import ConvertForm from "./components/ConvertForm";
import ProgressPanel from "./components/ProgressPanel";
import ResultPanel from "./components/ResultPanel";
import ErrorPanel from "./components/ErrorPanel";
import SupportedFormats from "./components/SupportedFormats";
import FeaturesSection from "./components/FeaturesSection";
import Footer from "./components/Footer";

export default function App() {
  const auth = useAuth();
  const conv = useConversion();

  return (
    <>
      <AuthOverlay
        visible={auth.isAuthOverlayVisible}
        errorMessage={auth.authError}
        signInButtonRef={auth.signInButtonRef}
      />

      <Navbar userProfile={auth.userProfile} onSignOut={auth.signOut} />

      <main className="flex-1 max-w-[960px] w-full mx-auto py-12 px-8">
        {/* Hero */}
        <div className="mb-10">
          <h1 className="text-4xl font-extrabold text-slate-900 leading-tight mb-4">
            Convert your presentations to{" "}
            <span className="text-blue-600">SCORM packages</span> in seconds.
          </h1>
          <p className="text-base text-slate-500 leading-relaxed max-w-xl">
            Professional tool for PPT and PDF SCORM conversion. Upload your
            files and get LMS-ready content instantly.
          </p>
        </div>

        {/* Upload / Convert / Progress / Result / Error */}
        {conv.phase === "idle" && (
          <UploadZone onFileSelect={conv.handleFileSelect} />
        )}

        {conv.phase === "file-selected" && conv.selectedFile && (
          <ConvertForm
            file={conv.selectedFile}
            title={conv.title}
            mode={conv.mode}
            theme={conv.theme}
            apiKey={conv.apiKey}
            authMode={auth.authMode}
            onTitleChange={conv.setTitle}
            onModeChange={conv.setMode}
            onThemeChange={conv.setTheme}
            onApiKeyChange={conv.setApiKey}
            onRemoveFile={conv.handleRemoveFile}
            onConvert={() => conv.handleConvert(auth.googleIdToken)}
          />
        )}

        {(conv.phase === "uploading" || conv.phase === "converting") && (
          <ProgressPanel
            phase={conv.phase}
            uploadProgress={conv.uploadProgress}
          />
        )}

        {conv.phase === "result" && (
          <ResultPanel
            filename={conv.resultFilename}
            onDownloadAgain={conv.handleDownloadAgain}
            onConvertAnother={conv.handleReset}
          />
        )}

        {conv.phase === "error" && (
          <ErrorPanel message={conv.errorMessage} onRetry={conv.handleReset} />
        )}

        {/* Supported Formats */}
        <SupportedFormats />
      </main>

      <hr className="border-t border-slate-200 m-0" />
      <FeaturesSection />
      <Footer />
    </>
  );
}
