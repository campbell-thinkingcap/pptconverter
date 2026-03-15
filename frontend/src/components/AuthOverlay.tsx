interface Props {
  visible: boolean;
  errorMessage: string;
  signInButtonRef: React.RefObject<HTMLDivElement | null>;
}

export default function AuthOverlay({
  visible,
  errorMessage,
  signInButtonRef,
}: Props) {
  if (!visible) return null;

  return (
    <div className="fixed inset-0 bg-slate-50 flex items-center justify-center z-50 flex-col gap-6">
      <div className="bg-white border border-slate-200 rounded-2xl p-12 px-10 text-center max-w-sm w-full shadow-sm">
        <h2 className="text-xl font-bold text-slate-900 mb-2">
          Sign in to continue
        </h2>
        <p className="text-sm text-slate-500 mb-8">
          Use your organisation Google account to access the converter.
        </p>
        <div ref={signInButtonRef} className="flex justify-center" />
        {errorMessage && (
          <div className="text-red-500 text-sm mt-4">{errorMessage}</div>
        )}
      </div>
    </div>
  );
}
