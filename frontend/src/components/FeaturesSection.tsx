export default function FeaturesSection() {
  return (
    <section className="bg-white py-12">
      <div className="max-w-[960px] mx-auto px-8">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
          <div>
            <div className="mb-3">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-6 h-6 text-amber-500"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="m3.75 13.5 10.5-11.25L12 10.5h8.25L9.75 21.75 12 13.5H3.75Z"
                />
              </svg>
            </div>
            <h3 className="text-base font-bold text-slate-900 mb-2">
              Lightning Fast
            </h3>
            <p className="text-sm text-slate-500 leading-relaxed">
              Our high-speed servers process your files in seconds, not minutes.
            </p>
          </div>

          <div>
            <div className="mb-3">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-6 h-6 text-blue-600"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M9 12.75 11.25 15 15 9.75m-3-7.036A11.959 11.959 0 0 1 3.598 6 11.99 11.99 0 0 0 3 9.749c0 5.592 3.824 10.29 9 11.623 5.176-1.332 9-6.03 9-11.622 0-1.31-.21-2.571-.598-3.751h-.152c-3.196 0-6.1-1.248-8.25-3.285Z"
                />
              </svg>
            </div>
            <h3 className="text-base font-bold text-slate-900 mb-2">
              Secure Processing
            </h3>
            <p className="text-sm text-slate-500 leading-relaxed">
              Files are automatically deleted after conversion for your privacy.
            </p>
          </div>

          <div>
            <div className="mb-3">
              <svg
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 24 24"
                strokeWidth={1.5}
                stroke="currentColor"
                className="w-6 h-6 text-blue-600"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="m2.25 15.75 5.159-5.159a2.25 2.25 0 0 1 3.182 0l5.159 5.159m-1.5-1.5 1.409-1.409a2.25 2.25 0 0 1 3.182 0l2.909 2.909M3.75 21h16.5A2.25 2.25 0 0 0 22.5 18.75V5.25A2.25 2.25 0 0 0 20.25 3H3.75A2.25 2.25 0 0 0 1.5 5.25v13.5A2.25 2.25 0 0 0 3.75 21Z"
                />
              </svg>
            </div>
            <h3 className="text-base font-bold text-slate-900 mb-2">
              LMS Compatible
            </h3>
            <p className="text-sm text-slate-500 leading-relaxed">
              Fully compatible with Moodle, Canvas, Blackboard, and more.
            </p>
          </div>
        </div>
      </div>
    </section>
  );
}
