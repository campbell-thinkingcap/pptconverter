export default function Footer() {
  return (
    <footer className="border-t border-slate-200 py-6 px-8 bg-white shrink-0">
      <div className="max-w-[960px] mx-auto flex items-center justify-between text-sm text-slate-400">
        <span>&copy; {new Date().getFullYear()} SCORMFlow. All rights reserved.</span>
        <div className="flex gap-6">
          <a
            href="/docs"
            className="text-slate-500 no-underline text-sm hover:text-slate-900"
          >
            API Documentation
          </a>
        </div>
      </div>
    </footer>
  );
}
