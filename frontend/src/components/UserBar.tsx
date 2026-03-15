import type { UserProfile } from "../hooks/useAuth";

interface Props {
  userProfile: UserProfile;
  onSignOut: () => void;
}

export default function UserBar({ userProfile, onSignOut }: Props) {
  return (
    <div className="flex items-center gap-2.5 text-sm text-slate-500">
      {userProfile.picture && (
        <img
          src={userProfile.picture}
          alt=""
          className="w-7 h-7 rounded-full border border-slate-200"
        />
      )}
      <span>{userProfile.email}</span>
      <button
        onClick={onSignOut}
        className="bg-transparent border border-slate-300 rounded-md text-slate-500 px-2.5 py-0.5 text-xs cursor-pointer hover:text-red-500 hover:border-red-500 transition-colors"
      >
        Sign out
      </button>
    </div>
  );
}
