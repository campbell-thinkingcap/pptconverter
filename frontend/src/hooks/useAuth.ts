import { useCallback, useEffect, useRef, useState } from "react";
import { getConfig } from "../api/client";

export type AuthMode = "loading" | "none" | "google" | "apikey";

export interface UserProfile {
  email: string;
  picture: string;
}

export interface AuthState {
  authMode: AuthMode;
  googleIdToken: string;
  userProfile: UserProfile | null;
  isAuthOverlayVisible: boolean;
  authError: string;
  signInButtonRef: React.RefObject<HTMLDivElement | null>;
  signOut: () => void;
  /** Force re-show the overlay (e.g. after 401) */
  showOverlay: (errorMsg?: string) => void;
}

export function useAuth(): AuthState {
  const [authMode, setAuthMode] = useState<AuthMode>("loading");
  const [googleIdToken, setGoogleIdToken] = useState("");
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [overlayVisible, setOverlayVisible] = useState(false);
  const [authError, setAuthError] = useState("");
  const signInButtonRef = useRef<HTMLDivElement | null>(null);
  const clientIdRef = useRef("");

  const renderButton = useCallback(() => {
    if (signInButtonRef.current) {
      google.accounts.id.renderButton(signInButtonRef.current, {
        theme: "filled_black",
        size: "large",
        shape: "rectangular",
        width: 280,
      });
    }
  }, []);

  const handleCredential = useCallback(
    (response: { credential: string }) => {
      const token = response.credential;
      setGoogleIdToken(token);
      try {
        const parts = token.split(".");
        const payload = JSON.parse(
          atob(parts[1]!.replace(/-/g, "+").replace(/_/g, "/")),
        ) as { email?: string; picture?: string };
        setUserProfile({
          email: payload.email ?? "",
          picture: payload.picture ?? "",
        });
      } catch {
        // If JWT decode fails, still accept the token
      }
      setOverlayVisible(false);
      setAuthError("");
    },
    [],
  );

  const waitForGIS = useCallback(
    (fn: () => void) => {
      if (window.google?.accounts?.id) {
        fn();
      } else {
        setTimeout(() => waitForGIS(fn), 100);
      }
    },
    [],
  );

  // Init: fetch /api/config and decide auth mode
  useEffect(() => {
    getConfig()
      .then((cfg) => {
        const clientId = cfg.googleClientId || "";
        clientIdRef.current = clientId;
        if (clientId) {
          setAuthMode("google");
          setOverlayVisible(true);
          waitForGIS(() => {
            google.accounts.id.initialize({
              client_id: clientId,
              callback: handleCredential,
              auto_select: false,
              cancel_on_tap_outside: false,
            });
            renderButton();
          });
        } else {
          setAuthMode("apikey");
          setOverlayVisible(false);
        }
      })
      .catch(() => {
        setAuthMode("apikey");
        setOverlayVisible(false);
      });
  }, [handleCredential, renderButton, waitForGIS]);

  // Re-render GIS button whenever overlay becomes visible
  useEffect(() => {
    if (overlayVisible && authMode === "google") {
      waitForGIS(renderButton);
    }
  }, [overlayVisible, authMode, renderButton, waitForGIS]);

  const signOut = useCallback(() => {
    setGoogleIdToken("");
    setUserProfile(null);
    if (authMode === "google") {
      google.accounts.id.disableAutoSelect();
      setOverlayVisible(true);
      setAuthError("");
    }
  }, [authMode]);

  const showOverlay = useCallback((errorMsg?: string) => {
    setGoogleIdToken("");
    setUserProfile(null);
    if (errorMsg) setAuthError(errorMsg);
    setOverlayVisible(true);
  }, []);

  return {
    authMode,
    googleIdToken,
    userProfile,
    isAuthOverlayVisible: overlayVisible,
    authError,
    signInButtonRef,
    signOut,
    showOverlay,
  };
}
