/* Google Identity Services (GIS) SDK type stubs */

interface GoogleAccountsId {
  initialize(config: {
    client_id: string;
    callback: (response: { credential: string }) => void;
    auto_select?: boolean;
    cancel_on_tap_outside?: boolean;
  }): void;
  renderButton(
    parent: HTMLElement,
    options: {
      theme?: string;
      size?: string;
      shape?: string;
      width?: number;
    },
  ): void;
  disableAutoSelect(): void;
}

interface GoogleAccounts {
  id: GoogleAccountsId;
}

interface Google {
  accounts: GoogleAccounts;
}

declare const google: Google;

interface Window {
  google?: Google;
}
