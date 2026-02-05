interface ImportMetaEnv {
  readonly VITE_BETTER_AUTH_URL?: string;
  readonly VITE_GOOGLE_CLIENT_ID?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
