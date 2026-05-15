/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_USE_REAL_CONNECTORS?: string;
}
interface ImportMeta {
  readonly env: ImportMetaEnv;
}
