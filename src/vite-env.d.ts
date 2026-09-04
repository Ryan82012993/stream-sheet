/// <reference types="vite/client" />

interface Window {
  showOpenFilePicker?: (options?: {
    types?: Array<{
      description?: string;
      accept?: Record<string, string[]>;
    }>;
    multiple?: boolean;
  }) => Promise<any[]>;
}

