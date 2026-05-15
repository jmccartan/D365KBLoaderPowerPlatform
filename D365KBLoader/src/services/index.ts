import type { KbLoaderService } from './KbLoaderService';
import { MockKbLoaderService } from './MockKbLoaderService';
import { PowerPlatformKbLoaderService } from './PowerPlatformKbLoaderService';

/**
 * Returns the mock service unless VITE_USE_REAL_CONNECTORS=true is set,
 * in which case the Power Platform service is used. The flag lives in
 * `.env.local` so the same code runs both during local dev and after deploy.
 */
export function getService(): KbLoaderService {
  const useReal = import.meta.env?.VITE_USE_REAL_CONNECTORS === 'true';
  return useReal ? new PowerPlatformKbLoaderService() : new MockKbLoaderService();
}
