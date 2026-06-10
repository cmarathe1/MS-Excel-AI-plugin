import { mkdirSync, readFileSync, writeFileSync, chmodSync, existsSync } from 'node:fs';
import { homedir } from 'node:os';
import { join } from 'node:path';
import type { ProviderConfig } from './providers/types.js';

export interface Settings {
  provider?: ProviderConfig;
}

export function dataDir(): string {
  const dir = process.env.EXCELAI_DATA_DIR ?? join(homedir(), '.excelai');
  mkdirSync(dir, { recursive: true });
  return dir;
}

const SETTINGS_FILE = 'settings.json';

/**
 * Provider settings live in a 0600 JSON file in the data directory.
 * (OS keychain integration is planned; a file the user's account owns is the
 * honest v1 — it is never sent anywhere and never committed.)
 */
export function loadSettings(): Settings {
  const file = join(dataDir(), SETTINGS_FILE);
  if (!existsSync(file)) return {};
  try {
    return JSON.parse(readFileSync(file, 'utf8')) as Settings;
  } catch {
    return {};
  }
}

export function saveSettings(settings: Settings): void {
  const file = join(dataDir(), SETTINGS_FILE);
  writeFileSync(file, JSON.stringify(settings, null, 2), { mode: 0o600 });
  chmodSync(file, 0o600);
}

export function redactSettings(settings: Settings): Settings {
  if (!settings.provider) return settings;
  const { apiKey, ...rest } = settings.provider;
  return { provider: { ...rest, ...(apiKey ? { apiKey: '••••' + apiKey.slice(-4) } : {}) } };
}
