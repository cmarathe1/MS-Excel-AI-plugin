import { mkdirSync, readFileSync, writeFileSync, chmodSync, existsSync } from 'node:fs';
import { randomUUID } from 'node:crypto';
import { homedir } from 'node:os';
import { join } from 'node:path';
import type { ProviderConfig } from './providers/types.js';

/** A saved model configuration the user can switch between. */
export interface StoredProvider extends ProviderConfig {
  id: string;
}

export interface Settings {
  providers: StoredProvider[];
  activeProviderId: string | null;
}

interface RawSettingsFile {
  providers?: StoredProvider[];
  activeProviderId?: string | null;
  /** Legacy single-provider shape (pre model-list). */
  provider?: ProviderConfig;
}

export function dataDir(): string {
  const dir = process.env.EXCELAI_DATA_DIR ?? join(homedir(), '.excelai');
  mkdirSync(dir, { recursive: true });
  return dir;
}

const SETTINGS_FILE = 'settings.json';

/**
 * Settings live in a 0600 JSON file in the data directory. (OS keychain
 * integration is planned; a file owned by the user's account is the honest
 * v1 — it is never sent anywhere and never committed.)
 */
export function loadSettings(): Settings {
  const file = join(dataDir(), SETTINGS_FILE);
  if (!existsSync(file)) return { providers: [], activeProviderId: null };
  let raw: RawSettingsFile;
  try {
    raw = JSON.parse(readFileSync(file, 'utf8')) as RawSettingsFile;
  } catch {
    return { providers: [], activeProviderId: null };
  }
  // Migrate the legacy single-provider shape.
  if (!raw.providers && raw.provider) {
    const migrated: StoredProvider = { ...raw.provider, id: randomUUID() };
    const settings: Settings = { providers: [migrated], activeProviderId: migrated.id };
    saveSettings(settings);
    return settings;
  }
  const providers = raw.providers ?? [];
  const activeProviderId =
    raw.activeProviderId && providers.some((p) => p.id === raw.activeProviderId)
      ? raw.activeProviderId
      : (providers[0]?.id ?? null);
  return { providers, activeProviderId };
}

export function saveSettings(settings: Settings): void {
  const file = join(dataDir(), SETTINGS_FILE);
  writeFileSync(file, JSON.stringify(settings, null, 2), { mode: 0o600 });
  chmodSync(file, 0o600);
}

export function activeProvider(settings: Settings): StoredProvider | null {
  return settings.providers.find((p) => p.id === settings.activeProviderId) ?? null;
}

export interface RedactedProvider {
  id: string;
  kind: StoredProvider['kind'];
  model: string;
  baseUrl?: string;
  /** Masked: present only to show that a key is stored. */
  apiKey?: string;
}

export function redactProvider(p: StoredProvider): RedactedProvider {
  const out: RedactedProvider = { id: p.id, kind: p.kind, model: p.model };
  if (p.baseUrl) out.baseUrl = p.baseUrl;
  if (p.apiKey) out.apiKey = `••••${p.apiKey.slice(-4)}`;
  return out;
}
