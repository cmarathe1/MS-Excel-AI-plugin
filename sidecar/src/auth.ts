import { randomBytes, randomInt, timingSafeEqual } from 'node:crypto';
import type { Db } from './db.js';

/**
 * Pairing: the sidecar prints a one-time code; the add-in exchanges it for a
 * long-lived bearer token. Every HTTP request and WebSocket connection must
 * present a valid token, so other local processes cannot drive Excel
 * through us.
 */
export class Auth {
  private pairingCode: string | null;

  constructor(private readonly db: Db) {
    this.pairingCode = generateCode();
  }

  get currentPairingCode(): string | null {
    return this.pairingCode;
  }

  /** Exchange the one-time pairing code for a bearer token. */
  pair(code: string): string | null {
    if (this.pairingCode === null) return null;
    const a = Buffer.from(code.toUpperCase().padEnd(8, ' '));
    const b = Buffer.from(this.pairingCode.padEnd(8, ' '));
    if (a.length !== b.length || !timingSafeEqual(a, b)) return null;
    this.pairingCode = null; // single use
    const token = randomBytes(32).toString('hex');
    this.db
      .prepare('INSERT INTO tokens (token, created_at) VALUES (?, ?)')
      .run(token, new Date().toISOString());
    return token;
  }

  /** Re-arm pairing (e.g. user clicks "pair another client"). */
  rearm(): string {
    this.pairingCode = generateCode();
    return this.pairingCode;
  }

  verify(token: string | undefined | null): boolean {
    if (!token) return false;
    const row = this.db.prepare('SELECT token FROM tokens WHERE token = ?').get(token);
    return row !== undefined;
  }
}

function generateCode(): string {
  // 6 chars from an unambiguous alphabet.
  const alphabet = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) code += alphabet[randomInt(alphabet.length)];
  return code;
}
