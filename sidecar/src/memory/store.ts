import type { Db } from '../db.js';

export interface MemoryEntry {
  id: number;
  scope: string;
  content: string;
  createdAt: string;
}

/**
 * Layered memory store (v1): user-scope and workbook-scope entries with FTS5
 * keyword retrieval. Embedding-based retrieval plugs in behind the same
 * interface later; FTS keeps it dependency-free and fully local.
 */
export class MemoryStore {
  constructor(private readonly db: Db) {}

  add(scope: string, content: string): MemoryEntry {
    const createdAt = new Date().toISOString();
    const result = this.db
      .prepare('INSERT INTO memory (scope, content, created_at) VALUES (?, ?, ?)')
      .run(scope, content.trim(), createdAt);
    return { id: Number(result.lastInsertRowid), scope, content: content.trim(), createdAt };
  }

  remove(id: number): void {
    this.db.prepare('DELETE FROM memory WHERE id = ?').run(id);
  }

  list(scope?: string): MemoryEntry[] {
    const rows = scope
      ? this.db
          .prepare('SELECT id, scope, content, created_at FROM memory WHERE scope = ? ORDER BY id DESC')
          .all(scope)
      : this.db.prepare('SELECT id, scope, content, created_at FROM memory ORDER BY id DESC').all();
    return (rows as { id: number; scope: string; content: string; created_at: string }[]).map(mapRow);
  }

  /**
   * Retrieve entries relevant to a query from the given scopes
   * (typically ['user', workbookId]).
   */
  search(query: string, scopes: string[], limit = 5): MemoryEntry[] {
    const terms = query
      .toLowerCase()
      .split(/[^a-z0-9]+/)
      .filter((t) => t.length > 2)
      .slice(0, 12);
    if (terms.length === 0) return [];
    const ftsQuery = terms.map((t) => `"${t}"`).join(' OR ');
    const placeholders = scopes.map(() => '?').join(',');
    try {
      const rows = this.db
        .prepare(
          `SELECT m.id, m.scope, m.content, m.created_at
           FROM memory_fts f JOIN memory m ON m.id = f.rowid
           WHERE memory_fts MATCH ? AND m.scope IN (${placeholders})
           ORDER BY rank LIMIT ?`,
        )
        .all(ftsQuery, ...scopes, limit);
      return (rows as { id: number; scope: string; content: string; created_at: string }[]).map(mapRow);
    } catch {
      return []; // malformed FTS query must never break a chat turn
    }
  }
}

function mapRow(r: { id: number; scope: string; content: string; created_at: string }): MemoryEntry {
  return { id: r.id, scope: r.scope, content: r.content, createdAt: r.created_at };
}
