import { DatabaseSync } from 'node:sqlite';
import { join } from 'node:path';
import { dataDir } from './config.js';

export type Db = DatabaseSync;

const MIGRATIONS = `
CREATE TABLE IF NOT EXISTS tokens (
  token TEXT PRIMARY KEY,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS fn_cache (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS ledger (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  at TEXT NOT NULL,
  provider TEXT NOT NULL,
  model TEXT NOT NULL,
  feature TEXT NOT NULL,
  input_tokens INTEGER NOT NULL,
  output_tokens INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS memory (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  scope TEXT NOT NULL,            -- 'user' or workbook id
  content TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE VIRTUAL TABLE IF NOT EXISTS memory_fts USING fts5(
  content,
  content='memory',
  content_rowid='id'
);
CREATE TRIGGER IF NOT EXISTS memory_ai AFTER INSERT ON memory BEGIN
  INSERT INTO memory_fts(rowid, content) VALUES (new.id, new.content);
END;
CREATE TRIGGER IF NOT EXISTS memory_ad AFTER DELETE ON memory BEGIN
  INSERT INTO memory_fts(memory_fts, rowid, content) VALUES ('delete', old.id, old.content);
END;
`;

export function openDb(path?: string): Db {
  const db = new DatabaseSync(path ?? join(dataDir(), 'excelai.db'));
  db.exec('PRAGMA journal_mode = WAL;');
  db.exec(MIGRATIONS);
  return db;
}
