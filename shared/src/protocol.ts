import type { ChangeOp, ChangePreviewItem } from './changeset.js';

/**
 * WebSocket protocol between the add-in (browser) and the sidecar.
 *
 * The add-in is the workbook executor ("hands"): the sidecar sends it tool
 * execution requests and it answers. The add-in is also the UI: the sidecar
 * streams agent/chat events to it.
 *
 * Every message carries a protocol version; both sides refuse mismatched
 * majors so silent version skew is impossible.
 */

export const PROTOCOL_VERSION = 1;

export interface Envelope<T extends string, P> {
  v: number;
  id: string;
  type: T;
  payload: P;
}

/* ---------- add-in -> sidecar ---------- */

export type ClientHello = Envelope<
  'hello',
  { token: string; workbookId: string; workbookName: string; client: 'addin' | 'devtool' }
>;

export type ChatRequest = Envelope<'chat', { message: string }>;

export type ToolResult = Envelope<
  'tool_result',
  { requestId: string; ok: boolean; result?: unknown; error?: string }
>;

export type ApproveChangeSet = Envelope<'approve_changeset', { changeSetId: string }>;
export type RejectChangeSet = Envelope<'reject_changeset', { changeSetId: string }>;
export type UndoChangeSet = Envelope<'undo_changeset', { changeSetId: string }>;

export type ClientMessage =
  | ClientHello
  | ChatRequest
  | ToolResult
  | ApproveChangeSet
  | RejectChangeSet
  | UndoChangeSet;

/* ---------- sidecar -> add-in ---------- */

export type ServerHello = Envelope<
  'hello_ack',
  { ok: boolean; error?: string; sidecarVersion: string }
>;

/** Sidecar asks the add-in to execute a (read) tool or apply ops against Excel. */
export type ToolExecRequest = Envelope<
  'tool_exec',
  { requestId: string; tool: string; args: unknown }
>;

export type ApplyOpsRequest = Envelope<
  'apply_ops',
  { requestId: string; ops: ChangeOp[] }
>;

export type AgentEvent = Envelope<
  'agent_event',
  | { kind: 'text'; text: string }
  | { kind: 'tool_use'; tool: string; summary: string }
  | { kind: 'changeset_staged'; changeSetId: string; preview: ChangePreviewItem[] }
  | { kind: 'changeset_applied'; changeSetId: string }
  | { kind: 'changeset_undone'; changeSetId: string }
  | { kind: 'error'; message: string }
  | { kind: 'done' }
>;

export type ServerMessage = ServerHello | ToolExecRequest | ApplyOpsRequest | AgentEvent;

export function makeEnvelope<T extends string, P>(type: T, payload: P, id?: string): Envelope<T, P> {
  return {
    v: PROTOCOL_VERSION,
    id: id ?? globalThis.crypto.randomUUID(),
    type,
    payload,
  };
}
