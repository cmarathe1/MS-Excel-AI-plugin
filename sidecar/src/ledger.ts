import type { Db } from './db.js';

export interface LedgerSummary {
  totalCalls: number;
  inputTokens: number;
  outputTokens: number;
  byFeature: { feature: string; calls: number; inputTokens: number; outputTokens: number }[];
}

/** Cost & audit ledger: every model call is recorded locally. */
export class Ledger {
  constructor(private readonly db: Db) {}

  record(entry: {
    provider: string;
    model: string;
    feature: 'chat' | 'functions' | 'probe';
    inputTokens: number;
    outputTokens: number;
  }): void {
    this.db
      .prepare(
        'INSERT INTO ledger (at, provider, model, feature, input_tokens, output_tokens) VALUES (?, ?, ?, ?, ?, ?)',
      )
      .run(
        new Date().toISOString(),
        entry.provider,
        entry.model,
        entry.feature,
        entry.inputTokens,
        entry.outputTokens,
      );
  }

  summary(): LedgerSummary {
    const total = this.db
      .prepare(
        'SELECT COUNT(*) AS calls, COALESCE(SUM(input_tokens),0) AS inp, COALESCE(SUM(output_tokens),0) AS outp FROM ledger',
      )
      .get() as { calls: number; inp: number; outp: number };
    const byFeature = this.db
      .prepare(
        'SELECT feature, COUNT(*) AS calls, COALESCE(SUM(input_tokens),0) AS inp, COALESCE(SUM(output_tokens),0) AS outp FROM ledger GROUP BY feature',
      )
      .all() as { feature: string; calls: number; inp: number; outp: number }[];
    return {
      totalCalls: total.calls,
      inputTokens: total.inp,
      outputTokens: total.outp,
      byFeature: byFeature.map((r) => ({
        feature: r.feature,
        calls: r.calls,
        inputTokens: r.inp,
        outputTokens: r.outp,
      })),
    };
  }
}
