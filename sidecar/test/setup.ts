import { mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// Tests must never touch the real ~/.excelai data directory.
process.env.EXCELAI_DATA_DIR = mkdtempSync(join(tmpdir(), 'excelai-test-'));
