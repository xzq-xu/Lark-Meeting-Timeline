import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { demoTimeline } from './normalize.mjs';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const defaultDataDir = process.env.TIMELINE_DATA_DIR || join(root, 'data');
const defaultPath = process.env.TIMELINE_STATE_PATH || join(defaultDataDir, 'state.json');

export class TimelineStore {
  constructor(path = defaultPath) {
    this.path = path;
    this.state = null;
  }

  async load() {
    if (this.state) return this.state;
    try {
      this.state = JSON.parse(await readFile(this.path, 'utf8'));
    } catch {
      this.state = demoTimeline();
      await this.save(this.state);
    }
    return this.state;
  }

  async save(nextState) {
    this.state = {
      ...nextState,
      updated_at: new Date().toISOString(),
    };
    await mkdir(dirname(this.path), { recursive: true });
    await writeFile(this.path, JSON.stringify(this.state, null, 2));
    return this.state;
  }

  async resetDemo() {
    return this.save(demoTimeline());
  }
}
