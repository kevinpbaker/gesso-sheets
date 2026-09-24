import { defineConfig } from 'vitest/config';

/**
 * The sheet model runs in node, with no browser and no framework.
 *
 * That is the whole point of Phase 1: everything under `src/sheet` is
 * plain TypeScript, so its specs need no DOM, no canvas and no worker,
 * and they run in milliseconds. The spike in `src/` is excluded rather
 * than configured — it has no specs, and the day it does they will
 * want a different environment.
 */
export default defineConfig({
  test: {
    include: ['src/sheet/**/*.spec.ts'],
    environment: 'node'
  }
});
