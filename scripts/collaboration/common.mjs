// Copyright (c) 2026 Eigenpal, Inc. Licensed under Apache-2.0.
// Release baseline design adapted from DOCX Editor scripts/collaboration,
// commit 13f08556889ef7c192023f92aa865cc9cc8e3fe4.
import { createHash } from 'node:crypto';
import { fork, spawnSync } from 'node:child_process';
import { copyFileSync, readFileSync, realpathSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
export const root = resolve(import.meta.dirname, '../..');
export const baseline = '0.3.1';
export const names = ['@pptx-editor.dev/i18n', '@pptx-editor.dev/core', '@pptx-editor.dev/pro'];
export const read = (file) => JSON.parse(readFileSync(file, 'utf8'));
export const write = (file, value) => writeFileSync(file, `${JSON.stringify(value, null, 2)}\n`);
export const hash = (file) => createHash('sha256').update(readFileSync(file)).digest('hex');
export const stable = (version) => /^\d+\.\d+\.\d+$/.test(version);
export function compare(a, b) {
  const left = a.split('.').map(Number),
    right = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) if (left[i] !== right[i]) return left[i] - right[i];
  return 0;
}
export function run(command, args, cwd = root, trim = true) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: 'utf8',
    timeout: 180_000,
    maxBuffer: 16 * 1024 * 1024,
  });
  if (result.status !== 0) throw new Error(`${command} failed: ${result.stderr || result.error}`);
  return trim ? result.stdout.trim() : result.stdout;
}
export const registry = (name, version) =>
  JSON.parse(run('npm', ['view', `${name}@${version}`, '--json', '--prefer-online']));
export function install(directory, locked = false) {
  run('npm', [locked ? 'ci' : 'install', '--ignore-scripts', '--no-audit', '--no-fund'], directory);
  const lock = read(resolve(directory, 'package-lock.json'));
  for (const name of ['@pptx-editor.dev/core', 'yjs']) {
    const entries = Object.entries(lock.packages).filter(([key]) =>
      key.endsWith(`node_modules/${name}`)
    );
    if (entries.length !== 1 || entries[0][1].link)
      throw new Error(`Expected one isolated ${name}`);
    if (!realpathSync(resolve(directory, entries[0][0])).startsWith(`${realpathSync(directory)}/`))
      throw new Error(`External package link: ${name}`);
  }
  copyFileSync(resolve(root, 'scripts/collaboration/worker.mjs'), resolve(directory, 'worker.mjs'));
}
export function peer(directory) {
  const child = fork(resolve(directory, 'worker.mjs'), [], {
    cwd: directory,
    stdio: ['ignore', 'ignore', 'inherit', 'ipc'],
  });
  return {
    call(message) {
      return new Promise((resolveCall, reject) => {
        const timer = setTimeout(() => finish(new Error('Worker timed out')), 30_000);
        const onExit = (code) => finish(new Error(`Worker exited: ${code}`));
        const onMessage = (reply) =>
          finish(reply.error ? new Error(reply.error) : null, reply.result);
        function finish(error, value) {
          clearTimeout(timer);
          child.off('exit', onExit);
          child.off('message', onMessage);
          if (error) reject(error);
          else resolveCall(value);
        }
        child.once('exit', onExit);
        child.once('message', onMessage);
        child.send(message);
      });
    },
    close() {
      child.kill();
    },
  };
}
