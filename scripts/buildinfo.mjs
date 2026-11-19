// Build identity: the VSIX version and file name are derived from git, so every commit gives a new
// minor version and every build can be traced to a commit.
//   version = 0.<commits + (uncommitted changes ? 1 : 0)>.0   (committing the build's changes gives the same number)
//   commit  = short hash of HEAD, "-dirty" appended if the working tree has uncommitted changes
// Without any commit (or without git) the version from package.json is kept and the commit is "nocommit".
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim();

export function getBuildInfo() {
  const base = JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8')).version;
  try {
    const hash = git('rev-parse', '--short=8', 'HEAD');
    const count = Number(git('rev-list', '--count', 'HEAD'));
    const dirty = git('status', '--porcelain').length > 0;
    return { version: `0.${count + (dirty ? 1 : 0)}.0`, commit: hash, dirty, label: dirty ? `${hash}-dirty` : hash };
  } catch {
    try {
      // git works but there is no commit yet
      git('rev-parse', '--git-dir');
      return { version: base, commit: 'nocommit', dirty: true, label: 'nocommit' };
    } catch {
      return { version: base, commit: 'unknown', dirty: false, label: 'unknown' };
    }
  }
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  console.log(JSON.stringify(getBuildInfo()));
}
