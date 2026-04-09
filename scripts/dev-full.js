import { spawn } from 'child_process';

const children = [];

function startProcess(name, command, args) {
  const child = spawn(command, args, {
    stdio: 'inherit',
    shell: process.platform === 'win32'
  });

  child.on('exit', (code, signal) => {
    if (signal) {
      console.log(`${name} stopped with signal ${signal}`);
      return;
    }

    if (code !== 0) {
      console.error(`${name} exited with code ${code}`);
      shutdown(code ?? 1);
    }
  });

  children.push(child);
  return child;
}

function shutdown(exitCode = 0) {
  while (children.length) {
    const child = children.pop();
    if (!child.killed) {
      child.kill('SIGTERM');
    }
  }

  process.exit(exitCode);
}

process.on('SIGINT', () => shutdown(0));
process.on('SIGTERM', () => shutdown(0));

console.log('Starting backend on http://localhost:3000 and frontend on http://127.0.0.1:5173');

startProcess('API server', 'node', ['server.js']);
startProcess('Vite dev server', 'npm', ['run', 'dev', '--', '--host', '127.0.0.1']);
