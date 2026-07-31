/**
 * Pull Execution Logs locally from Google Apps Script / Google Cloud Logging
 * Zero CPU overhead on Apps Script execution.
 */

const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const LOGS_DIR = path.join(__dirname, 'logs');
const LOG_FILE = path.join(LOGS_DIR, `execution_${new Date().toISOString().replace(/[:.]/g, '-')}.log`);
const LATEST_LOG_FILE = path.join(LOGS_DIR, 'latest_execution.log');

if (!fs.existsSync(LOGS_DIR)) {
  fs.mkdirSync(LOGS_DIR, { recursive: true });
}

console.log('📡 Fetching execution logs from Google Apps Script...');

try {
  const nodeExe = `C:\\Users\\ben.arthur\\node-v24.14.1-win-x64\\node.exe`;
  const claspJs = `C:\\Users\\ben.arthur\\node-v24.14.1-win-x64\\node_modules\\@google\\clasp\\build\\src\\index.js`;

  const result = spawnSync(nodeExe, [claspJs, 'logs', '--json'], { encoding: 'utf8' });

  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || `Exit code ${result.status}`);
  }

  const output = result.stdout;

  fs.writeFileSync(LOG_FILE, output);
  fs.writeFileSync(LATEST_LOG_FILE, output);

  console.log(`✅ Logs saved successfully!`);
  console.log(`📁 Historical Log: ${LOG_FILE}`);
  console.log(`📁 Latest Log:     ${LATEST_LOG_FILE}`);

  // Display summary of latest lines to terminal
  const lines = output.trim().split('\n');
  console.log('\n--- Recent Log Activity (Last 15 Entries) ---');
  lines.slice(-15).forEach(line => console.log(line));
} catch (error) {
  console.error('❌ Error pulling logs:', error.message);
  if (error.stdout) console.log(error.stdout);
  if (error.stderr) console.error(error.stderr);
}
