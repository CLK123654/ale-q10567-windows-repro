import crypto from 'node:crypto';
import fs from 'node:fs';
import fsp from 'node:fs/promises';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { spawn } from 'node:child_process';
import { inflateRawSync } from 'node:zlib';

const repo = path.resolve(import.meta.dirname, '..');
const artifacts = path.join(repo, 'artifacts');
const evidence = path.join(repo, 'verification', 'evidence');
const names = ['输入数据包.zip', 'reference.zip', '关键标准答案.xlsx', '任务规格转化.xlsx'];
const delivery = [
  'output/chart/billing-service/Chart.yaml',
  'output/chart/billing-service/templates/_helpers.tpl',
  'output/chart/billing-service/templates/configmap.yaml',
  'output/chart/billing-service/templates/deployment.yaml',
  'output/chart/billing-service/templates/service.yaml',
  'output/chart/billing-service/templates/serviceaccount.yaml',
  'output/chart/billing-service/values.yaml',
  'output/chart/billing-service/values/env/dev.yaml',
  'output/chart/billing-service/values/env/prod.yaml',
  'output/rendered/dev.yaml',
  'output/rendered/prod.yaml',
  'output/reports/render_diff.csv',
].sort();
const assert = (value, message) => { if (!value) throw new Error(message); };
const sha = (value) => crypto.createHash('sha256').update(value).digest('hex');
const shaFile = (file) => sha(fs.readFileSync(file));

function zip(file) {
  const data = fs.readFileSync(file);
  let end = -1;
  for (let index = data.length - 22; index >= Math.max(0, data.length - 65557); index -= 1) if (data.readUInt32LE(index) === 0x06054b50) { end = index; break; }
  assert(end >= 0, 'ZIP目录缺失');
  const count = data.readUInt16LE(end + 10);
  let offset = data.readUInt32LE(end + 16);
  const output = new Map();
  for (let index = 0; index < count; index += 1) {
    assert(data.readUInt32LE(offset) === 0x02014b50, 'ZIP目录损坏');
    const method = data.readUInt16LE(offset + 10);
    const size = data.readUInt32LE(offset + 20);
    const plain = data.readUInt32LE(offset + 24);
    const nameLength = data.readUInt16LE(offset + 28);
    const extraLength = data.readUInt16LE(offset + 30);
    const commentLength = data.readUInt16LE(offset + 32);
    const local = data.readUInt32LE(offset + 42);
    const name = data.subarray(offset + 46, offset + 46 + nameLength).toString('utf8').replaceAll('\\', '/');
    const localNameLength = data.readUInt16LE(local + 26);
    const localExtraLength = data.readUInt16LE(local + 28);
    const start = local + 30 + localNameLength + localExtraLength;
    if (!name.endsWith('/')) {
      const raw = data.subarray(start, start + size);
      const body = method === 0 ? raw : method === 8 ? inflateRawSync(raw) : null;
      assert(body && body.length === plain, `ZIP成员损坏${name}`);
      output.set(name, body);
    }
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return output;
}

async function extract(file, destination) {
  for (const [name, bytes] of zip(file)) {
    const target = path.resolve(destination, ...name.split('/'));
    assert(target.startsWith(path.resolve(destination) + path.sep), 'ZIP路径越界');
    await fsp.mkdir(path.dirname(target), { recursive: true });
    await fsp.writeFile(target, bytes);
  }
}

function sheets(file) {
  const xml = zip(file).get('xl/workbook.xml')?.toString('utf8') ?? '';
  return [...xml.matchAll(/<(?:[A-Za-z]+:)?sheet[^>]+name="([^"]+)"/gu)].map((match) => match[1]);
}

async function run(command, args, cwd, env = {}) {
  const start = Date.now();
  return new Promise((resolve) => {
    const child = spawn(command, args, { cwd, env: { ...process.env, ...env }, windowsHide: true });
    let stdout = '';
    let stderr = '';
    let done = false;
    child.stdout.on('data', (value) => stdout += value);
    child.stderr.on('data', (value) => stderr += value);
    child.on('error', (error) => { if (!done) { done = true; resolve({ code: 1, stdout, stderr: stderr + error.message, elapsed_ms: Date.now() - start }); } });
    child.on('exit', (code) => { if (!done) { done = true; resolve({ code: code ?? 1, stdout, stderr, elapsed_ms: Date.now() - start }); } });
  });
}

async function download(url, file) {
  await new Promise((resolve, reject) => {
    const get = (current) => https.get(current, (response) => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) { response.resume(); get(new URL(response.headers.location, current)); return; }
      if (response.statusCode !== 200) { reject(new Error(`下载失败${response.statusCode}`)); return; }
      const output = fs.createWriteStream(file);
      response.pipe(output);
      output.on('finish', () => output.close(resolve));
      output.on('error', reject);
    }).on('error', reject);
    get(new URL(url));
  });
}

function tree(root) {
  const lines = [];
  function walk(directory, prefix = '') {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true }).toSorted((left, right) => left.name.localeCompare(right.name))) {
      const relative = prefix ? `${prefix}/${entry.name}` : entry.name;
      const full = path.join(directory, entry.name);
      if (entry.isDirectory()) walk(full, relative);
      else lines.push(`${relative}\0${shaFile(full)}`);
    }
  }
  walk(root);
  return sha(Buffer.from(lines.join('\n')));
}

function listed(root) {
  const output = [];
  function walk(directory, prefix = 'output') {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const relative = `${prefix}/${entry.name}`;
      const full = path.join(directory, entry.name);
      if (entry.isDirectory()) walk(full, relative);
      else output.push(relative.replaceAll('\\', '/'));
    }
  }
  walk(root);
  return output.sort();
}

function compare(output, standard, strict = true) {
  assert(JSON.stringify(listed(output)) === JSON.stringify(delivery), '交付成员错误');
  const digest = crypto.createHash('sha256');
  for (const name of delivery) {
    const actual = fs.readFileSync(path.join(output, name.slice(7))).toString('utf8').replaceAll('\r\n', '\n').trimEnd();
    const expected = standard.get(name).toString('utf8').replaceAll('\r\n', '\n').trimEnd();
    if (strict) assert(actual === expected, `交付内容不一致${name}`);
    digest.update(actual);
  }
  return digest.digest('hex');
}

const tool = path.join(os.tmpdir(), 'Q10567 Helm Tool');
await fsp.rm(tool, { recursive: true, force: true });
await fsp.mkdir(tool, { recursive: true });
const archive = path.join(tool, 'helm.zip');
await download('https://get.helm.sh/helm-v3.17.3-windows-amd64.zip', archive);
assert(shaFile(archive) === '8ea93e2f6285e649dede583ac90ff8cdb938ca53ec6cf5fe909f2303fbc22d96', 'Helm包校验错误');
await extract(archive, tool);
const helm = path.join(tool, 'windows-amd64', 'helm.exe');
assert(fs.existsSync(helm), 'helm.exe缺失');

await fsp.rm(evidence, { recursive: true, force: true });
await fsp.mkdir(evidence, { recursive: true });
assert(process.platform === 'win32' && process.env.GITHUB_ACTIONS === 'true', '需要GitHub托管Windows');
const attachmentSha = Object.fromEntries(names.map((name) => [name, shaFile(path.join(artifacts, name))]));
const input = zip(path.join(artifacts, '输入数据包.zip'));
const standard = zip(path.join(artifacts, 'reference.zip'));
assert(JSON.stringify([...standard.keys()].sort()) === JSON.stringify(delivery), 'Reference成员错误');
const platform = [...input].filter(([name, bytes]) => (bytes[0] === 0x7f && bytes.subarray(1, 4).toString() === 'ELF') || /\.(?:sh|bash|so)$/iu.test(name));
assert(platform.length === 0, '输入含平台专用成员');
assert(JSON.stringify(sheets(path.join(artifacts, '关键标准答案.xlsx'))) === JSON.stringify(['交付物答案清单', '固定字段答案', '固定集合答案', '固定数值答案', '允许变体答案']), '标答Sheet错误');
assert(JSON.stringify(sheets(path.join(artifacts, '任务规格转化.xlsx'))) === JSON.stringify(['任务规格转化']), '规格Sheet错误');

async function prepare(label, mutate) {
  const root = path.join(os.tmpdir(), label);
  await fsp.rm(root, { recursive: true, force: true });
  await fsp.mkdir(root, { recursive: true });
  await extract(path.join(artifacts, '输入数据包.zip'), root);
  for (const name of delivery.filter((item) => item.startsWith('output/chart/'))) {
    const file = path.join(root, name);
    await fsp.mkdir(path.dirname(file), { recursive: true });
    await fsp.writeFile(file, standard.get(name));
  }
  const inputRoot = path.join(root, 'input_data');
  if (mutate) await mutate(inputRoot, root);
  return { root, inputRoot, output: path.join(root, 'output') };
}
const execute = (inputRoot) => run('node', ['tools/run-task.mjs'], inputRoot, { ALE_HELM_BIN: helm });

const clean = [];
for (const label of ['Q10567 第一 中文 空目录', 'Q10567 第二 中文 空格目录']) {
  const prepared = await prepare(label);
  const before = tree(prepared.inputRoot);
  const result = await execute(prepared.inputRoot);
  assert(result.code === 0, `正式运行失败${result.stdout}${result.stderr}`);
  const after = tree(prepared.inputRoot);
  assert(before === after, '输入被修改');
  const digest = compare(prepared.output, standard);
  clean.push({ directory_label: label, exit_code: 0, input_digest_before: before, input_digest_after: after, semantic_digest: digest, elapsed_ms: result.elapsed_ms });
}
assert(clean[0].semantic_digest === clean[1].semantic_digest, '双目录结果不同');

const crlf = await prepare('Q10567 CRLF 合同', async (inputRoot) => {
  const file = path.join(inputRoot, 'rules', 'billing_render_contract.yaml');
  const text = await fsp.readFile(file, 'utf8');
  await fsp.writeFile(file, text.replace(/\r?\n/gu, '\r\n'));
});
let result = await execute(crlf.inputRoot);
assert(result.code === 0, `CRLF运行失败${result.stderr}`);
assert(compare(crlf.output, standard) === clean[0].semantic_digest, 'CRLF改变结果');

const changed = await prepare('Q10567 生产副本变化', async (_inputRoot, root) => {
  const file = path.join(root, 'output', 'chart', 'billing-service', 'values', 'env', 'prod.yaml');
  const text = await fsp.readFile(file, 'utf8');
  await fsp.writeFile(file, text.replace('replicaCount: 6', 'replicaCount: 8'));
});
result = await execute(changed.inputRoot);
assert(result.code === 0, `副本变化运行失败${result.stderr}`);
compare(changed.output, standard, false);
const changedRender = await fsp.readFile(path.join(changed.output, 'rendered', 'prod.yaml'), 'utf8');
const changedDiff = await fsp.readFile(path.join(changed.output, 'reports', 'render_diff.csv'), 'utf8');
assert(changedRender.includes('replicas: 8') && changedDiff.includes('replica_count,2,8,'), '副本变化未联动清单与报告');

const invalid = await prepare('Q10567 缺少配置模板', async (_inputRoot, root) => {
  await fsp.rm(path.join(root, 'output', 'chart', 'billing-service', 'templates', 'configmap.yaml'));
});
result = await execute(invalid.inputRoot);
assert(result.code !== 0, '缺少模板被接受');
assert(!fs.existsSync(path.join(invalid.output, 'reports', 'render_diff.csv')), '无效Chart仍生成差异报告');

const version = await run(helm, ['version', '--short'], repo);
const proof = {
  schema_version: 1,
  task_asset_id: 'helm_billing_environment_release',
  result: 'PASS',
  generated_at_utc: new Date().toISOString(),
  git_commit_sha: process.env.GITHUB_SHA,
  workflow_run_id: process.env.GITHUB_RUN_ID,
  runner: { os: process.env.RUNNER_OS, arch: process.env.RUNNER_ARCH, image_os: process.env.ImageOS, image_version: process.env.ImageVersion, node: process.version, actual_windows_run: true },
  software: { main: 'Helm', version: version.stdout.trim(), archive_sha256: shaFile(archive), executed: true },
  attachment_sha256: attachmentSha,
  reference_members: delivery,
  workbook_checks: { answer_sheet_names: sheets(path.join(artifacts, '关键标准答案.xlsx')), specification_sheet_names: ['任务规格转化'] },
  platform_audit: { platform_specific_members: platform.map((item) => item[0]), no_wsl_required: true, no_linux_container_required: true },
  clean_runs: clean,
  crlf_input: { exit_code: 0, reference_match: true },
  positive_mutation: { changed_rule: 'prod replicaCount', new_value: 8, rendered_replicas: 8, diff_report_value: 8 },
  invalid_input: { changed_rule: 'missing configmap template', exit_code: result.code, report_absent: true },
  network: { installation: 'official Helm release archive', formal_run: 'offline' },
};
await fsp.writeFile(path.join(evidence, 'windows-verification.json'), `${JSON.stringify(proof, null, 2)}\n`);
console.log(JSON.stringify(proof, null, 2));
