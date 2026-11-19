// OpenVBS++ VS Code integration.
//
// Deliberately dependency-free: it shells out to the `openvbs` CLI and consumes the
// machine-readable JSON that `run --json` and `check --json` already emit. A future LSP
// server (see doc/IDE_LSP_DESIGN.md) can replace this file without touching the grammar.
"use strict";

const { execFile } = require("node:child_process");
const path = require("node:path");
const vscode = require("vscode");

const MAX_BUFFER = 8 * 1024 * 1024;

function configuration() {
  const cfg = vscode.workspace.getConfiguration("openvbs");
  const executable = cfg.get("executable", "openvbs");
  const extra = cfg.get("arguments", []);
  return {
    executable: executable,
    extra: Array.isArray(extra) ? extra : [],
    checkOnSave: cfg.get("checkOnSave", true)
  };
}

function projectCwd(document) {
  const folder = vscode.workspace.getWorkspaceFolder(document.uri);
  if (folder) {
    return folder.uri.fsPath;
  }
  return path.dirname(document.uri.fsPath);
}

function invoke(document, args) {
  const settings = configuration();
  const argv = [...settings.extra, ...args, document.uri.fsPath];
  return new Promise((resolve) => {
    execFile(
      settings.executable,
      argv,
      { cwd: projectCwd(document), maxBuffer: MAX_BUFFER, timeout: 60000, windowsHide: true },
      (error, stdout, stderr) => {
        resolve({ error, stdout: stdout || "", stderr: stderr || "" });
      }
    );
  });
}

function parseJson(text) {
  const trimmed = text.trim();
  if (!trimmed) {
    return null;
  }
  // Take the last non-empty line that parses: the CLI may print warnings first.
  const lines = trimmed.split(/\r?\n/).reverse();
  for (const line of lines) {
    const candidate = line.trim();
    if (!candidate.startsWith("{")) {
      continue;
    }
    try {
      return JSON.parse(candidate);
    } catch (error) {
      // keep looking
    }
  }
  return null;
}

function offsetOf(column) {
  const zeroBased = Number.isFinite(column) ? column - 1 : 0;
  return zeroBased > 0 ? zeroBased : 0;
}

function toDiagnostics(document, payload) {
  const result = [];
  const items = payload && Array.isArray(payload.diagnostics) ? payload.diagnostics : [];
  for (const item of items) {
    const line = Number.isFinite(item.line) ? item.line - 1 : 0;
    const clampedLine = Math.max(0, Math.min(line, document.lineCount - 1));
    const textLine = document.lineAt(clampedLine).text;
    const start = Math.min(offsetOf(item.column), textLine.length);
    const range = new vscode.Range(clampedLine, start, clampedLine, textLine.length);
    const message = item.message || "OpenVBS++ reported an error";
    const diagnostic = new vscode.Diagnostic(range, message, vscode.DiagnosticSeverity.Error);
    diagnostic.source = "openvbs";
    diagnostic.code = item.kind || (item.number === null ? undefined : String(item.number));
    result.push(diagnostic);
  }
  return result;
}

function activate(context) {
  const output = vscode.window.createOutputChannel("OpenVBS++");
  const diagnostics = vscode.languages.createDiagnosticCollection("openvbspp");

  const reportFailure = (document, response, action) => {
    const detail = response.error && response.error.message ? response.error.message : response.stderr.trim();
    if (response.error && response.error.code === "ENOENT") {
      vscode.window.showErrorMessage(
        `OpenVBS++: 找不到命令 "${configuration().executable}"。请在设置里配置 openvbs.executable（或设为 python 并把 openvbs.arguments 设为 ["-m","OpenVBSPP"]）。`
      );
      return;
    }
    vscode.window.showErrorMessage(`OpenVBS++: ${action} 失败 — ${detail.split(/\r?\n/)[0]}`);
    output.appendLine(detail);
    output.show(true);
  };

  const check = async (document, silent) => {
    const response = await invoke(document, ["check", "--json"]);
    const payload = parseJson(response.stdout);
    if (!payload) {
      if (!silent) {
        reportFailure(document, response, "编译检查");
      } else {
        diagnostics.delete(document.uri);
      }
      return false;
    }
    diagnostics.set(document.uri, toDiagnostics(document, payload));
    if (!payload.ok && !silent) {
      output.appendLine(`check: ${payload.diagnostics.length} 个问题`);
      for (const item of payload.diagnostics) {
        output.appendLine(`  line ${item.line}: ${item.message}`);
      }
      output.show(true);
    }
    return Boolean(payload.ok);
  };

  const run = async (document) => {
    const ok = await check(document, true);
    if (!ok) {
      output.appendLine("run: 已跳过执行（编译检查未通过）");
      output.show(true);
      return;
    }
    const response = await invoke(document, ["run", "--json"]);
    const payload = parseJson(response.stdout);
    output.show(true);
    if (!payload) {
      reportFailure(document, response, "运行");
      return;
    }
    output.appendLine(`$ openvbs run ${path.basename(document.uri.fsPath)}  (${payload.duration_ms ?? "?"} ms)`);
    if (payload.output) {
      output.append(payload.output.endsWith("\n") ? payload.output : `${payload.output}\n`);
    }
    if (!payload.ok) {
      output.appendLine(`[失败] ${payload.error || "unknown error"}`);
    }
    output.appendLine("");
  };

  context.subscriptions.push(
    output,
    diagnostics,
    vscode.commands.registerCommand("openvbs.runFile", async () => {
      const document = vscode.window.activeTextEditor && vscode.window.activeTextEditor.document;
      if (document) {
        await run(document);
      }
    }),
    vscode.commands.registerCommand("openvbs.checkFile", async () => {
      const document = vscode.window.activeTextEditor && vscode.window.activeTextEditor.document;
      if (document) {
        await check(document, false);
      }
    }),
    vscode.workspace.onDidSaveTextDocument(async (document) => {
      if (document.languageId !== "openvbs") {
        return;
      }
      if (!configuration().checkOnSave) {
        return;
      }
      await check(document, true);
    }),
    vscode.workspace.onDidOpenTextDocument(async (document) => {
      if (document.languageId === "openvbs" && configuration().checkOnSave) {
        await check(document, true);
      }
    })
  );
}

function deactivate() {}

module.exports = { activate, deactivate };
