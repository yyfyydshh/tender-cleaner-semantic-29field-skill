#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""八爪鱼任务采集 + 招投标29字段清洗的统一入口。"""

from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import ssl
import sys
import time
import unicodedata
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from tender_clean import (
    FIELDS,
    clean_records,
    deduplicate_records,
    load_records,
    write_results,
)


DEFAULT_MCP_URL = "https://mcp.bazhuayu.com"
PROTOCOL_VERSION = "2025-03-26"
API_KEY_ENV = "BAZHUAYU_API_KEY"


class WorkflowError(RuntimeError):
    """可直接展示给用户的工作流错误。"""


def normalize_name(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "")
    return re.sub(r"\s+", " ", value).strip()


def safe_filename(value: str, fallback: str = "task") -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", normalize_name(value))
    return value.strip(". ") or fallback


def validate_api_key(value: str | None) -> str:
    key = (value or "").strip()
    if not key:
        raise WorkflowError(
            f"未提供八爪鱼 API Key。请设置环境变量 {API_KEY_ENV}；"
            "不要把密钥写入 SKILL.md、脚本或输出文件。"
        )
    lowered = key.lower()
    if "<api" in lowered or "api_key" in lowered or key.startswith("<") or key.endswith(">"):
        raise WorkflowError("检测到 API Key 占位符，而不是真实密钥。请提供实际密钥。")
    if len(key) < 20:
        raise WorkflowError("API Key 长度异常，请确认复制完整。")
    return key


def parse_mcp_body(raw: bytes) -> dict[str, Any] | None:
    if not raw:
        return None
    text = raw.decode("utf-8-sig", errors="strict").strip()
    if not text:
        return None
    data_lines = [line[5:].strip() for line in text.splitlines() if line.startswith("data:")]
    payload = data_lines[-1] if data_lines else text
    return json.loads(payload)


class BazhuayuMCP:
    """标准 Streamable HTTP MCP 客户端；使用 x-api-key 鉴权。"""

    def __init__(self, api_key: str, url: str = DEFAULT_MCP_URL, timeout: int = 90):
        self.api_key = validate_api_key(api_key)
        self.url = url.rstrip("/")
        self.timeout = timeout
        self.session_id: str | None = None
        self.request_id = 0
        self.ssl_context = ssl.create_default_context()

    def _next_id(self) -> int:
        self.request_id += 1
        return self.request_id

    def _headers(self) -> dict[str, str]:
        headers = {
            "Accept": "application/json, text/event-stream",
            "Content-Type": "application/json",
            "x-api-key": self.api_key,
        }
        if self.session_id:
            headers["Mcp-Session-Id"] = self.session_id
            headers["MCP-Protocol-Version"] = PROTOCOL_VERSION
        return headers

    def _post(self, payload: dict[str, Any]) -> dict[str, Any] | None:
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        request = urllib.request.Request(
            self.url,
            data=body,
            headers=self._headers(),
            method="POST",
        )
        try:
            with urllib.request.urlopen(
                request,
                timeout=self.timeout,
                context=self.ssl_context,
            ) as response:
                session_id = response.headers.get("Mcp-Session-Id")
                if session_id:
                    self.session_id = session_id
                return parse_mcp_body(response.read())
        except urllib.error.HTTPError as exc:
            detail = exc.read(1200).decode("utf-8", errors="replace")
            detail = detail.replace(self.api_key, "<redacted>")
            if exc.code in (401, 403):
                raise WorkflowError(
                    f"八爪鱼 MCP 鉴权失败（HTTP {exc.code}）。请检查 API Key 是否有效。"
                ) from exc
            raise WorkflowError(f"八爪鱼 MCP HTTP {exc.code}: {detail[:500]}") from exc
        except urllib.error.URLError as exc:
            raise WorkflowError(f"无法连接八爪鱼 MCP：{exc.reason}") from exc
        except TimeoutError as exc:
            raise WorkflowError("连接八爪鱼 MCP 超时。") from exc

    def initialize(self) -> dict[str, Any]:
        response = self._post(
            {
                "jsonrpc": "2.0",
                "id": self._next_id(),
                "method": "initialize",
                "params": {
                    "protocolVersion": PROTOCOL_VERSION,
                    "capabilities": {},
                    "clientInfo": {"name": "tender-cleaner", "version": "3.0"},
                },
            }
        )
        self._raise_rpc_error(response)
        if not self.session_id:
            raise WorkflowError("MCP 初始化成功但未返回 Mcp-Session-Id。")
        self._post({"jsonrpc": "2.0", "method": "notifications/initialized", "params": {}})
        return (response or {}).get("result", {})

    @staticmethod
    def _raise_rpc_error(response: dict[str, Any] | None) -> None:
        if response and response.get("error"):
            error = response["error"]
            raise WorkflowError(f"MCP RPC {error.get('code')}: {error.get('message')}")

    def call_tool(self, name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        response = self._post(
            {
                "jsonrpc": "2.0",
                "id": self._next_id(),
                "method": "tools/call",
                "params": {"name": name, "arguments": arguments},
            }
        )
        self._raise_rpc_error(response)
        result = (response or {}).get("result", {})
        if result.get("isError"):
            detail = self._content_text(result) or "未知工具错误"
            raise WorkflowError(f"八爪鱼工具 {name} 调用失败：{detail}")
        structured = result.get("structuredContent")
        if isinstance(structured, dict):
            return structured
        text = self._content_text(result)
        if not text:
            return {}
        try:
            parsed = json.loads(text)
            return parsed if isinstance(parsed, dict) else {"data": parsed}
        except json.JSONDecodeError:
            return {"message": text}

    @staticmethod
    def _content_text(result: dict[str, Any]) -> str:
        for item in result.get("content", []):
            if isinstance(item, dict) and item.get("type") == "text":
                return str(item.get("text", ""))
        return ""

    def search_all_tasks(self, page_size: int = 50) -> list[dict[str, Any]]:
        """不传中文 keyword，规避服务端中文关键词误判；客户端本地匹配。"""
        first = self.call_tool("search_tasks", {"page": 1, "size": page_size})
        tasks = list(first.get("tasks") or [])
        total_pages = int(first.get("totalPages") or 1)
        for page in range(2, total_pages + 1):
            data = self.call_tool("search_tasks", {"page": page, "size": page_size})
            tasks.extend(data.get("tasks") or [])
        by_id = {str(task.get("taskId")): task for task in tasks if task.get("taskId")}
        return list(by_id.values())

    def search_tasks_by_ids(self, task_ids: Iterable[str]) -> list[dict[str, Any]]:
        ids = list(dict.fromkeys(str(task_id) for task_id in task_ids))
        if not ids:
            return []
        data = self.call_tool(
            "search_tasks",
            {"taskIds": ids, "page": 1, "size": max(10, len(ids))},
        )
        return list(data.get("tasks") or [])


def task_label(task: dict[str, Any]) -> str:
    return str(task.get("taskStatusLabel") or task.get("status") or "未知")


def is_running(task: dict[str, Any]) -> bool:
    return task.get("rawTaskStatusCode") == 0 or task_label(task).lower() in {"运行中", "running"}


def is_completed(task: dict[str, Any]) -> bool:
    label = task_label(task).lower()
    return task.get("rawTaskStatusCode") == 2 or label in {"已完成", "completed", "finished"}


def is_failed(task: dict[str, Any]) -> bool:
    label = task_label(task).lower()
    return any(token in label for token in ("失败", "failed", "error"))


def candidate_tasks(query: str, tasks: list[dict[str, Any]], limit: int = 8) -> list[dict[str, Any]]:
    normalized = normalize_name(query)
    contains = [task for task in tasks if normalized in normalize_name(str(task.get("taskName", "")))]
    if contains:
        return contains[:limit]
    names = [normalize_name(str(task.get("taskName", ""))) for task in tasks]
    close = set(difflib.get_close_matches(normalized, names, n=limit, cutoff=0.35))
    return [task for task in tasks if normalize_name(str(task.get("taskName", ""))) in close][:limit]


def resolve_tasks(
    requested_names: list[str],
    requested_ids: list[str],
    tasks: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    errors: list[str] = []
    resolved: list[dict[str, Any]] = []

    for query in requested_names:
        exact = [
            task
            for task in tasks
            if normalize_name(str(task.get("taskName", ""))) == normalize_name(query)
        ]
        if len(exact) == 1:
            resolved.append(exact[0])
            continue
        if len(exact) > 1:
            options = ", ".join(str(task.get("taskId")) for task in exact)
            errors.append(f"任务名“{query}”有多个精确匹配，请改用 --task-id：{options}")
            continue
        candidates = candidate_tasks(query, tasks)
        preview = "；".join(
            f"{task.get('taskName')} ({task.get('taskId')})" for task in candidates
        ) or "无候选"
        errors.append(f"未找到任务名“{query}”的精确匹配。候选：{preview}")

    by_id = {str(task.get("taskId")): task for task in tasks}
    for task_id in requested_ids:
        task = by_id.get(str(task_id))
        if task:
            resolved.append(task)
        else:
            errors.append(f"账户中不存在 taskId：{task_id}")

    if errors:
        raise WorkflowError("任务解析失败，尚未启动任何任务：\n- " + "\n- ".join(errors))
    unique = {str(task["taskId"]): task for task in resolved}
    if not unique:
        raise WorkflowError("至少提供一个 --task-name 或 --task-id。")
    return list(unique.values())


def start_tasks(client: BazhuayuMCP, tasks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    states: list[dict[str, Any]] = []
    for task in tasks:
        state = {
            "taskId": str(task["taskId"]),
            "taskName": str(task.get("taskName") or task["taskId"]),
            "previousStatus": task_label(task),
            "startStatus": "already_running" if is_running(task) else "pending",
            "finalStatus": "",
            "error": "",
        }
        try:
            if not is_running(task):
                response = client.call_tool(
                    "start_or_stop_task",
                    {"taskId": state["taskId"], "action": "start"},
                )
                if response.get("success") is False:
                    raise WorkflowError(str(response.get("message") or "启动请求被拒绝"))
                state["startStatus"] = str(response.get("status") or "start_requested")
            print(f"已提交: {state['taskName']} | {state['taskId']} | {state['startStatus']}")
        except Exception as exc:  # 保留其他任务的成功结果
            state["startStatus"] = "failed"
            state["finalStatus"] = "start_failed"
            state["error"] = str(exc)
            print(f"启动失败: {state['taskName']} | {exc}", file=sys.stderr)
        states.append(state)
    return states


def poll_tasks(
    client: BazhuayuMCP,
    states: list[dict[str, Any]],
    timeout_seconds: int,
    interval_seconds: int,
) -> None:
    pending = {
        state["taskId"]: state
        for state in states
        if state["finalStatus"] != "start_failed"
    }
    if not pending:
        return
    deadline = time.monotonic() + timeout_seconds
    time.sleep(min(3, interval_seconds))
    last_report = 0.0

    while pending and time.monotonic() < deadline:
        found = {task["taskId"]: task for task in client.search_tasks_by_ids(pending)}
        now = time.monotonic()
        for task_id, state in list(pending.items()):
            task = found.get(task_id)
            if not task:
                continue
            state["lastStatus"] = task_label(task)
            if is_completed(task):
                state["finalStatus"] = "completed"
                pending.pop(task_id, None)
            elif is_failed(task):
                state["finalStatus"] = "failed"
                state["error"] = task_label(task)
                pending.pop(task_id, None)
        if now - last_report >= 30 or not pending:
            summary = ", ".join(
                f"{state['taskName']}={state.get('lastStatus') or state['finalStatus'] or '等待中'}"
                for state in states
            )
            print(f"任务状态: {summary}")
            last_report = now
        if pending:
            time.sleep(interval_seconds)

    for state in pending.values():
        state["finalStatus"] = "timeout"
        state["error"] = f"超过 {timeout_seconds} 秒仍未完成"


def download_file(url: str, destination: Path, timeout: int = 180) -> None:
    request = urllib.request.Request(url, headers={"User-Agent": "tender-cleaner/3.0"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            destination.write_bytes(response.read())
    except (urllib.error.URLError, TimeoutError) as exc:
        raise WorkflowError(f"下载八爪鱼导出文件失败：{exc}") from exc


def export_task(
    client: BazhuayuMCP,
    state: dict[str, Any],
    raw_dir: Path,
    timeout_seconds: int,
    interval_seconds: int,
) -> Path | None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        data = client.call_tool(
            "export_data",
            {
                "taskId": state["taskId"],
                "exportFileType": "JSON",
                "previewRows": 5,
            },
        )
        status = str(data.get("status") or "").lower()
        state["exportStatus"] = status or "unknown"
        state["rawRows"] = int(data.get("dataTotal") or 0)
        state["exportFileUrl"] = str(data.get("exportFileUrl") or data.get("exportUrl") or "")

        if status in {"exported", "completed"} or state["exportFileUrl"]:
            output = raw_dir / f"{safe_filename(state['taskName'])}_raw.json"
            if state["exportFileUrl"]:
                download_file(state["exportFileUrl"], output)
                return output
            sample = data.get("sampleData") or []
            if sample and state["rawRows"] <= len(sample):
                output.write_text(json.dumps(sample, ensure_ascii=False, indent=2), encoding="utf-8")
                return output
            raise WorkflowError("导出已完成但未返回完整文件链接，不能用预览数据冒充全量数据。")
        if status == "no_data":
            state["error"] = "任务完成但没有可导出数据"
            return None
        if status in {"failed", "error"}:
            raise WorkflowError(str(data.get("message") or "导出失败"))
        time.sleep(interval_seconds)
    raise WorkflowError(f"导出超过 {timeout_seconds} 秒仍未完成")


def clean_inputs(
    input_paths: list[Path],
    output_dir: Path,
    output_name: str,
    sheet_name: str | None,
    keep_duplicates: bool,
    verbose: bool,
) -> dict[str, Any]:
    records, source_counts = load_records([str(path) for path in input_paths], sheet_name)
    if not records:
        raise WorkflowError("输入文件已读取，但没有可清洗的数据行。")
    input_rows = len(records)
    unique_records = records if keep_duplicates else deduplicate_records(records)
    results = clean_records(unique_records, verbose=verbose)
    csv_path, json_path = write_results(results, output_dir, output_name)
    non_empty = sum(1 for row in results for value in row.values() if value)
    total_cells = len(results) * len(FIELDS)
    return {
        "inputs": source_counts,
        "inputRows": input_rows,
        "outputRows": len(results),
        "deduplicated": input_rows - len(results),
        "fieldCoverage": round(non_empty / total_cells * 100, 2) if total_cells else 0,
        "csv": str(csv_path),
        "json": str(json_path),
    }


def write_manifest(path: Path, manifest: dict[str, Any]) -> None:
    path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def run_collect(args: argparse.Namespace) -> int:
    api_key = validate_api_key(args.api_key or os.environ.get(API_KEY_ENV))
    output_dir = Path(args.output_dir).expanduser().resolve()
    raw_dir = output_dir / "raw"
    clean_dir = output_dir / "clean"
    raw_dir.mkdir(parents=True, exist_ok=True)
    clean_dir.mkdir(parents=True, exist_ok=True)
    manifest: dict[str, Any] = {
        "generatedAt": datetime.now().astimezone().isoformat(),
        "mcpUrl": args.mcp_url,
        "tasks": [],
        "outputs": {},
    }
    manifest_path = output_dir / "manifest.json"

    client = BazhuayuMCP(api_key, args.mcp_url, args.http_timeout)
    info = client.initialize()
    server = info.get("serverInfo") or {}
    print(f"MCP 已连接: {server.get('name', 'bazhuayu')} {server.get('version', '')}".strip())

    all_tasks = client.search_all_tasks()
    tasks = resolve_tasks(args.task_name or [], args.task_id or [], all_tasks)
    print(f"已唯一解析 {len(tasks)} 个任务；账户任务总数 {len(all_tasks)}。")
    if args.resolve_only:
        manifest["tasks"] = tasks
        write_manifest(manifest_path, manifest)
        for task in tasks:
            print(f"{task.get('taskName')} | {task.get('taskId')} | {task_label(task)}")
        return 0

    states = start_tasks(client, tasks)
    manifest["tasks"] = states
    write_manifest(manifest_path, manifest)
    poll_tasks(client, states, args.task_timeout, args.poll_interval)

    successful_raw: list[Path] = []
    for state in states:
        if state["finalStatus"] != "completed":
            continue
        try:
            raw_path = export_task(
                client,
                state,
                raw_dir,
                args.export_timeout,
                args.poll_interval,
            )
            if not raw_path:
                continue
            state["rawFile"] = str(raw_path)
            successful_raw.append(raw_path)
            individual = clean_inputs(
                [raw_path],
                clean_dir,
                f"{state['taskName']}_29字段",
                args.sheet,
                args.keep_duplicates,
                not args.quiet,
            )
            state["cleanOutput"] = individual
        except Exception as exc:
            state["error"] = str(exc)
            state["exportStatus"] = "failed"
            print(f"导出/清洗失败: {state['taskName']} | {exc}", file=sys.stderr)
        finally:
            write_manifest(manifest_path, manifest)

    extra_inputs = [Path(path).expanduser().resolve() for path in (args.input or [])]
    combined_inputs = successful_raw + extra_inputs
    if combined_inputs:
        combined_name = args.name or (
            "招标公告_合并清洗" if len(combined_inputs) > 1 else f"{combined_inputs[0].stem}_29字段"
        )
        manifest["outputs"]["combined"] = clean_inputs(
            combined_inputs,
            clean_dir,
            combined_name,
            args.sheet,
            args.keep_duplicates,
            not args.quiet,
        )
    write_manifest(manifest_path, manifest)

    failures = [state for state in states if state.get("error")]
    print(f"运行清单: {manifest_path}")
    if failures:
        print(f"部分完成：{len(states) - len(failures)}/{len(states)} 个任务成功。", file=sys.stderr)
        return 1
    print(f"全部完成：{len(states)} 个任务。")
    return 0


def run_files(args: argparse.Namespace) -> int:
    inputs = [Path(path).expanduser().resolve() for path in args.input]
    output_dir = Path(args.output_dir).expanduser().resolve()
    output_name = args.name or (
        f"{inputs[0].stem}_29字段" if len(inputs) == 1 else "招标公告_合并清洗"
    )
    result = clean_inputs(
        inputs,
        output_dir,
        output_name,
        args.sheet,
        args.keep_duplicates,
        not args.quiet,
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="八爪鱼现有任务采集 + 招投标29字段清洗；或直接清洗本地文件。"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    collect = subparsers.add_parser("collect", help="按现有任务名/ID启动、导出并清洗")
    collect.add_argument("--task-name", action="append", default=[], help="八爪鱼现有任务名，可重复")
    collect.add_argument("--task-id", action="append", default=[], help="八爪鱼 taskId，可重复")
    collect.add_argument("--input", nargs="*", default=[], help="可选：与采集结果一起合并清洗的本地文件")
    collect.add_argument("--api-key", help=f"可选；默认读取 {API_KEY_ENV}（环境变量更安全）")
    collect.add_argument("--mcp-url", default=DEFAULT_MCP_URL)
    collect.add_argument("--output-dir", default=f"tender-run-{datetime.now():%Y%m%d-%H%M%S}")
    collect.add_argument("--name", help="合并输出文件名，不含扩展名")
    collect.add_argument("--sheet", help="Excel 工作表名")
    collect.add_argument("--keep-duplicates", action="store_true")
    collect.add_argument("--resolve-only", action="store_true", help="只解析任务，不启动")
    collect.add_argument("--task-timeout", type=int, default=1800, help="等待任务完成秒数")
    collect.add_argument("--export-timeout", type=int, default=600, help="等待导出完成秒数")
    collect.add_argument("--poll-interval", type=int, default=10)
    collect.add_argument("--http-timeout", type=int, default=90)
    collect.add_argument("--quiet", action="store_true")
    collect.set_defaults(func=run_collect)

    files = subparsers.add_parser("files", help="直接清洗 JSON/CSV/TSV/Excel")
    files.add_argument("-i", "--input", nargs="+", required=True)
    files.add_argument("-o", "--output-dir", default=".")
    files.add_argument("--name", help="输出文件名，不含扩展名")
    files.add_argument("--sheet", help="Excel 工作表名")
    files.add_argument("--keep-duplicates", action="store_true")
    files.add_argument("--quiet", action="store_true")
    files.set_defaults(func=run_files)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        return int(args.func(args))
    except (WorkflowError, FileNotFoundError, ValueError, ImportError) as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("已取消。", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
