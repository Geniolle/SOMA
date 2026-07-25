from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

log_lock = threading.Lock()


def _utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S")


def _normalize_sensitive_url(raw: str) -> str:
    text = raw.strip()
    if not text:
        return text

    try:
        parts = urlsplit(text)
    except Exception:
        return text

    if not parts.scheme or not parts.netloc:
        return text

    sensitive_keys = {"token", "password", "passwd", "secret", "cookie", "session", "sid", "auth", "key", "user"}
    query_items = []
    for key, value in parse_qsl(parts.query, keep_blank_values=True):
        if key.lower() in sensitive_keys:
            query_items.append((key, "[redacted]"))
        else:
            query_items.append((key, value))
    return urlunsplit((parts.scheme, parts.netloc, parts.path, urlencode(query_items), parts.fragment))


def redact_line(text: str) -> str:
    line = text.rstrip("\n")
    if not line:
        return line

    line = re.sub(
        r"(?i)\b(?:password|passwd|token|cookie|cookies|secret|authorization|auth|session|sid|user|username)\b\s*[:=]\s*([^\s,;]+)",
        "[redacted]",
        line,
    )
    line = re.sub(r"https?://\S+", lambda m: _normalize_sensitive_url(m.group(0)), line)
    line = re.sub(r"(?i)(SITE_PASSWORD|SITE_USER|GOOGLE_CREDENTIALS_PATH)\s*=\s*[^ \t]+", r"\1=[redacted]", line)
    line = re.sub(r"(?i)\b(token|cookie|password|secret|authorization)\b[^ \t]*", "[redacted]", line)
    return line


@dataclass
class SmokeTestConfig:
    process_name: str = "Teste SOMA E2E"
    prompt_timeout_seconds: float = 180.0
    record_active_timeout_seconds: float = 90.0
    command_timeout_seconds: float = 60.0
    exit_timeout_seconds: float = 120.0
    line_poll_seconds: float = 0.1
    run_id_prefix: str = "teste-browser-e2e"
    site_mapper_args: list[str] = field(
        default_factory=lambda: ["--mode", "record", "--no-headless"]
    )


@dataclass
class SmokeTestResult:
    run_id: str
    status: str
    duration_seconds: float
    returncode: int | None
    artifacts_dir: Path
    stdout_lines: list[str]
    stderr_lines: list[str]
    commands_sent: list[str]
    validations: list[str]
    errors: list[str]


class RecordBrowserSmokeTest:
    def __init__(self, config: SmokeTestConfig | None = None) -> None:
        self.config = config or SmokeTestConfig()
        self.project_root = Path(__file__).resolve().parents[3]
        self.python_executable = sys.executable
        self.env_file = self.project_root / "deploy" / ".env"
        self.stdout_lines: list[str] = []
        self.stderr_lines: list[str] = []
        self.combined_lines: list[str] = []
        self.commands_sent: list[str] = []
        self.validations: list[str] = []
        self.errors: list[str] = []

    def _build_run_id(self) -> str:
        return f"{self.config.run_id_prefix}-{_utc_stamp()}"

    def _build_command(self, run_id: str) -> list[str]:
        return [
            self.python_executable,
            "-u",
            "-m",
            "soma_app.tools.site_mapper",
            *self.config.site_mapper_args,
            "--run-id",
            run_id,
        ]

    def _log_line(self, stream_name: str, line: str) -> None:
        redacted = redact_line(line)
        if not redacted:
            return
        with log_lock:
            print(f"[{stream_name}] {redacted}", flush=True)

    def _reader(self, stream, sink: list[str], stream_name: str, stop_event: threading.Event) -> None:
        try:
            for raw in iter(stream.readline, ""):
                if stop_event.is_set() and not raw:
                    break
                line = raw.rstrip("\n")
                redacted = redact_line(line)
                sink.append(redacted)
                self.combined_lines.append(redacted)
                self._log_line(stream_name, line)
        finally:
            try:
                stream.close()
            except Exception:
                pass

    def _wait_for(self, predicate: Callable[[str], bool], *, timeout_seconds: float, lines: list[str], label: str) -> None:
        deadline = time.monotonic() + timeout_seconds
        seen = 0
        while time.monotonic() < deadline:
            while seen < len(lines):
                current = lines[seen]
                seen += 1
                if predicate(current):
                    return
            time.sleep(self.config.line_poll_seconds)
        raise TimeoutError(f"Timeout a aguardar {label} ({timeout_seconds:.0f}s).")

    def _wait_for_any(self, predicates: list[tuple[str, Callable[[str], bool]]], *, timeout_seconds: float, lines: list[str]) -> str:
        deadline = time.monotonic() + timeout_seconds
        seen = 0
        while time.monotonic() < deadline:
            while seen < len(lines):
                current = lines[seen]
                seen += 1
                for label, predicate in predicates:
                    if predicate(current):
                        return label
            time.sleep(self.config.line_poll_seconds)
        labels = ", ".join(label for label, _ in predicates)
        raise TimeoutError(f"Timeout a aguardar um de: {labels}.")

    def _send_command(self, proc: subprocess.Popen[str], command: str) -> None:
        if proc.stdin is None:
            raise RuntimeError("stdin do subprocess indisponível.")
        proc.stdin.write(command + "\n")
        proc.stdin.flush()
        self.commands_sent.append(command)

    def _terminate_process(self, proc: subprocess.Popen[str]) -> None:
        try:
            if proc.poll() is not None:
                return
        except Exception:
            return

        if os.name == "nt":
            try:
                subprocess.run(
                    ["taskkill", "/PID", str(proc.pid), "/T", "/F"],
                    capture_output=True,
                    text=True,
                    timeout=20,
                )
                return
            except Exception:
                pass

        try:
            proc.terminate()
            proc.wait(timeout=10)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass

    def _assert_artifacts(self, run_id: str) -> Path:
        artifacts_dir = self.project_root / "artifacts" / "site-map" / run_id
        required = [
            "record_status.json",
            "steps.json",
            "workflow.json",
            "workflow_summary.json",
            "elements_used.json",
            "locator_candidates_used.json",
        ]
        missing = [name for name in required if not (artifacts_dir / name).exists()]
        if missing:
            raise AssertionError(f"Artefactos em falta: {', '.join(missing)}")

        status = json.loads((artifacts_dir / "record_status.json").read_text(encoding="utf-8"))
        if status.get("status") != "complete":
            raise AssertionError(f"record_status.json não indica complete: {status!r}")

        steps = json.loads((artifacts_dir / "steps.json").read_text(encoding="utf-8"))
        actions = [str(item.get("action", "")) for item in steps if isinstance(item, dict)]
        labels = [str(item.get("label", "")) for item in steps if isinstance(item, dict)]
        for expected_action in ("checkpoint", "marker"):
            if expected_action not in actions:
                raise AssertionError(f"steps.json não contém a ação esperada: {expected_action}")
        for expected_label in ("mark", "pause", "resume", "stop"):
            if expected_label not in labels:
                raise AssertionError(f"steps.json não contém o marcador esperado: {expected_label}")
        if labels.count("checkpoint") == 0:
            raise AssertionError("steps.json não contém checkpoint")

        for name in ("workflow.json", "workflow_summary.json", "elements_used.json", "locator_candidates_used.json"):
            text = (artifacts_dir / name).read_text(encoding="utf-8")
            if re.search(r"(?i)\b(password|token|cookie|authorization|secret)\b", text):
                raise AssertionError(f"Segredo em texto aberto encontrado em {name}")
            if "javascript" in text.lower() and "error" in text.lower():
                raise AssertionError(f"Erro JavaScript encontrado em {name}")

        steps_text = (artifacts_dir / "steps.json").read_text(encoding="utf-8")
        if "[redacted]" not in steps_text:
            raise AssertionError("Valores sensíveis não foram redatados em steps.json")

        return artifacts_dir

    def run(self) -> SmokeTestResult:
        run_id = self._build_run_id()
        command = self._build_command(run_id)
        env = os.environ.copy()
        if self.env_file.exists():
            env.setdefault("ENV_FILE", str(self.env_file))
        env["PYTHONUNBUFFERED"] = "1"

        start = time.monotonic()
        proc = subprocess.Popen(
            command,
            cwd=self.project_root,
            env=env,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
        )

        stop_event = threading.Event()
        stdout_thread = threading.Thread(target=self._reader, args=(proc.stdout, self.stdout_lines, "stdout", stop_event), daemon=True)
        stderr_thread = threading.Thread(target=self._reader, args=(proc.stderr, self.stderr_lines, "stderr", stop_event), daemon=True)
        stdout_thread.start()
        stderr_thread.start()

        try:
            self._wait_for(
                lambda line: "Página carregada com sucesso após clicar no botão 'SOMA'." in line,
                timeout_seconds=self.config.prompt_timeout_seconds,
                lines=self.combined_lines,
                label="bootstrap do SOMA",
            )
            self._send_command(proc, self.config.process_name)

            self._wait_for(
                lambda line: "Modo record ativo. Comandos:" in line,
                timeout_seconds=self.config.record_active_timeout_seconds,
                lines=self.combined_lines,
                label="mensagem do modo record",
            )
            self.validations.append("record_activated")

            command_plan = [
                ("mark", "Marcador registado."),
                ("pause", "Gravação pausada."),
                ("__smoke__:pause_menu", "Smoke action concluída: pause_menu."),
                ("resume", "Gravação retomada."),
                ("__smoke__:resume_new_form", "Smoke action concluída: resume_new_form."),
                ("", "Checkpoint manual registado."),
                ("__smoke__:final_note", "Smoke action concluída: final_note."),
                ("q", None),
            ]
            for command_text, ack in command_plan:
                self._send_command(proc, command_text)
                if ack:
                    self._wait_for(
                        lambda line, needle=ack: needle in line,
                        timeout_seconds=self.config.command_timeout_seconds,
                        lines=self.combined_lines,
                        label=f"ack de {command_text or 'Enter'}",
                    )
                    self.validations.append(command_text or "checkpoint")

            returncode = proc.wait(timeout=self.config.exit_timeout_seconds)
            self.validations.append("process_exited")

            artifacts_dir = self._assert_artifacts(run_id)
            self.validations.append("artifacts_validated")

            if any("Traceback" in line for line in self.stderr_lines + self.stdout_lines):
                raise AssertionError("Traceback encontrado na saída do processo.")

            if any("JavaScript" in line and "Error" in line for line in self.stdout_lines + self.stderr_lines):
                raise AssertionError("Erro JavaScript encontrado na saída do processo.")

            duration = time.monotonic() - start
            return SmokeTestResult(
                run_id=run_id,
                status="complete",
                duration_seconds=duration,
                returncode=returncode,
                artifacts_dir=artifacts_dir,
                stdout_lines=list(self.stdout_lines),
                stderr_lines=list(self.stderr_lines),
                commands_sent=list(self.commands_sent),
                validations=list(self.validations),
                errors=list(self.errors),
            )
        except Exception as exc:
            self.errors.append(str(exc))
            self._terminate_process(proc)
            raise
        finally:
            stop_event.set()
            try:
                if proc.stdin is not None:
                    proc.stdin.close()
            except Exception:
                pass
            self._terminate_process(proc)
            stdout_thread.join(timeout=5)
            stderr_thread.join(timeout=5)


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Smoke test automático do modo record do SOMA.")
    parser.add_argument("--process-name", default="Teste SOMA E2E")
    parser.add_argument("--run-id", default="")
    parser.add_argument("--prompt-timeout", type=float, default=180.0)
    parser.add_argument("--record-active-timeout", type=float, default=90.0)
    parser.add_argument("--command-timeout", type=float, default=60.0)
    parser.add_argument("--exit-timeout", type=float, default=120.0)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    smoke = RecordBrowserSmokeTest(
        SmokeTestConfig(
            process_name=args.process_name,
            prompt_timeout_seconds=args.prompt_timeout,
            record_active_timeout_seconds=args.record_active_timeout,
            command_timeout_seconds=args.command_timeout,
            exit_timeout_seconds=args.exit_timeout,
        )
    )
    if args.run_id:
        smoke._build_run_id = lambda: args.run_id  # type: ignore[method-assign]

    result = smoke.run()
    print(
        json.dumps(
            {
                "status": result.status,
                "run_id": result.run_id,
                "duration_seconds": round(result.duration_seconds, 2),
                "returncode": result.returncode,
                "artifacts_dir": str(result.artifacts_dir),
                "commands_sent": result.commands_sent,
                "validations": result.validations,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
