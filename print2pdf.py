from __future__ import annotations

import queue
import traceback
import threading
import multiprocessing
from dataclasses import dataclass
from pathlib import Path
from datetime import datetime
import re

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


VALID_EXTS = {".hwp", ".hwpx", ".hml"}


@dataclass
class ConvertConfig:
    root_dir: str
    output_dir: str
    include_hwp: bool
    include_hwpx: bool
    include_hml: bool
    include_name_contains: str
    exclude_name_contains: str
    save_with_timestamp: bool
    skip_if_same_stem_pdf_exists: bool


def make_allowed_exts(cfg: ConvertConfig) -> set[str]:
    exts: set[str] = set()
    if cfg.include_hwp:
        exts.add(".hwp")
    if cfg.include_hwpx:
        exts.add(".hwpx")
    if cfg.include_hml:
        exts.add(".hml")
    return exts


def _safe_folder_name(text: str) -> str:
    text = text.strip()
    if not text:
        return "필터링"

    text = re.sub(r"\s+", "_", text)
    text = re.sub(r'[\\/:*?"<>|]', "", text)
    text = text.strip(" .")

    if not text:
        return "필터링"

    return text[:120]


def _tokenize_filter_expr(expr: str) -> list[str]:
    """
    예:
    '중간 and (공통 or 수학)'
    -> ['중간', 'and', '(', '공통', 'or', '수학', ')']
    """
    tokens: list[str] = []
    i = 0
    n = len(expr)

    while i < n:
        ch = expr[i]

        if ch.isspace():
            i += 1
            continue

        if ch in "()":
            tokens.append(ch)
            i += 1
            continue

        j = i
        while j < n and expr[j] not in "()":
            j += 1

        chunk = expr[i:j].strip()
        if chunk:
            parts = re.split(r"(?i)\b(and|or)\b", chunk)
            for part in parts:
                part = part.strip()
                if part:
                    tokens.append(part)

        i = j

    return tokens


class _FilterExprParser:
    def __init__(self, tokens: list[str], filename: str):
        self.tokens = tokens
        self.filename = filename.lower()
        self.pos = 0

    def current(self) -> str | None:
        if self.pos >= len(self.tokens):
            return None
        return self.tokens[self.pos]

    def consume(self) -> str:
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def parse(self) -> bool:
        if not self.tokens:
            return True

        value = self.parse_or()

        if self.current() is not None:
            raise ValueError(f"잘못된 토큰: {self.tokens[self.pos:]}")

        return value

    def parse_or(self) -> bool:
        value = self.parse_and()

        while True:
            tok = self.current()

            if tok is not None and tok.lower() == "or":
                self.consume()
                rhs = self.parse_and()   # 반드시 파싱해서 토큰 소비
                value = value or rhs
            else:
                break

        return value

    def parse_and(self) -> bool:
        value = self.parse_factor()

        while True:
            tok = self.current()

            if tok is not None and tok.lower() == "and":
                self.consume()
                rhs = self.parse_factor()   # 반드시 파싱해서 토큰 소비
                value = value and rhs
            else:
                break

        return value

    def parse_factor(self) -> bool:
        tok = self.current()

        if tok is None:
            raise ValueError("필터식이 비정상적으로 끝났습니다.")

        if tok == "(":
            self.consume()

            value = self.parse_or()

            if self.current() != ")":
                raise ValueError("닫는 괄호 ')'가 필요합니다.")

            self.consume()
            return value

        if tok == ")":
            raise ValueError("여는 괄호 '(' 없이 닫는 괄호 ')'가 나왔습니다.")

        term = self.consume().lower()
        return term in self.filename


def _match_filter_expr(filename: str, expr: str) -> bool:
    expr = expr.strip()
    if not expr:
        return True

    tokens = _tokenize_filter_expr(expr)
    parser = _FilterExprParser(tokens, filename)
    return parser.parse()


def make_pdf_path(
    src_path: str,
    save_with_timestamp: bool,
    output_dir: str = "",
    filter_text: str = "",
) -> str:
    src = Path(src_path)

    if save_with_timestamp:
        now = datetime.now().strftime("%Y%m%d-%H%M%S-%f")[:-3]
        filename = f"{src.stem}-{now}.pdf"
    else:
        filename = f"{src.stem}.pdf"

    if output_dir.strip():
        out_dir = Path(output_dir).expanduser().resolve()
    else:
        out_dir = (src.parent / "pdf" / _safe_folder_name(filter_text)).resolve()

    out_dir.mkdir(parents=True, exist_ok=True)
    return str((out_dir / filename).resolve())


def find_target_files(cfg: ConvertConfig) -> list[str]:
    root = Path(cfg.root_dir)
    if not root.exists():
        raise FileNotFoundError(f"폴더를 찾을 수 없습니다: {cfg.root_dir}")

    allowed_exts = make_allowed_exts(cfg)
    include_text = cfg.include_name_contains.strip()
    exclude_text = cfg.exclude_name_contains.strip()

    files: list[str] = []

    for p in root.rglob("*"):
        if not p.is_file():
            continue

        if p.suffix.lower() not in allowed_exts:
            continue

        name_lower = p.name.lower()

        if include_text and not _match_filter_expr(name_lower, include_text):
            continue

        if exclude_text and _match_filter_expr(name_lower, exclude_text):
            continue

        if cfg.skip_if_same_stem_pdf_exists:
            if cfg.output_dir.strip():
                parent = Path(cfg.output_dir).expanduser().resolve()
            else:
                parent = (p.parent / "pdf" / _safe_folder_name(include_text)).resolve()

            stem = p.stem
            if parent.exists() and any(parent.glob(f"{stem}*.pdf")):
                continue

        files.append(str(p.resolve()))

    return files


def convert_one_fresh_instance(
    src_path: str,
    save_with_timestamp: bool,
    output_dir: str = "",
    filter_text: str = "",
) -> tuple[str, bool, str]:
    hwp = None

    try:
        from pyhwpx import Hwp

        src = Path(src_path).resolve()
        pdf_path = make_pdf_path(
            str(src),
            save_with_timestamp,
            output_dir,
            filter_text,
        )

        hwp = Hwp(visible=False)
        hwp.Open(str(src))
        hwp.SaveAs(pdf_path, "PDF")

        return (str(src), True, pdf_path)

    except Exception as e:
        err = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
        return (src_path, False, err)

    finally:
        try:
            if hwp is not None:
                hwp.Clear()
        except Exception:
            pass

        try:
            if hwp is not None:
                hwp.Quit()
        except Exception:
            pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP/HWPX/HML → PDF 변환기")
        self.geometry("920x700")
        self.minsize(840, 640)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False
        self.stop_requested = False
        self.worker_thread: threading.Thread | None = None

        self.root_dir_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.include_hwp_var = tk.BooleanVar(value=True)
        self.include_hwpx_var = tk.BooleanVar(value=True)
        self.include_hml_var = tk.BooleanVar(value=True)
        self.include_name_var = tk.StringVar()
        self.exclude_name_var = tk.StringVar()
        self.save_timestamp_var = tk.BooleanVar(value=False)
        self.skip_existing_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.after(120, self._drain_log_queue)

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=12)

        ttk.Label(top, text="대상 폴더").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.root_dir_var).grid(row=0, column=1, sticky="ew", padx=(8, 8))
        ttk.Button(top, text="찾아보기", command=self.pick_folder).grid(row=0, column=2)

        ttk.Label(top, text="저장 폴더").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.output_dir_var).grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(8, 0))
        ttk.Button(top, text="찾아보기", command=self.pick_output_folder).grid(row=1, column=2, pady=(8, 0))

        top.columnconfigure(1, weight=1)

        opt = ttk.LabelFrame(self, text="옵션")
        opt.pack(fill="x", padx=12, pady=(0, 12))

        ttk.Checkbutton(opt, text=".hwp", variable=self.include_hwp_var).grid(row=0, column=0, sticky="w", **pad)
        ttk.Checkbutton(opt, text=".hwpx", variable=self.include_hwpx_var).grid(row=0, column=1, sticky="w", **pad)
        ttk.Checkbutton(opt, text=".hml", variable=self.include_hml_var).grid(row=0, column=2, sticky="w", **pad)

        ttk.Label(opt, text="파일명 포함 필터 (and/or, 괄호 지원)").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(opt, textvariable=self.include_name_var).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        ttk.Label(opt, text="파일명 제외 필터 (and/or, 괄호 지원)").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(opt, textvariable=self.exclude_name_var).grid(row=2, column=1, columnspan=2, sticky="ew", **pad)

        ttk.Checkbutton(opt, text="파일명 뒤에 날짜시간 붙이기", variable=self.save_timestamp_var).grid(
            row=3, column=0, columnspan=2, sticky="w", **pad
        )
        ttk.Checkbutton(opt, text="같은 이름 계열 PDF가 이미 있으면 건너뛰기", variable=self.skip_existing_var).grid(
            row=4, column=0, columnspan=2, sticky="w", **pad
        )

        opt.columnconfigure(1, weight=1)
        opt.columnconfigure(2, weight=1)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=(0, 8))
        self.scan_btn = ttk.Button(btns, text="대상 미리보기", command=self.preview_targets)
        self.scan_btn.pack(side="left")
        self.start_btn = ttk.Button(btns, text="변환 시작", command=self.start_conversion)
        self.start_btn.pack(side="left", padx=(8, 0))
        self.stop_btn = ttk.Button(btns, text="중단", command=self.request_stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))
        self.clear_btn = ttk.Button(btns, text="로그 지우기", command=self.clear_log)
        self.clear_btn.pack(side="right")

        self.status_var = tk.StringVar(value="대기 중")
        ttk.Label(self, textvariable=self.status_var).pack(fill="x", padx=12, pady=(0, 4))
        ttk.Label(
            self,
            text='필터 예: "수학 and 중간", "부산 or 울산", "고1 and 1학기", "중간 and (공통 or 수학)"',
        ).pack(fill="x", padx=12, pady=(0, 6))

        log_frame = ttk.LabelFrame(self, text="로그")
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.log_text = tk.Text(log_frame, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.pack(side="right", fill="y")

        self.log_text.configure(yscrollcommand=scroll.set)

    def pick_folder(self):
        path = filedialog.askdirectory(title="변환할 폴더 선택")
        if path:
            self.root_dir_var.set(path)

    def pick_output_folder(self):
        path = filedialog.askdirectory(title="PDF 저장 폴더 선택")
        if path:
            self.output_dir_var.set(path)

    def clear_log(self):
        self.log_text.delete("1.0", tk.END)

    def log(self, message: str):
        self.log_queue.put(message)

    def _drain_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass

        self.after(120, self._drain_log_queue)

    def get_config(self) -> ConvertConfig:
        root_dir = self.root_dir_var.get().strip()
        if not root_dir:
            raise ValueError("대상 폴더를 선택해 주세요.")

        cfg = ConvertConfig(
            root_dir=root_dir,
            output_dir=self.output_dir_var.get().strip(),
            include_hwp=self.include_hwp_var.get(),
            include_hwpx=self.include_hwpx_var.get(),
            include_hml=self.include_hml_var.get(),
            include_name_contains=self.include_name_var.get(),
            exclude_name_contains=self.exclude_name_var.get(),
            save_with_timestamp=self.save_timestamp_var.get(),
            skip_if_same_stem_pdf_exists=self.skip_existing_var.get(),
        )

        if not make_allowed_exts(cfg):
            raise ValueError("최소 1개 이상의 확장자를 선택해 주세요.")

        return cfg

    def preview_targets(self):
        try:
            cfg = self.get_config()
            files = find_target_files(cfg)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return

        self.clear_log()
        self.log(f"대상 파일 수: {len(files)}")
        for p in files[:300]:
            self.log(p)
        if len(files) > 300:
            self.log(f"... 외 {len(files) - 300}개")

        self.status_var.set(f"대상 미리보기 완료: {len(files)}개")

    def start_conversion(self):
        if self.running:
            messagebox.showinfo("안내", "이미 변환 중입니다.")
            return

        try:
            cfg = self.get_config()
            files = find_target_files(cfg)
        except Exception as e:
            messagebox.showerror("오류", str(e))
            return

        if not files:
            messagebox.showinfo("안내", "조건에 맞는 파일이 없습니다.")
            return

        self.running = True
        self.stop_requested = False
        self.start_btn.configure(state="disabled")
        self.scan_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_var.set(f"변환 준비 중: {len(files)}개")
        self.clear_log()

        self.worker_thread = threading.Thread(
            target=self._run_conversion,
            args=(cfg, files),
            daemon=True,
        )
        self.worker_thread.start()

    def request_stop(self):
        if not self.running:
            return

        self.stop_requested = True
        self.stop_btn.configure(state="disabled")
        self.status_var.set("중단 요청됨: 현재 파일 처리 후 멈춤")
        self.log("[안내] 중단 요청됨 - 현재 파일 처리 후 작업을 멈춥니다.")

    def _run_conversion(self, cfg: ConvertConfig, files: list[str]):
        success_count = 0
        fail_count = 0
        total = len(files)

        self.log(f"대상 파일 수: {total}")
        self.log("변환 방식: 순차 처리 / 파일마다 새 한글 인스턴스")

        if cfg.output_dir:
            self.log(f"저장 폴더: {cfg.output_dir}")
        else:
            self.log(f"저장 폴더: 각 원본 폴더 아래의 pdf\\{_safe_folder_name(cfg.include_name_contains)}")

        try:
            for idx, src_path in enumerate(files, start=1):
                if self.stop_requested:
                    self.log("[중단] 사용자 요청으로 작업을 중단했습니다.")
                    self.status_var.set(f"중단됨 | 성공 {success_count} | 실패 {fail_count}")
                    break

                self.log(f"[{idx}/{total}] 변환 시작: {src_path}")

                src, ok, msg = convert_one_fresh_instance(
                    src_path,
                    cfg.save_with_timestamp,
                    cfg.output_dir,
                    cfg.include_name_contains,
                )

                if ok:
                    success_count += 1
                    self.log(f"[성공] {src} -> {msg}")
                else:
                    fail_count += 1
                    self.log(f"[실패] {src}")
                    self.log(msg)

                self.status_var.set(
                    f"진행 중: {idx}/{total} | 성공 {success_count} | 실패 {fail_count}"
                )

            else:
                self.status_var.set(f"완료 | 성공 {success_count} | 실패 {fail_count}")

            self.log("-" * 60)
            self.log(f"총 파일 수 : {total}")
            self.log(f"성공      : {success_count}")
            self.log(f"실패      : {fail_count}")

        except Exception as e:
            self.log("[치명적 오류]")
            self.log(f"{type(e).__name__}: {e}")
            self.log(traceback.format_exc())
            self.status_var.set("오류로 중단됨")

        finally:
            self.running = False
            self.start_btn.configure(state="normal")
            self.scan_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")


def main():
    multiprocessing.freeze_support()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()