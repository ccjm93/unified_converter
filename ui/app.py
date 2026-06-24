"""통합 변환기 메인 윈도우 (단일 화면 + 출력형식 다중선택)."""
from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import theme as T
from .widgets import make_button
from ..core import Orchestrator, supported_inputs
from ..core.engines import DEFAULT_FACTORIES
from ..utils.filescan import INPUT_EXTS, scan_folder, split_name_ext
from ..utils.logging_setup import setup_file_logging

# 드래그&드롭(tkinterdnd2)은 선택적 의존성. 없으면 DnD 없이 정상 동작한다
# (한컴/Excel 미설치 시 해당 변환만 비활성화되는 것과 동일한 graceful degradation).
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _DND_OK = True
    _TkBase = TkinterDnD.Tk
except Exception:  # 라이브러리·네이티브 tkdnd 부재 등 모든 실패를 포괄
    _DND_OK = False
    _TkBase = tk.Tk
    DND_FILES = None


def _progid_available(progid: str) -> bool:
    """레지스트리로 COM ProgID 설치 여부를 확인(앱을 실제로 띄우지 않음)."""
    try:
        import winreg
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid).Close()
        return True
    except OSError:
        return False
    except Exception:
        return False


# 출력형식 → (라벨, 필요 엔진 ProgID). 표시 순서: HWPX → XLSX → PDF.
_OUTPUT_FORMATS = [
    (".hwpx", "HWPX"),
    (".xlsx", "XLSX"),
    (".pdf", "PDF"),
]

# 변환 후 원본 처리 콤보박스 라벨
_KEEP_ORIGINAL = "원본 유지"
_DELETE_ORIGINAL = "원본 삭제"


class App(_TkBase):
    def __init__(self) -> None:
        super().__init__()
        self.title("통합 파일형식 변환기")
        self.geometry("760x680")
        self.minsize(680, 560)
        self.configure(bg=T.BG)

        self._files: list[str] = []
        self._out_dir: str = ""
        self._running = False
        self._item_status: dict[int, str] = {}  # index → 'ok'|'fail'|'skip'

        self._logger = setup_file_logging()
        self._hancom_ok = _progid_available("HWPFrame.HwpObject")
        self._excel_ok = _progid_available("Excel.Application")

        T.apply_ttk(self)
        self._build_menu()
        self._build_ui()
        self._announce_environment()

    # ── 메뉴 (부가 도구) ─────────────────────────────────────────
    def _build_menu(self) -> None:
        menubar = tk.Menu(self)
        tools = tk.Menu(menubar, tearoff=0)
        tools.add_command(label="Excel 외부 링크 일괄 해제...", command=self._open_link_breaker)
        menubar.add_cascade(label="도구", menu=tools)
        self.config(menu=menubar)

    # ── UI 빌드 ──────────────────────────────────────────────────
    def _build_ui(self) -> None:
        PAD = 16
        # 헤더
        header = tk.Frame(self, bg=T.ACCENT, pady=12)
        header.pack(fill="x")
        tk.Label(header, text="통합 파일형식 변환기", font=T.FONT_H,
                 bg=T.ACCENT, fg=T.CRUST).pack()

        # 파일 영역
        file_wrap = tk.Frame(self, bg=T.BG, padx=PAD, pady=10)
        file_wrap.pack(fill="both", expand=True)
        title = "변환할 파일" + ("  (여기로 파일·폴더를 끌어다 놓을 수 있어요)" if _DND_OK else "")
        tk.Label(file_wrap, text=title, font=T.FONT_B, bg=T.BG, fg=T.TEXT).pack(anchor="w")

        btn_row = tk.Frame(file_wrap, bg=T.BG)
        btn_row.pack(fill="x", pady=(4, 6))
        make_button(btn_row, "📄 파일 추가", self._add_files, bg=T.BLUE).pack(side="left", padx=(0, 6))
        make_button(btn_row, "📁 폴더 추가", self._add_folder, bg=T.BLUE).pack(side="left", padx=(0, 6))
        make_button(btn_row, "초기화", self._clear_files, bg=T.SURFACE2, fg=T.TEXT,
                    hover=T.SURFACE).pack(side="right")

        list_frame = tk.Frame(file_wrap, bg=T.SURFACE)
        list_frame.pack(fill="both", expand=True)
        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            list_frame, yscrollcommand=sb.set, bg=T.SURFACE, fg=T.TEXT,
            selectbackground=T.ACCENT, selectforeground=T.CRUST, font=T.FONT_N,
            bd=0, highlightthickness=0, activestyle="none", relief="flat",
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.config(command=self.listbox.yview)
        self._enable_dnd(self.listbox)

        self.count_var = tk.StringVar(value="선택된 파일: 0개")
        tk.Label(file_wrap, textvariable=self.count_var, font=T.FONT_N,
                 bg=T.BG, fg=T.SUBTEXT).pack(anchor="e", pady=(4, 0))

        # 출력 형식 (다중선택)
        opt_frame = tk.Frame(self, bg=T.BG, padx=PAD, pady=4)
        opt_frame.pack(fill="x")
        tk.Label(opt_frame, text="출력 형식 (복수 선택 가능)", font=T.FONT_B,
                 bg=T.BG, fg=T.TEXT).pack(anchor="w")
        cb_row = tk.Frame(opt_frame, bg=T.BG)
        cb_row.pack(fill="x", pady=(4, 0))
        self._fmt_vars: dict[str, tk.BooleanVar] = {}
        for ext, label in _OUTPUT_FORMATS:
            var = tk.BooleanVar(value=(ext in (".hwpx", ".xlsx")))
            self._fmt_vars[ext] = var
            tk.Checkbutton(
                cb_row, text=label, variable=var, font=T.FONT_N,
                bg=T.BG, fg=T.TEXT, selectcolor=T.SURFACE,
                activebackground=T.BG, activeforeground=T.TEXT,
                highlightthickness=0, bd=0,
            ).pack(side="left", padx=(0, 16))

        # 변환 후 원본 처리 (기본: 유지) — 데이터 손실 방지를 위해 명시적 opt-in.
        self._delete_var = tk.StringVar(value=_KEEP_ORIGINAL)
        ttk.Combobox(
            cb_row, textvariable=self._delete_var, state="readonly", width=9,
            values=(_KEEP_ORIGINAL, _DELETE_ORIGINAL),
        ).pack(side="right")
        tk.Label(cb_row, text="변환 후:", font=T.FONT_N,
                 bg=T.BG, fg=T.SUBTEXT).pack(side="right", padx=(0, 6))

        # 출력 폴더
        out_frame = tk.Frame(self, bg=T.BG, padx=PAD, pady=4)
        out_frame.pack(fill="x")
        tk.Label(out_frame, text="저장 폴더", font=T.FONT_B, bg=T.BG, fg=T.TEXT).pack(side="left", padx=(0, 8))
        self.out_var = tk.StringVar(value="(원본과 같은 폴더)")
        self.out_entry = tk.Entry(out_frame, textvariable=self.out_var, font=T.FONT_N,
                                  bg=T.SURFACE, fg=T.TEXT, relief="flat", bd=4,
                                  readonlybackground=T.SURFACE, state="readonly")
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self._refresh_out_style()
        make_button(out_frame, "선택", self._choose_out, bg=T.SURFACE2, fg=T.TEXT,
                    hover=T.SURFACE).pack(side="left", padx=(0, 4))
        make_button(out_frame, "초기화", self._reset_out, bg=T.SURFACE2, fg=T.TEXT,
                    hover=T.SURFACE).pack(side="left")

        # 진행/상태
        prog_frame = tk.Frame(self, bg=T.BG, padx=PAD, pady=6)
        prog_frame.pack(fill="x")
        self.status_var = tk.StringVar(value="파일을 추가하고 변환을 시작하세요.")
        tk.Label(prog_frame, textvariable=self.status_var, font=T.FONT_N,
                 bg=T.BG, fg=T.SUBTEXT).pack(anchor="w", pady=(0, 4))
        self.progress = ttk.Progressbar(prog_frame, style="Accent.Horizontal.TProgressbar",
                                        mode="determinate")
        self.progress.pack(fill="x")

        # 변환 버튼
        self.convert_btn = make_button(
            self, "⚡ 변환 시작", self._start, font=("Segoe UI", 11, "bold"), pady=10)
        self.convert_btn.pack(fill="x", padx=PAD, pady=(6, 4))

        # 로그
        log_frame = tk.Frame(self, bg=T.BG, padx=PAD)
        log_frame.pack(fill="both", expand=True, pady=(0, PAD))
        tk.Label(log_frame, text="로그", font=T.FONT_B, bg=T.BG, fg=T.TEXT).pack(anchor="w")
        self.log_text = tk.Text(log_frame, font=T.MONO, bg=T.SURFACE, fg=T.TEXT,
                                relief="flat", bd=6, wrap="word", state="disabled", height=8)
        self.log_text.pack(fill="both", expand=True, pady=(4, 0))
        for tag, color in (("info", T.TEXT), ("success", T.GREEN),
                           ("warn", T.YELLOW), ("error", T.RED)):
            self.log_text.tag_configure(tag, foreground=color)

    # ── 환경 안내 (graceful degradation) ─────────────────────────
    def _announce_environment(self) -> None:
        if not self._hancom_ok:
            self._log("⚠ 한컴오피스(한글)가 감지되지 않음 — HWP/HWPX 변환 불가", "warn")
        if not self._excel_ok:
            self._log("⚠ Microsoft Excel이 감지되지 않음 — XLS/XLSX 변환 불가", "warn")
        if self._hancom_ok and self._excel_ok:
            self._log("✓ 한컴오피스 · Excel 감지됨 — 모든 변환 사용 가능", "success")

    # ── 파일 선택 ────────────────────────────────────────────────
    def _add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="변환할 파일 선택",
            filetypes=[("지원 문서", "*.hwp *.hwpx *.xls *.xlsx"), ("모든 파일", "*.*")],
        )
        self._add_paths(paths)

    def _add_folder(self) -> None:
        folder = filedialog.askdirectory(title="폴더 선택")
        if not folder:
            return
        paths = scan_folder(folder, INPUT_EXTS)
        if not paths:
            messagebox.showinfo("알림", "선택한 폴더에 지원 파일이 없습니다.")
            return
        self._add_paths(paths)

    def _add_paths(self, paths) -> None:
        for p in paths:
            if p not in self._files:
                self._files.append(p)
                self.listbox.insert("end", p)
        self.count_var.set(f"선택된 파일: {len(self._files)}개")

    # ── 드래그 & 드롭 ────────────────────────────────────────────
    def _enable_dnd(self, widget) -> None:
        """위젯을 파일 드롭 타깃으로 등록(tkinterdnd2 없으면 무동작)."""
        if not _DND_OK:
            return
        try:
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_drop)
        except Exception:
            pass  # 등록 실패해도 버튼 추가는 그대로 사용 가능

    @staticmethod
    def _parse_drop_data(data: str) -> list[str]:
        """tkdnd 드롭 문자열을 경로 목록으로 분해.

        공백 포함 경로는 '{...}'로 감싸여 온다. tk.splitlist는 Windows 경로의
        백슬래시(\\new, \\tab 등)를 Tcl 이스케이프로 오인하므로, 백슬래시를
        건드리지 않는 정규식으로 직접 토큰을 추출한다.
        """
        import re
        return [a or b for a, b in re.findall(r"\{([^}]*)\}|(\S+)", data)]

    def _on_drop(self, event) -> None:
        raw = self._parse_drop_data(event.data)
        paths: list[str] = []
        skipped = 0
        for p in raw:
            if os.path.isdir(p):
                paths.extend(scan_folder(p, INPUT_EXTS))
            elif split_name_ext(p)[1] in INPUT_EXTS:
                paths.append(p)
            else:
                skipped += 1
        if paths:
            self._add_paths(paths)
        if skipped:
            self._log(f"⚠ 지원하지 않는 항목 {skipped}개는 제외했습니다 "
                      "(hwp·hwpx·xls·xlsx만 지원).", "warn")

    def _clear_files(self) -> None:
        self._files.clear()
        self.listbox.delete(0, "end")
        self._item_status.clear()
        self.count_var.set("선택된 파일: 0개")

    def _choose_out(self) -> None:
        folder = filedialog.askdirectory(title="저장 폴더 선택")
        if folder:
            self._out_dir = folder
            self.out_var.set(folder)
            self._refresh_out_style()

    def _reset_out(self) -> None:
        self._out_dir = ""
        self.out_var.set("(원본과 같은 폴더)")
        self._refresh_out_style()

    def _refresh_out_style(self) -> None:
        """저장 폴더가 미지정(placeholder)이면 눈에 띄게 빨간 굵은 글씨,
        실제 폴더가 지정되면 일반 텍스트 색으로 표시한다."""
        if self._out_dir:
            self.out_entry.config(fg=T.TEXT, font=T.FONT_N)
        else:
            self.out_entry.config(fg=T.RED, font=("Segoe UI Light", T.FONT_N[1]))

    # ── 변환 시작 ────────────────────────────────────────────────
    def _selected_targets(self) -> list[str]:
        return [ext for ext, _ in _OUTPUT_FORMATS if self._fmt_vars[ext].get()]

    def _start(self) -> None:
        if self._running:
            return
        if not self._files:
            messagebox.showwarning("경고", "변환할 파일을 먼저 추가하세요.")
            return
        targets = self._selected_targets()
        if not targets:
            messagebox.showwarning("경고", "출력 형식을 하나 이상 선택하세요.")
            return

        delete_originals = self._delete_var.get() == _DELETE_ORIGINAL
        if delete_originals and not messagebox.askyesno(
            "원본 삭제 확인",
            "변환에 성공한 파일의 원본(hwp·hwpx·xls·xlsx)을 삭제합니다.\n"
            "출력형식이 하나라도 실패한 파일의 원본은 보존됩니다.\n\n계속할까요?",
        ):
            return

        self._running = True
        self._item_status.clear()
        self.convert_btn.config(state="disabled", text="변환 중...")
        self.progress["value"] = 0
        self._clear_log()
        self.status_var.set("변환 중...")

        orch = Orchestrator(DEFAULT_FACTORIES)
        files = list(self._files)
        out_dir = self._out_dir or None
        threading.Thread(
            target=self._run_worker,
            args=(orch, files, targets, out_dir, delete_originals), daemon=True,
        ).start()

    def _run_worker(self, orch, files, targets, out_dir, delete_originals) -> None:
        orch.run(
            files, targets, out_dir=out_dir,
            on_log=self._log, on_progress=self._on_progress,
            on_item_result=self._on_item, on_done=self._on_done,
            init_com=True, delete_originals=delete_originals,
        )

    # ── 콜백 (메인 스레드로 마샬링) ───────────────────────────────
    def _log(self, message: str, level: str = "info") -> None:
        self._logger.info("[%s] %s", level, message)

        def _write():
            self.log_text.config(state="normal")
            self.log_text.insert("end", message + "\n", level)
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, _write)

    def _on_progress(self, done: int, total: int) -> None:
        def _u():
            self.progress["maximum"] = max(total, 1)
            self.progress["value"] = done
            self.status_var.set(f"변환 중... {done}/{total}")
        self.after(0, _u)

    def _on_item(self, index: int, result) -> None:
        # 같은 파일에 여러 출력형식 → 하나라도 실패하면 빨강, 건너뜀은 회색, 전부 성공 초록
        prev = self._item_status.get(index)
        if not result.ok and result.error == "건너뜀":
            status = "skip" if prev is None else prev
        elif not result.ok:
            status = "fail"
        else:
            status = "ok" if prev in (None, "ok") else prev
        self._item_status[index] = status
        color = {"ok": T.GREEN, "fail": T.RED, "skip": T.SUBTEXT}[status]
        self.after(0, lambda: self.listbox.itemconfig(index, fg=color))

    def _on_done(self, summary) -> None:
        def _finish():
            self._running = False
            self.convert_btn.config(state="normal", text="⚡ 변환 시작")
            s, f, sk = summary.success, len(summary.failed), len(summary.skipped)
            d = len(summary.deleted)
            del_msg = f" / 원본삭제 {d}" if d else ""
            self._log(f"\n🎉 완료 — 성공 {s} / 실패 {f} / 건너뜀 {sk}{del_msg}", "success")
            self.status_var.set(f"완료 — 성공 {s} · 실패 {f} · 건너뜀 {sk}{del_msg}")
            if f == 0:
                messagebox.showinfo("완료", f"변환 완료!\n성공 {s} / 건너뜀 {sk}{del_msg}")
            else:
                detail = "\n".join(f"• {os.path.basename(p)}: {e}" for p, e in summary.failed[:10])
                messagebox.showwarning("완료(일부 실패)",
                                       f"성공 {s} / 실패 {f} / 건너뜀 {sk}{del_msg}\n\n{detail}")
        self.after(0, _finish)

    def _clear_log(self) -> None:
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")

    # ── 부가 도구: 링크 해제 ─────────────────────────────────────
    def _open_link_breaker(self) -> None:
        if not self._excel_ok:
            messagebox.showwarning("Excel 필요", "Microsoft Excel이 설치되어 있지 않습니다.")
            return
        folder = filedialog.askdirectory(title="링크를 해제할 엑셀 폴더 선택")
        if not folder:
            return
        if not messagebox.askyesno(
            "확인", "원본은 .bak으로 백업된 뒤 링크가 제거됩니다.\n계속할까요?"):
            return
        self._log(f"\n[도구] Excel 링크 해제 시작: {folder}", "info")

        def _worker():
            from ..tools.link_breaker import break_links_in_folder
            break_links_in_folder(folder, on_log=self._log, backup=True)
        threading.Thread(target=_worker, daemon=True).start()
