# ui_gui.py
# -*- coding: utf-8 -*-

from __future__ import annotations
import os
import sys
import threading
import traceback
from pathlib import Path
from datetime import date
import re

import tkinter as tk
from tkinter import filedialog, messagebox

# Theming
import ttkbootstrap as tb
from ttkbootstrap.constants import *

# Drag & Drop (opcjonalnie)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_OK = True
except Exception:
    DND_OK = False

# Core
from core.utils import base_dir, wybierz_sciezke_wyjsciowa, znajdz_plik_operations, znajdz_plik_loyalty
from core.io_operations import wczytaj_operations, wczytaj_operations_many
from core.io_loyalty import wczytaj_loyalty, wczytaj_loyalty_many
from core.compare import porownaj
from core.report import zapisz_do_excela


SUPPORTED_EXT = {".xls", ".xlsx"}


def is_excel_path(p: Path) -> bool:
    return p.suffix.lower() in SUPPORTED_EXT


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.style = tb.Style(theme="cosmo")
        self.root.title("Raport PMID — Operations ↔ Loyalty")
        self.root.geometry("900x640")
        self.root.minsize(860, 600)

        # --- Zmienne stanu ---
        self.auto_mode  = tk.BooleanVar(value=True)
        self.ops_path   = tk.StringVar(value="")
        self.loy_paths  = tk.StringVar(value="")
        self.out_path   = tk.StringVar(value="")
        self.tolerance  = tk.StringVar(value="0.10")
        self.open_after = tk.BooleanVar(value=True)
        self.timestamp  = tk.BooleanVar(value=False)

        # --- UI ---
        self._build_ui()
        self._bind_state()

    def _extract_paths_from_dnd(self, data: str) -> list[Path]:
        out = []
        for token in self.root.tk.splitlist(data):
            s = str(token).strip().strip('"').strip()
            s = re.sub(r'[\u202A-\u202E\u200E\u200F]', '', s)  # niewidoczne
            m = re.search(r"[A-Za-z]:\\", s)
            if m: s = s[m.start():]
            if s.startswith("{") and s.endswith("}"): s = s[1:-1]
            p = Path(s)
            if p.exists(): out.append(p)
        return out

    # ---------- UI ----------
    def _build_ui(self):
        pad = dict(padx=10, pady=10)

        # Pliki
        frm_files = tb.Labelframe(self.root, text="Pliki wejściowe", padding=10)
        frm_files.pack(fill=X, **pad)

        tb.Checkbutton(
            frm_files,
            text="Automatycznie znajdź pliki w folderze programu",
            variable=self.auto_mode,
            command=self._bind_state
        ).grid(row=0, column=0, columnspan=6, sticky=W)

        # Operations
        tb.Label(frm_files, text="Operations (.xls/.xlsx):").grid(row=1, column=0, sticky=E, pady=(10, 3))
        self.ent_ops = tb.Entry(frm_files, textvariable=self.ops_path, width=70)
        self.ent_ops.grid(row=1, column=1, columnspan=4, sticky=EW, pady=(10, 3))
        self.btn_pick_ops = tb.Button(frm_files, text="Wybierz…", command=self._choose_ops)
        self.btn_pick_ops.grid(row=1, column=5, sticky=EW, pady=(10, 3))

        self.lbl_ops_hint = tb.Label(
            frm_files,
            text="Wskazówka: możesz też przeciągnąć i upuścić plik tutaj",
            bootstyle=INFO
        )
        self.lbl_ops_hint.grid(row=2, column=1, columnspan=5, sticky=W, pady=(0, 6))

        # Loyalty
        tb.Label(frm_files, text="Loyalty (.xls/.xlsx):").grid(row=3, column=0, sticky=E)
        self.ent_loy = tb.Entry(frm_files, textvariable=self.loy_paths, width=70)
        self.ent_loy.grid(row=3, column=1, columnspan=4, sticky=EW)
        self.btn_pick_loy = tb.Button(frm_files, text="Wybierz (1 lub wiele)…", command=self._choose_loy_many)
        self.btn_pick_loy.grid(row=3, column=5, sticky=EW)

        self.lbl_loy_hint = tb.Label(
            frm_files,
            text="Wskazówka: możesz upuścić jeden lub wiele plików (przytrzymaj Ctrl)",
            bootstyle=INFO
        )
        self.lbl_loy_hint.grid(row=4, column=1, columnspan=5, sticky=W, pady=(0, 6))

        frm_files.columnconfigure(1, weight=1)
        frm_files.columnconfigure(2, weight=1)
        frm_files.columnconfigure(3, weight=1)
        frm_files.columnconfigure(4, weight=1)

        # Ustawienia
        frm_settings = tb.Labelframe(self.root, text="Ustawienia", padding=10)
        frm_settings.pack(fill=X, **pad)

        tb.Label(frm_settings, text="Tolerancja Δ").grid(row=0, column=0, sticky=E)
        self.ent_tol = tb.Entry(frm_settings, textvariable=self.tolerance, width=8)
        self.ent_tol.grid(row=0, column=1, sticky=W)

        tb.Label(frm_settings, text="Plik wyjściowy (.xlsx)").grid(row=0, column=2, sticky=E)
        self.ent_out = tb.Entry(frm_settings, textvariable=self.out_path)
        self.ent_out.grid(row=0, column=3, columnspan=2, sticky=EW)
        self.btn_pick_out = tb.Button(frm_settings, text="Zapisz jako…", command=self._choose_out)
        self.btn_pick_out.grid(row=0, column=5, sticky=EW)

        tb.Checkbutton(
            frm_settings, text="Otwórz raport po zapisie", variable=self.open_after
        ).grid(row=1, column=0, columnspan=2, sticky=W, pady=(8, 0))
        tb.Checkbutton(
            frm_settings, text="Dodać znacznik czasu do nazwy", variable=self.timestamp
        ).grid(row=1, column=2, columnspan=2, sticky=W, pady=(8, 0))

        frm_settings.columnconfigure(3, weight=1)
        frm_settings.columnconfigure(4, weight=1)

        # Akcje
        frm_actions = tb.Frame(self.root)
        frm_actions.pack(fill=X, **pad)
        self.btn_run = tb.Button(frm_actions, text="📊 Generuj raport", bootstyle=SUCCESS, command=self._run_clicked)
        self.btn_run.pack(side=LEFT)
        tb.Button(frm_actions, text="Zamknij", command=self.root.destroy).pack(side=RIGHT)

        # Progress + log
        frm_log = tb.Labelframe(self.root, text="Log", padding=10)
        frm_log.pack(fill=BOTH, expand=True, **pad)
        self.prog = tb.Progressbar(frm_log, mode="indeterminate")
        self.prog.pack(fill=X, pady=(0, 8))
        self.txt = tk.Text(frm_log, height=18, wrap="word")
        self.txt.pack(fill=BOTH, expand=True)

        # Drag & Drop – rejestracja
        self._setup_dnd()

    def _setup_dnd(self):
        if not DND_OK:
            # Brak tkinterdnd2 — pokaż info
            info = "Przeciągnij i upuść: niedostępne (zainstaluj 'tkinterdnd2')"
            self.lbl_ops_hint.configure(text=info, bootstyle=WARNING)
            self.lbl_loy_hint.configure(text=info, bootstyle=WARNING)
            return

        # Rejestracja celów Drop dla pól Entry
        for widget in (self.ent_ops, self.ent_loy):
            try:
                widget.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
                widget.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]
            except Exception:
                pass

    def _on_drop(self, event):
        """Obsługa upuszczenia plików na Entry (Operations i Loyalty)."""
        try:
            if self.auto_mode.get():
                self.auto_mode.set(False)
                self._bind_state()

            paths = self._extract_paths_from_dnd(event.data)
            excel_paths = [p for p in paths if is_excel_path(p)]
            if not excel_paths:
                messagebox.showwarning("Nieobsługiwany plik", "Upuść pliki .xls lub .xlsx")
                return

            widget = event.widget
            # deduplikacja
            uniq, seen = [], set()
            for p in excel_paths:
                s = str(p)
                if s not in seen:
                    uniq.append(p);
                    seen.add(s)

            if widget is self.ent_ops:
                self.ops_path.set("; ".join(str(p) for p in uniq))
                self.log(f"📥 Ustawiono Operations ({len(uniq)}): " + ", ".join(p.name for p in uniq))
            elif widget is self.ent_loy:
                self.loy_paths.set("; ".join(str(p) for p in uniq))
                self.log(f"📥 Ustawiono Loyalty ({len(uniq)}): " + ", ".join(p.name for p in uniq))
        except Exception:
            self.log("❌ Błąd parsowania DnD:")
            self.log(traceback.format_exc())

    def _bind_state(self):
        auto = self.auto_mode.get()
        state = tk.DISABLED if auto else tk.NORMAL

        # Pola i przyciski wyboru
        for w in (self.ent_ops, self.ent_loy, self.ent_out,
                  self.btn_pick_ops, self.btn_pick_loy, self.btn_pick_out):
            w.configure(state=state)

        # Podpowiedzi DnD – pokazuj tylko w trybie ręcznym
        visible = not auto and DND_OK
        self.lbl_ops_hint.configure(text=(
            "Wskazówka: możesz też przeciągnąć i upuścić plik tutaj" if visible
            else ("Przeciągnij i upuść: włącz tryb ręczny" if auto else "Przeciągnij i upuść: niedostępne")
        ), bootstyle=INFO if visible else WARNING)

        self.lbl_loy_hint.configure(text=(
            "Wskazówka: możesz upuścić jeden lub wiele plików (przytrzymaj Ctrl)" if visible
            else ("Przeciągnij i upuść: włącz tryb ręczny" if auto else "Przeciągnij i upuść: niedostępne")
        ), bootstyle=INFO if visible else WARNING)

    # ---------- Helpers ----------
    def log(self, msg: str):
        self.txt.insert(tk.END, msg + "\n")
        self.txt.see(tk.END)
        self.root.update_idletasks()

    def _choose_ops(self):
        if self.auto_mode.get():
            return
        paths = filedialog.askopenfilenames(
            title="Wybierz plik(i) Operations",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Wszystkie pliki", "*.*")]
        )
        if paths:
            self.ops_path.set("; ".join(paths))

    def _choose_loy_many(self):
        if self.auto_mode.get():
            return
        paths = filedialog.askopenfilenames(
            title="Wybierz plik(i) Loyalty",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Wszystkie pliki", "*.*")]
        )
        if paths:
            self.loy_paths.set("; ".join(paths))

    def _choose_out(self):
        if self.auto_mode.get():
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            title="Zapisz raport jako"
        )
        if path:
            self.out_path.set(path)

    # ---------- Run ----------
    def _run_clicked(self):
        t = threading.Thread(target=self._run_safe, daemon=True)
        self._set_busy(True)
        t.start()

    def _set_busy(self, busy: bool):
        if busy:
            self.btn_run.configure(state=DISABLED)
            self.prog.start(8)
        else:
            self.btn_run.configure(state=NORMAL)
            self.prog.stop()

    def _run_safe(self):
        try:
            self._run_job()
        except Exception:
            self.log("❌ Błąd:")
            self.log(traceback.format_exc())
            messagebox.showerror("Błąd", "Wystąpił błąd. Szczegóły w logu.")
        finally:
            self._set_busy(False)

    def _run_job(self):
        # tolerancja
        tol_str = (self.tolerance.get() or "0.10").replace(",", ".")
        try:
            tol = float(tol_str)
        except ValueError:
            messagebox.showerror("Błąd", "Błędna tolerancja. Użyj np. 0.10")
            return

        # --- ŹRÓDŁA DANYCH ---
        if self.auto_mode.get():
            root = base_dir()

            # Operations (auto – jeden plik)
            p_ops = znajdz_plik_operations(root)
            ops_df = wczytaj_operations(str(p_ops))
            ops_names = [Path(p_ops).name]

            # Loyalty (auto – jeden plik)
            p_loy = znajdz_plik_loyalty(root)
            loy_paths = [Path(p_loy)]
            loy_names = [Path(p_loy).name]
        else:
            # Operations (ręcznie – jeden lub wiele)
            ops = self.ops_path.get().strip()
            if not ops:
                messagebox.showwarning("Brak plików", "Wybierz lub upuść plik(i) Operations.")
                return
            ops_list = [Path(p) for p in ops.split(";") if p.strip()]
            ops_names = [p.name for p in ops_list]
            ops_df = wczytaj_operations_many(ops_list) if len(ops_list) > 1 else wczytaj_operations(str(ops_list[0]))

            # Loyalty (ręcznie – jeden lub wiele)
            loy = self.loy_paths.get().strip()
            if not loy:
                messagebox.showwarning("Brak plików", "Wybierz lub upuść co najmniej jeden plik Loyalty.")
                return
            loy_paths = [Path(p) for p in loy.split(";") if p.strip()]
            loy_names = [p.name for p in loy_paths]

        # log: Operations
        if len(ops_names) == 1:
            self.log(f"🔎 Operations: {ops_names[0]}")
        else:
            self.log(f"🔎 Operations (x{len(ops_names)}): " + ", ".join(ops_names))

        # log: Loyalty
        if len(loy_names) == 1:
            self.log(f"🔎 Loyalty:    {loy_names[0]}")
        else:
            self.log(f"🔎 Loyalty (x{len(loy_names)}): " + ", ".join(loy_names))

        # wczytanie Loyalty
        lojal_df = wczytaj_loyalty(str(loy_paths[0])) if len(loy_paths) == 1 else wczytaj_loyalty_many(loy_paths)

        # porównanie
        wyniki = porownaj(lojal_df, ops_df, tolerancja=tol)

        # wyjściowa ścieżka
        out = Path(self.out_path.get()) if self.out_path.get().strip() else wybierz_sciezke_wyjsciowa(base_dir())
        if self.timestamp.get():
            stem = out.stem
            suf  = out.suffix or ".xlsx"
            out  = out.with_name(f"{stem} - {date.today().isoformat()}{suf}")

        if out.exists():
            self.log(f"ℹ️ Uwaga: {out.name} zostanie nadpisany (najstarszy w cyklu 01..31).")

        # zapis
        zapisz_do_excela(wyniki, out)
        self.log(f"✅ Gotowe. Otwórz plik: {out.name}")

        if self.open_after.get():
            try:
                if os.name == "nt":
                    os.startfile(out)  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    os.system(f"open '{out}'")
                else:
                    os.system(f"xdg-open '{out}'")
            except Exception:
                pass

    def run(self):
        self.root.mainloop()


def run_gui():
    # Jeśli mamy tkinterdnd2 – użyj jego klasy okna; w przeciwnym razie zwykły Tk
    if DND_OK:
        root = TkinterDnD.Tk()  # type: ignore[call-arg]
    else:
        root = tk.Tk()
    app = App(root)
    app.run()


if __name__ == "__main__":
    run_gui()
