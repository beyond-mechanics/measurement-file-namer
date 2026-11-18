# -*- coding: utf-8 -*-
"""
計測ファイル一括命名・保存ツール（試作 v0.1）

1回の試験で取れた .csv / .xlsm / .xlsx / .dwd などを一度に並べ、
共通の試験条件（試験No・周波数・電圧・回数・日付）と機器タグから
ファイル名を決めて、指定フォルダへまとめて「コピー」する。

  * 元ファイルは既定で消さない（コピー）。移動はチェックを入れたときだけ。
  * 同名ファイルがあっても上書きしない（_2, _3 … を付ける）。
  * コピー後にサイズを照合し、保存先に rename_log.csv（元パス→新名）を追記。
  * CSV の先頭数行から機器を推定（KEYENCE / O-Solution / DynoWare）。

起動: python sokutei_namer.py
exe化: build_exe.bat を実行（PyInstaller）
"""
from __future__ import annotations

import csv
import datetime as _dt
import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path

APP_TITLE = "計測ファイル一括命名ツール v0.1"
DEFAULT_TEMPLATE = "{test}-{freq}-{volt}-{run}_{inst}"
INST_PRESETS = ["NRX", "DS5100", "DYNO", "LK", "RP", ""]
TOKENS_HELP = ("{date} 日付YYYYMMDD / {test} 試験No / {freq} 周波数 / "
               "{volt} 電圧 / {run} 回数 / {inst} 機器 / {memo} メモ / {orig} 元の名前")
_BAD_CHARS = re.compile(r'[\\/:*?"<>|]')


# ---------------------------------------------------------------- 命名ロジック
def guess_instrument(path: Path) -> str:
    """拡張子と先頭数行から機器タグを推定する（外れたら画面で直せばよい）。"""
    ext = path.suffix.lower()
    if ext == ".dwd":
        return "DYNO"
    if ext in (".xlsm", ".xlsx", ".xls"):
        return "NRX"  # WAVE LOGGER X の Excel 出力を想定
    if ext in (".csv", ".txt"):
        try:
            with open(path, "rb") as f:
                head = f.read(4096)
            text = ""
            for enc in ("utf-8", "cp932"):
                try:
                    text = head.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            t = text.upper()
            if "KEYENCE" in t or "NR-X" in t or "WAVE LOGGER" in t:
                return "NRX"
            if "O-SOLUTION" in t or "DS-5" in t or "ONO SOKKI" in t or "小野測器" in text:
                return "DS5100"
            if "DYNOWARE" in t or "KISTLER" in t:
                return "DYNO"
        except OSError:
            pass
    return ""


def clean_name(stem: str) -> str:
    """禁止文字を _ に、空トークンで生じた区切りの重複・端の区切りを整理する。"""
    s = _BAD_CHARS.sub("_", stem)
    s = re.sub(r"([-_ ])[-_ ]+", r"\1", s)   # "--" "_-" などを1つに
    return s.strip("-_ .")


def build_name(template: str, fields: dict, orig: Path) -> str:
    vals = {k: str(v).strip() for k, v in fields.items()}
    vals["orig"] = orig.stem

    def rep(m):
        return vals.get(m.group(1), m.group(0))

    stem = re.sub(r"\{(\w+)\}", rep, template)
    stem = clean_name(stem) or orig.stem
    return stem + orig.suffix.lower()


def unique_dest(dest_dir: Path, name: str, taken: set[str]) -> Path:
    """保存先の既存ファイルと今回の予定名の両方とぶつからない名前を返す。"""
    p = Path(name)
    cand, i = name, 2
    while (dest_dir / cand).exists() or cand.lower() in taken:
        cand = f"{p.stem}_{i}{p.suffix}"
        i += 1
    taken.add(cand.lower())
    return dest_dir / cand


@dataclass
class Item:
    src: Path
    inst: str = ""
    memo: str = ""
    new_name: str = ""
    manual: bool = False           # 新しい名前を手で直したか
    extra: dict = field(default_factory=dict)


def plan(items: list[Item], template: str, common: dict, dest_dir: Path) -> list[tuple[Item, Path]]:
    taken: set[str] = set()
    out = []
    for it in items:
        if not it.manual:
            f = dict(common, inst=it.inst, memo=it.memo)
            it.new_name = build_name(template, f, it.src)
        out.append((it, unique_dest(dest_dir, it.new_name, taken)))
    return out


def execute(pairs, move: bool = False, log_name: str = "rename_log.csv", progress=None):
    """コピー（または移動）してサイズ照合し、ログを追記。戻り値: (成功数, エラー一覧)"""
    if not pairs:
        return 0, []
    dest_dir = pairs[0][1].parent
    dest_dir.mkdir(parents=True, exist_ok=True)
    ok, errors, rows = 0, [], []
    now = _dt.datetime.now().isoformat(timespec="seconds")
    for n, (it, dst) in enumerate(pairs, 1):
        try:
            size = it.src.stat().st_size
            shutil.copy2(it.src, dst)
            if dst.stat().st_size != size:
                raise IOError("コピー後のサイズが一致しません")
            if move:
                it.src.unlink()
            ok += 1
            rows.append([now, str(it.src), dst.name, size, it.inst, "move" if move else "copy"])
        except Exception as e:  # noqa: BLE001
            errors.append(f"{it.src.name}: {e}")
        if progress:
            progress(n, len(pairs))
    log = dest_dir / log_name
    new = not log.exists()
    with open(log, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["日時", "元ファイル", "新しい名前", "サイズ[byte]", "機器", "操作"])
        w.writerows(rows)
    return ok, errors


# ---------------------------------------------------------------- GUI
def main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    try:  # ドラッグ＆ドロップは tkinterdnd2 があれば有効（pip install tkinterdnd2）
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        has_dnd = True
    except Exception:  # noqa: BLE001
        root = tk.Tk()
        has_dnd = False

    root.title(APP_TITLE)
    root.geometry("1050x620")
    items: list[Item] = []

    # --- 共通条件
    top = ttk.LabelFrame(root, text="試験条件（全ファイル共通）")
    top.pack(fill="x", padx=8, pady=6)
    v = {k: tk.StringVar() for k in ("date", "test", "freq", "volt", "run")}
    v["date"].set(_dt.date.today().strftime("%Y%m%d"))
    v["run"].set("1")
    for col, (k, lab, w) in enumerate([("date", "日付", 10), ("test", "試験No", 6),
                                       ("freq", "周波数[Hz]", 8), ("volt", "電圧[V]", 7),
                                       ("run", "回数", 4)]):
        ttk.Label(top, text=lab).grid(row=0, column=col * 2, padx=(8, 2), pady=4, sticky="e")
        ttk.Entry(top, textvariable=v[k], width=w).grid(row=0, column=col * 2 + 1, sticky="w")

    def bump_run():
        try:
            v["run"].set(str(int(v["run"].get()) + 1))
        except ValueError:
            v["run"].set("1")
    ttk.Button(top, text="回数+1", command=bump_run).grid(row=0, column=10, padx=6)

    tmpl = tk.StringVar(value=DEFAULT_TEMPLATE)
    ttk.Label(top, text="命名ルール").grid(row=1, column=0, padx=(8, 2), sticky="e")
    ttk.Entry(top, textvariable=tmpl, width=48).grid(row=1, column=1, columnspan=6, sticky="w", pady=4)
    ttk.Label(top, text=TOKENS_HELP, foreground="#666").grid(row=2, column=0, columnspan=12,
                                                              sticky="w", padx=8, pady=(0, 4))

    # --- 保存先
    dfrm = ttk.Frame(root)
    dfrm.pack(fill="x", padx=8)
    dest = tk.StringVar()
    ttk.Label(dfrm, text="保存先フォルダ").pack(side="left")
    ttk.Entry(dfrm, textvariable=dest).pack(side="left", fill="x", expand=True, padx=4)
    ttk.Button(dfrm, text="選択…",
               command=lambda: dest.set(filedialog.askdirectory() or dest.get())).pack(side="left")
    mkdir_sub = tk.BooleanVar(value=False)
    ttk.Checkbutton(dfrm, text="日付_試験No のサブフォルダを作る", variable=mkdir_sub).pack(side="left", padx=6)
    move = tk.BooleanVar(value=False)
    ttk.Checkbutton(dfrm, text="移動（元を消す）", variable=move).pack(side="left", padx=6)

    # --- 一覧
    cols = ("src", "inst", "memo", "new")
    tree = ttk.Treeview(root, columns=cols, show="headings", selectmode="extended")
    for c, t, w in [("src", "元ファイル", 360), ("inst", "機器（ダブルクリックで編集）", 170),
                    ("memo", "メモ", 120), ("new", "新しい名前（ダブルクリックで直接編集）", 360)]:
        tree.heading(c, text=t)
        tree.column(c, width=w, anchor="w")
    tree.pack(fill="both", expand=True, padx=8, pady=6)
    if has_dnd:
        hint = "ここにファイルをドラッグ＆ドロップ、または「ファイル追加」"
    else:
        hint = "「ファイル追加」で選択（tkinterdnd2 を入れるとドラッグ＆ドロップも可）"
    status = tk.StringVar(value=hint)

    def dest_dir() -> Path:
        d = Path(dest.get().strip() or ".")
        if mkdir_sub.get():
            d = d / clean_name(f"{v['date'].get()}_{v['test'].get()}")
        return d

    def common():
        return {k: x.get() for k, x in v.items()}

    def refresh(*_):
        pairs = plan(items, tmpl.get(), common(), dest_dir())
        tree.delete(*tree.get_children())
        for i, (it, dst) in enumerate(pairs):
            tree.insert("", "end", iid=str(i),
                        values=(str(it.src), it.inst, it.memo, dst.name))
        status.set(f"{len(items)} 件 → {dest_dir()}")

    for x in list(v.values()) + [tmpl, dest, mkdir_sub]:
        x.trace_add("write", refresh)

    def add_paths(paths):
        known = {it.src.resolve() for it in items}
        for p in paths:
            p = Path(p)
            if p.is_dir():
                add_paths([q for q in sorted(p.iterdir()) if q.is_file()])
                continue
            if p.is_file() and p.resolve() not in known:
                items.append(Item(src=p, inst=guess_instrument(p)))
                known.add(p.resolve())
        refresh()

    def add_files():
        add_paths(filedialog.askopenfilenames(
            filetypes=[("計測ファイル", "*.csv *.xlsm *.xlsx *.xls *.dwd *.txt"), ("すべて", "*.*")]))

    def remove_sel():
        for iid in sorted(map(int, tree.selection()), reverse=True):
            items.pop(iid)
        refresh()

    def edit_cell(event):
        iid, col = tree.identify_row(event.y), tree.identify_column(event.x)
        if not iid or col not in ("#2", "#3", "#4"):
            return
        it = items[int(iid)]
        x, y, w, h = tree.bbox(iid, col)
        key = {"#2": "inst", "#3": "memo", "#4": "new"}[col]
        cur = {"inst": it.inst, "memo": it.memo, "new": tree.set(iid, "new")}[key]
        if key == "inst":
            ed = ttk.Combobox(tree, values=INST_PRESETS)
        else:
            ed = ttk.Entry(tree)
        ed.insert(0, cur)
        ed.place(x=x, y=y, width=w, height=h)
        ed.focus_set()

        def done(_=None):
            val = ed.get().strip()
            ed.destroy()
            if key == "inst":
                it.inst = val
            elif key == "memo":
                it.memo = val
            else:
                it.new_name, it.manual = (val, True) if val else ("", False)
            refresh()
        ed.bind("<Return>", done)
        ed.bind("<FocusOut>", done)
        ed.bind("<Escape>", lambda _e: ed.destroy())

    tree.bind("<Double-1>", edit_cell)
    tree.bind("<Delete>", lambda _e: remove_sel())
    if has_dnd:
        tree.drop_target_register(DND_FILES)
        tree.dnd_bind("<<Drop>>", lambda e: add_paths(root.tk.splitlist(e.data)))

    def run():
        if not items:
            messagebox.showinfo(APP_TITLE, "ファイルがまだありません。")
            return
        if not dest.get().strip():
            messagebox.showwarning(APP_TITLE, "保存先フォルダを選んでください。")
            return
        pairs = plan(items, tmpl.get(), common(), dest_dir())
        verb = "移動" if move.get() else "コピー"
        if not messagebox.askokcancel(APP_TITLE, f"{len(pairs)} 件を {dest_dir()} に{verb}します。"):
            return
        ok, errs = execute(pairs, move=move.get(),
                           progress=lambda n, t: (status.set(f"{verb}中… {n}/{t}"), root.update()))
        msg = f"{ok} 件 {verb}しました。ログ: rename_log.csv"
        if errs:
            msg += "\n\n失敗:\n" + "\n".join(errs)
            messagebox.showwarning(APP_TITLE, msg)
        else:
            messagebox.showinfo(APP_TITLE, msg)
            items.clear()
            bump_run()   # 次の1回に備えて回数を進める
        refresh()

    bar = ttk.Frame(root)
    bar.pack(fill="x", padx=8, pady=(0, 8))
    ttk.Button(bar, text="ファイル追加", command=add_files).pack(side="left")
    ttk.Button(bar, text="フォルダ追加",
               command=lambda: add_paths([d] if (d := filedialog.askdirectory()) else [])).pack(side="left", padx=4)
    ttk.Button(bar, text="選択を外す", command=remove_sel).pack(side="left")
    ttk.Button(bar, text="全部クリア", command=lambda: (items.clear(), refresh())).pack(side="left", padx=4)
    ttk.Label(bar, textvariable=status, foreground="#335").pack(side="left", padx=12)
    ttk.Button(bar, text="一括保存 ▶", command=run).pack(side="right")

    refresh()
    root.mainloop()


if __name__ == "__main__":
    main()
