"""命名・コピー部分の自動テスト（pytest）。GUIは対象外。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from sokutei_namer import DEFAULT_TEMPLATE, Item, build_name, execute, guess_instrument, plan  # noqa: E402

COMMON = dict(date="20261118", test="2", freq="500", volt="0.65", run="1")


def test_guess_instrument(tmp_path):
    (tmp_path / "a.csv").write_text("KEYENCE NR-X100\n1,2\n")
    (tmp_path / "b.CSV").write_text("O-Solution DS-5100\n")
    (tmp_path / "c.xlsm").write_bytes(b"x")
    (tmp_path / "d.dwd").write_bytes(b"x")
    (tmp_path / "e.csv").write_text("plain\n")
    got = {p.name: guess_instrument(p) for p in tmp_path.iterdir()}
    assert got == {"a.csv": "NRX", "b.CSV": "DS5100", "c.xlsm": "NRX", "d.dwd": "DYNO", "e.csv": ""}


def test_build_name_drops_empty_tokens_and_bad_chars():
    f = dict(COMMON, inst="", memo="")
    assert build_name(DEFAULT_TEMPLATE, f, Path("x.CSV")) == "2-500-0.65-1.csv"
    assert build_name("{test}:{freq}", f, Path("x.csv")) == "2_500.csv"


def test_plan_never_overwrites_and_execute_logs(tmp_path):
    src, out = tmp_path / "src", tmp_path / "out"
    src.mkdir(); out.mkdir()
    (src / "a.csv").write_text("KEYENCE\n")
    (src / "b.csv").write_text("KEYENCE\n")
    (out / "2-500-0.65-1_NRX.csv").write_text("existing")
    items = [Item(src=src / "a.csv", inst="NRX"), Item(src=src / "b.csv", inst="NRX")]
    pairs = plan(items, DEFAULT_TEMPLATE, COMMON, out)
    assert [d.name for _, d in pairs] == ["2-500-0.65-1_NRX_2.csv", "2-500-0.65-1_NRX_3.csv"]
    ok, errs = execute(pairs)
    assert ok == 2 and not errs
    assert (out / "2-500-0.65-1_NRX.csv").read_text() == "existing"
    assert (src / "a.csv").exists()  # 既定はコピー
    assert "2-500-0.65-1_NRX_2.csv" in (out / "rename_log.csv").read_text(encoding="utf-8-sig")


# ---- 開いている Excel（COM）の代わりに模擬オブジェクトで確認
class _FakeApp:
    def __init__(self):
        self.DisplayAlerts = True
        self.ActiveWorkbook = None
        self.saved = []


class _FakeSheet:
    def __init__(self, wb):
        self.wb = wb

    def Copy(self):
        self.wb.Application.ActiveWorkbook = _FakeWB("Sheet1", self.wb.Application)

    def Range(self, _):
        class R:
            Value = (("KEYENCE WAVE LOGGER X", None), (1.0, 2.0))
        return R()


class _FakeWB:
    def __init__(self, name, app=None):
        self.Name, self.Path, self.Application = name, "", app or _FakeApp()
        self.closed = False

    def Worksheets(self, _):
        return _FakeSheet(self)

    def SaveAs(self, path, fmt):
        assert self.Application.DisplayAlerts is False
        Path(path).write_text(f"fmt={fmt}")
        self.Application.saved.append((Path(path).name, fmt))

    def Close(self, _):
        self.closed = True


def test_open_excel_saved_with_rule_and_csv(tmp_path):
    from sokutei_namer import guess_instrument_wb
    wb = _FakeWB("Book1")
    assert guess_instrument_wb(wb) == "NRX"
    items = [Item(src=Path("Book1"), inst="NRX", wb=wb)]
    pairs = plan(items, DEFAULT_TEMPLATE, COMMON, tmp_path, xl_ext=".xlsx", also_csv=True)
    assert pairs[0][1].name == "2-500-0.65-1_NRX.xlsx"
    assert items[0].csv_dst.name == "2-500-0.65-1_NRX.csv"
    ok, errs = execute(pairs, close_excel=True)
    assert ok == 1 and not errs, errs
    assert wb.Application.saved == [("2-500-0.65-1_NRX.xlsx", 51), ("2-500-0.65-1_NRX.csv", 6)]
    assert wb.Application.DisplayAlerts is True  # 元に戻す
    assert wb.closed
    log = (tmp_path / "rename_log.csv").read_text(encoding="utf-8-sig")
    assert "[開いているExcel] Book1" in log and "excel-csv" in log


def test_open_excel_xlsm_and_manual_name(tmp_path):
    wb = _FakeWB("Book2")
    it = Item(src=Path("Book2"), inst="NRX", wb=wb, new_name="特別.xlsx", manual=True)
    pairs = plan([it], DEFAULT_TEMPLATE, COMMON, tmp_path, xl_ext=".xlsm")
    assert pairs[0][1].name == "特別.xlsm" and it.csv_dst is None
