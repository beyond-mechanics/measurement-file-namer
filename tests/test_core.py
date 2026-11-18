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
