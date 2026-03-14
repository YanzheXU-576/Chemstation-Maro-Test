from pathlib import Path


MACRO_FILE = Path(__file__).resolve().parents[1] / "start_single_run.mac"


def _macro_text() -> str:
    return MACRO_FILE.read_text(encoding="utf-8")


def test_macro_file_exists() -> None:
    assert MACRO_FILE.exists(), "start_single_run.mac 文件不存在"


def test_macro_contains_key_commands() -> None:
    text = _macro_text()

    assert "PrepRun" in text
    assert "while loop = 1" in text
    assert 'if ACQSTATUS$ = "PRERUN"' in text
    assert "StartMethod" in text
    assert "EndMacro" in text


def test_macro_execution_order() -> None:
    text = _macro_text()

    prep_idx = text.index("PrepRun")
    while_idx = text.index("while loop = 1")
    status_check_idx = text.index('if ACQSTATUS$ = "PRERUN"')
    start_method_idx = text.index("StartMethod")
    end_macro_idx = text.index("EndMacro")

    assert prep_idx < while_idx < status_check_idx < start_method_idx < end_macro_idx
