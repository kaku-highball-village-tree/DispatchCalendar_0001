from __future__ import annotations

import importlib
import sys
import types
from pathlib import Path


def import_cmd_module():
    """Import the CMD module with optional runtime dependencies stubbed for unit tests."""
    obj_openpyxl_module = types.ModuleType("openpyxl")
    obj_openpyxl_module.load_workbook = lambda *args, **kwargs: None
    sys.modules.setdefault("openpyxl", obj_openpyxl_module)

    for psz_module_name in [
        "google",
        "google.auth",
        "google.auth.transport",
        "google.auth.transport.requests",
        "google.oauth2",
        "google.oauth2.credentials",
        "google_auth_oauthlib",
        "google_auth_oauthlib.flow",
        "googleapiclient",
        "googleapiclient.discovery",
        "googleapiclient.errors",
    ]:
        sys.modules.setdefault(psz_module_name, types.ModuleType(psz_module_name))

    sys.modules["google.auth.transport.requests"].Request = object
    sys.modules["google.oauth2.credentials"].Credentials = type(
        "Credentials",
        (),
        {"from_authorized_user_file": classmethod(lambda cls, *args, **kwargs: cls())},
    )
    sys.modules["google_auth_oauthlib.flow"].InstalledAppFlow = object
    sys.modules["googleapiclient.discovery"].build = lambda *args, **kwargs: None
    sys.modules["googleapiclient.errors"].HttpError = Exception

    return importlib.import_module("src.DispatchCalendar_Cmd")


def test_step0001_5_expands_slots_without_deleting_step0001_third_column(tmp_path: Path) -> None:
    obj_cmd_module = import_cmd_module()
    obj_step0001_path = tmp_path / "配車_0508_step0001.tsv"
    obj_step0001_path.write_text(
        "令和8年　5月　8日（金）\n"
        "header\n"
        "浅野　龍太\t7002\t旧D列_A\t旧E列_B\t旧F列_C\n"
        "\t\t補足_a\t補足_b\t補足_c\n",
        encoding="utf-8",
        newline="\r\n",
    )

    obj_step0001_5_path = Path(obj_cmd_module.create_step0001_5_tsv_from_step0001_tsv(str(obj_step0001_path)))

    assert obj_step0001_5_path.name == "配車_0508_step0001_5.tsv"
    assert obj_step0001_5_path.read_text(encoding="utf-8").replace("\r\n", "\n") == (
        "令和8年　5月　8日（金）\n"
        "header\n"
        "浅野　龍太\t7002\t旧D列_A\n"
        "\t\t補足_a\n"
        "浅野　龍太\t7002\t旧E列_B\n"
        "\t\t補足_b\n"
        "浅野　龍太\t7002\t旧F列_C\n"
        "\t\t補足_c\n"
    )
