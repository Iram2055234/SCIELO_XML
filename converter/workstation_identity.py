from __future__ import annotations

import json
import os
import socket
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class WorkstationIdentity:
    alias: str
    hostname: str
    config_path: Path


def get_local_appdata_directory() -> Path:
    value = os.environ.get("LOCALAPPDATA")

    if value:
        return Path(value).resolve()

    return (
        Path.home()
        / "AppData"
        / "Local"
    ).resolve()


def get_station_config_path() -> Path:
    return (
        get_local_appdata_directory()
        / "SCIELO_FACPyA"
        / "station.json"
    )


def get_hostname() -> str:
    hostname = str(
        socket.gethostname()
        or os.environ.get("COMPUTERNAME")
        or "Equipo"
    ).strip()

    return hostname or "Equipo"


def clean_alias(value: str) -> str:
    alias = " ".join(
        str(value or "").split()
    ).strip()

    if len(alias) > 80:
        alias = alias[:80].rstrip()

    return alias


def load_station_identity() -> WorkstationIdentity:
    path = get_station_config_path()
    hostname = get_hostname()
    alias = ""

    if path.is_file():
        try:
            data = json.loads(
                path.read_text(
                    encoding="utf-8",
                )
            )

            alias = clean_alias(
                str(
                    data.get(
                        "alias",
                        "",
                    )
                )
            )

        except Exception:
            alias = ""

    if not alias:
        alias = hostname

    return WorkstationIdentity(
        alias=alias,
        hostname=hostname,
        config_path=path,
    )


def save_station_alias(
    alias: str,
) -> WorkstationIdentity:
    cleaned_alias = clean_alias(alias)

    if not cleaned_alias:
        raise ValueError(
            "El nombre del equipo no puede quedar vacío."
        )

    path = get_station_config_path()
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    data = {
        "alias": cleaned_alias,
        "hostname": get_hostname(),
    }

    temp_path = path.with_suffix(
        ".json.tmp"
    )

    temp_path.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    temp_path.replace(path)

    return load_station_identity()
