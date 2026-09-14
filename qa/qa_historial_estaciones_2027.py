from __future__ import annotations

import json
import os
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from converter.conversion_history import (
    STATUS_CONVERTED,
    STATUS_FAILED,
    create_history_record,
    finish_history_record,
    list_recent_history,
)
from converter.workstation_identity import (
    load_station_identity,
    save_station_alias,
)


def main() -> None:
    original_localappdata = os.environ.get(
        "LOCALAPPDATA"
    )

    with tempfile.TemporaryDirectory() as td:
        temp_root = Path(td)
        os.environ["LOCALAPPDATA"] = str(
            temp_root / "localappdata"
        )

        results_root = (
            temp_root
            / "SCIELO_RESULTADOS_QA"
        )
        results_root.mkdir()

        source = temp_root / "1354 Completo.docx"
        source.write_bytes(b"QA")

        identity = save_station_alias(
            "QA Equipo 01"
        )

        loaded = load_station_identity()

        assert loaded.alias == "QA Equipo 01"
        assert loaded.hostname
        assert loaded.config_path.is_file()

        handle = create_history_record(
            source,
            results_root=results_root,
            station=identity,
        )

        finish_history_record(
            handle,
            STATUS_CONVERTED,
            result_directory=(
                results_root
                / "1354 Completo"
            ),
            warning_count=2,
        )

        data = json.loads(
            handle.path.read_text(
                encoding="utf-8"
            )
        )

        assert data["status"] == STATUS_CONVERTED
        assert data["workstation_alias"] == "QA Equipo 01"
        assert data["warning_count"] == 2

        def worker(index: int) -> str:
            docx = temp_root / f"articulo_{index}.docx"
            docx.write_bytes(b"QA")

            item = create_history_record(
                docx,
                results_root=results_root,
                station=identity,
            )

            finish_history_record(
                item,
                (
                    STATUS_FAILED
                    if index % 4 == 0
                    else STATUS_CONVERTED
                ),
                result_directory=(
                    results_root
                    / docx.stem
                ),
                error=(
                    "Fallo QA"
                    if index % 4 == 0
                    else ""
                ),
            )

            return str(item.path)

        with ThreadPoolExecutor(
            max_workers=8
        ) as executor:
            paths = list(
                executor.map(
                    worker,
                    range(20),
                )
            )

        assert len(set(paths)) == 20

        recent = list_recent_history(
            results_root=results_root,
            limit=50,
        )

        assert len(recent) == 21
        assert all(
            record.get("status")
            in {STATUS_CONVERTED, STATUS_FAILED}
            for record in recent
        )

        print("IDENTIDAD DE ESTACIÓN: APROBADA")
        print("ALIAS PERSISTENTE: APROBADO")
        print("HISTORIAL CONVERTIDO/FALLIDO: APROBADO")
        print("ARCHIVOS JSON INDEPENDIENTES: APROBADO")
        print("CONCURRENCIA BÁSICA 20 REGISTROS: APROBADA")
        print("HISTORIAL + ESTACIONES 2027: APROBADO")

    if original_localappdata is None:
        os.environ.pop(
            "LOCALAPPDATA",
            None,
        )
    else:
        os.environ["LOCALAPPDATA"] = (
            original_localappdata
        )


if __name__ == "__main__":
    main()
