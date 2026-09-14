from __future__ import annotations

import os
import subprocess
from pathlib import Path


def validate_file(
    file_path: str | Path,
) -> Path:
    """
    Comprueba que exista un archivo.
    """

    path = Path(
        file_path
    ).resolve()

    if not path.is_file():
        raise FileNotFoundError(
            f"No se encontró el archivo: {path}"
        )

    return path


def validate_directory(
    directory_path: str | Path,
) -> Path:
    """
    Comprueba que exista una carpeta.
    """

    path = Path(
        directory_path
    ).resolve()

    if not path.is_dir():
        raise FileNotFoundError(
            f"No se encontró la carpeta: {path}"
        )

    return path


def open_html(
    html_path: str | Path,
) -> None:
    """
    Abre el HTML con el navegador predeterminado
    de Windows.
    """

    path = validate_file(
        html_path
    )

    os.startfile(
        str(path)
    )


def open_xml_in_notepad(
    xml_path: str | Path,
) -> None:
    """
    Abre explícitamente el XML con Bloc de notas.

    No utiliza la aplicación predeterminada de
    Windows para evitar que se abra VS Code,
    Visual Studio u otro editor técnico.
    """

    path = validate_file(
        xml_path
    )

    subprocess.Popen(
        [
            "notepad.exe",
            str(path),
        ]
    )


def open_directory(
    directory_path: str | Path,
) -> None:
    """
    Abre una carpeta con el Explorador de Windows.
    """

    path = validate_directory(
        directory_path
    )

    os.startfile(
        str(path)
    )


def reveal_file_in_explorer(
    file_path: str | Path,
) -> None:
    """
    Abre el Explorador de Windows seleccionando
    un archivo específico.

    Esta función queda disponible para usos
    posteriores, por ejemplo para localizar el
    XML antes de subirlo al validador de SciELO.
    """

    path = validate_file(
        file_path
    )

    subprocess.Popen(
        [
            "explorer.exe",
            "/select,",
            str(path),
        ]
    )