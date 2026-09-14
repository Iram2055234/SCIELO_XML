from __future__ import annotations

from pathlib import Path

from converter.docx_reader import read_docx
from converter.metadata_parser import parse_metadata


def find_1355() -> Path:
    candidates = [
        Path.cwd() / "1355 Completo.docx",
        Path.cwd() / "input" / "1355 Completo.docx",
        Path.home() / "Downloads" / "1355 Completo.docx",
        Path(
            r"\\148.234.45.124\Secretaria de Investigacion"
            r"\CP\VinculaTégica EFAN1\2026"
            r"\Vinculategica 12_4\2. Producción"
            r"\1355 Completo.docx"
        ),
    ]

    for path in candidates:
        if path.is_file():
            return path

    raise FileNotFoundError(
        "No encontré 1355 Completo.docx."
    )


def main() -> None:
    path = find_1355()
    article = parse_metadata(
        read_docx(path)
    )

    assert article.title.startswith(
        "Determinantes de la capacidad de innovación"
    )
    assert article.translated_title.startswith(
        "Determinants of Innovation Capability"
    )
    assert len(article.authors) == 3
    assert len(article.affiliations) == 3
    assert article.jel_codes == [
        "O31",
        "O32",
        "O34",
        "O54",
    ]

    print("TÍTULO ESPAÑOL RECUPERADO: APROBADO")
    print("TÍTULO INGLÉS PRESERVADO: APROBADO")
    print("AUTORES/AFILIACIONES: APROBADOS")
    print("JEL: APROBADO")
    print("METADATA FALLBACK PARCIAL 1355: APROBADO")


if __name__ == "__main__":
    main()
