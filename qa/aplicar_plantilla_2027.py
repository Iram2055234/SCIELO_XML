from __future__ import annotations

import py_compile
import re
import shutil
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

METADATA_PATH = (
    PROJECT_ROOT
    / "converter"
    / "metadata_parser.py"
)

BODY_PATH = (
    PROJECT_ROOT
    / "converter"
    / "body_parser.py"
)

METADATA_BACKUP = (
    PROJECT_ROOT
    / "converter"
    / "metadata_parser_antes_plantilla_2027.py"
)

BODY_BACKUP = (
    PROJECT_ROOT
    / "converter"
    / "body_parser_antes_plantilla_2027.py"
)


TITLE_CONSTANTS = '''TITLE_STYLE = "Título_Artículo"
SPANISH_TITLE_STYLE = "Título_Español"
ENGLISH_TITLE_STYLE = "Título_Ingles"
AUTHORS_STYLE = "Autores"
AUTHOR_DATA_STYLE = "ID_Autores"
ABSTRACT_STYLE = "Resumen_Abstract"

IGNORED_TEMPLATE_STYLES = {
    "Instruccion_Plantilla",
    "Ejemplo_Plantilla",
}
'''


TITLE_BRANCH = '''        if style in IGNORED_TEMPLATE_STYLES:
            continue

        if style == SPANISH_TITLE_STYLE:
            article.title = text
            continue

        if style == ENGLISH_TITLE_STYLE:
            article.translated_title = (
                remove_title_parentheses(
                    text
                )
            )
            continue

        if style == TITLE_STYLE:
            titles.append(text)
            continue

        if style == AUTHORS_STYLE:'''


BODY_STYLES = '''METADATA_STYLES = {
    "Título_Artículo",
    "Título_Español",
    "Título_Ingles",
    "Autores",
    "ID_Autores",
    "Resumen_Abstract",
}

IGNORED_TEMPLATE_STYLES = {
    "Instruccion_Plantilla",
    "Ejemplo_Plantilla",
}
'''


BODY_SKIP = '''        if (
            style in METADATA_STYLES
            or style in IGNORED_TEMPLATE_STYLES
        ):
            continue
'''


def backup_once(
    source: Path,
    backup: Path,
) -> None:
    if not backup.exists():
        shutil.copy2(
            source,
            backup,
        )


def patch_metadata() -> None:
    backup_once(
        METADATA_PATH,
        METADATA_BACKUP,
    )

    text = METADATA_PATH.read_text(
        encoding="utf-8"
    )

    if "SPANISH_TITLE_STYLE" not in text:
        pattern = re.compile(
            r'TITLE_STYLE = "Título_Artículo"\n'
            r'AUTHORS_STYLE = "Autores"\n'
            r'AUTHOR_DATA_STYLE = "ID_Autores"\n'
            r'ABSTRACT_STYLE = "Resumen_Abstract"\n'
        )

        text, count = pattern.subn(
            TITLE_CONSTANTS,
            text,
            count=1,
        )

        if count != 1:
            raise RuntimeError(
                "No se pudieron actualizar las "
                "constantes de metadata_parser.py."
            )

    if (
        "if style == SPANISH_TITLE_STYLE:"
        not in text
    ):
        pattern = re.compile(
            r'        if style == TITLE_STYLE:\n'
            r'.*?'
            r'        if style == AUTHORS_STYLE:',
            flags=re.DOTALL,
        )

        text, count = pattern.subn(
            TITLE_BRANCH,
            text,
            count=1,
        )

        if count != 1:
            raise RuntimeError(
                "No se pudo actualizar la lectura "
                "de títulos en metadata_parser.py."
            )

    METADATA_PATH.write_text(
        text,
        encoding="utf-8",
    )


def patch_body() -> None:
    backup_once(
        BODY_PATH,
        BODY_BACKUP,
    )

    text = BODY_PATH.read_text(
        encoding="utf-8"
    )

    if '"Título_Español"' not in text:
        pattern = re.compile(
            r'METADATA_STYLES = \{\n'
            r'.*?'
            r'\}\n',
            flags=re.DOTALL,
        )

        text, count = pattern.subn(
            BODY_STYLES,
            text,
            count=1,
        )

        if count != 1:
            raise RuntimeError(
                "No se pudo actualizar "
                "METADATA_STYLES en body_parser.py."
            )

    if (
        "or style in IGNORED_TEMPLATE_STYLES"
        not in text
    ):
        old = '''        if style in METADATA_STYLES:
            continue
'''

        if old not in text:
            raise RuntimeError(
                "No se encontró el filtro de estilos "
                "en body_parser.py."
            )

        text = text.replace(
            old,
            BODY_SKIP,
            1,
        )

    BODY_PATH.write_text(
        text,
        encoding="utf-8",
    )


def compile_files() -> None:
    py_compile.compile(
        str(METADATA_PATH),
        doraise=True,
    )

    py_compile.compile(
        str(BODY_PATH),
        doraise=True,
    )


def test_2027_metadata() -> None:
    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(
            0,
            str(PROJECT_ROOT),
        )

    from converter.metadata_parser import (
        parse_metadata,
    )

    data = {
        "stem": "9999 Completo",
        "paragraphs": [
            {
                "text": "Texto falso",
                "style": "Instruccion_Plantilla",
            },
            {
                "text": (
                    "Título de prueba en español"
                ),
                "style": "Título_Español",
            },
            {
                "text": "Test title in English",
                "style": "Título_Ingles",
            },
            {
                "text": (
                    "Ana María Pérez-López1"
                ),
                "style": "Autores",
            },
            {
                "text": (
                    "1 Universidad de Prueba "
                    "(México), ana@example.com"
                ),
                "style": "ID_Autores",
            },
            {
                "text": (
                    "Códigos JEL: M10, M12, M15"
                ),
                "style": "ID_Autores",
            },
            {
                "text": "Resumen",
                "style": "Resumen_Abstract",
            },
            {
                "text": "Resumen de prueba.",
                "style": "Resumen_Abstract",
            },
            {
                "text": (
                    "Palabras clave: "
                    "uno, dos, tres"
                ),
                "style": "Resumen_Abstract",
            },
            {
                "text": "Abstract",
                "style": "Resumen_Abstract",
            },
            {
                "text": "Test abstract.",
                "style": "Resumen_Abstract",
            },
            {
                "text": (
                    "Key words: one, two, three"
                ),
                "style": "Resumen_Abstract",
            },
            {
                "text": "Figura 99. Falso",
                "style": "Ejemplo_Plantilla",
            },
        ],
    }

    article = parse_metadata(
        data
    )

    assert article.title == (
        "Título de prueba en español"
    )

    assert article.translated_title == (
        "Test title in English"
    )

    assert len(
        article.authors
    ) == 1

    assert (
        article.authors[0].full_name
        == "Ana María Pérez-López"
    )

    assert (
        article.authors[0].email
        == "ana@example.com"
    )

    assert len(
        article.affiliations
    ) == 1

    assert article.abstract == (
        "Resumen de prueba."
    )

    assert (
        article.translated_abstract
        == "Test abstract."
    )

    assert article.keywords == [
        "uno",
        "dos",
        "tres",
    ]

    assert article.translated_keywords == [
        "one",
        "two",
        "three",
    ]

    assert article.jel_codes == [
        "M10",
        "M12",
        "M15",
    ]


def test_ignored_body_styles() -> None:
    from converter.body_parser import (
        parse_body,
    )

    from models import Article

    article = Article()

    data = {
        "paragraphs": [
            {
                "text": "Título ES",
                "style": "Título_Español",
            },
            {
                "text": "Title EN",
                "style": "Título_Ingles",
            },
            {
                "text": "Introducción",
                "style": "Título_Nivel1",
            },
            {
                "text": "Párrafo real.",
                "style": "Párrafo1",
            },
            {
                "text": "Instrucción falsa.",
                "style": "Instruccion_Plantilla",
            },
            {
                "text": "Referencias",
                "style": "Título_Nivel1",
            },
            {
                "text": "Referencia real.",
                "style": "Referencias",
            },
            {
                "text": "Referencia falsa.",
                "style": "Ejemplo_Plantilla",
            },
        ],
        "tables": [],
    }

    parse_body(
        article,
        data,
    )

    assert len(
        article.sections
    ) == 1

    assert (
        article.sections[0].paragraphs
        == ["Párrafo real."]
    )

    assert len(
        article.references
    ) == 1

    assert (
        article.references[0].raw_text
        == "Referencia real."
    )


def test_legacy_1293() -> None:
    from converter.docx_reader import (
        read_docx,
    )

    from converter.metadata_parser import (
        parse_metadata,
    )

    source = (
        PROJECT_ROOT
        / "input"
        / "1293 Completo.docx"
    )

    if not source.is_file():
        raise FileNotFoundError(
            source
        )

    article = parse_metadata(
        read_docx(
            source
        )
    )

    assert article.title
    assert article.translated_title

    assert len(
        article.authors
    ) == 3

    assert len(
        article.affiliations
    ) == 3

    assert article.abstract
    assert article.translated_abstract


def main() -> None:
    if not METADATA_PATH.is_file():
        raise FileNotFoundError(
            METADATA_PATH
        )

    if not BODY_PATH.is_file():
        raise FileNotFoundError(
            BODY_PATH
        )

    patch_metadata()
    patch_body()
    compile_files()

    test_2027_metadata()
    test_ignored_body_styles()
    test_legacy_1293()

    print(
        "PLANTILLA 2027: SOPORTE ACTIVADO"
    )

    print(
        "TITULO_ESPANOL: OK"
    )

    print(
        "TITULO_INGLES: OK"
    )

    print(
        "ESTILOS IGNORABLES: OK"
    )

    print(
        "BODY PARSER: OK"
    )

    print(
        "COMPATIBILIDAD 1293: OK"
    )

    print(
        "COMPILACION: OK"
    )

    print(
        "QA PLANTILLA 2027: APROBADA"
    )

    print(
        "BACKUP METADATA:",
        METADATA_BACKUP,
    )

    print(
        "BACKUP BODY    :",
        BODY_BACKUP,
    )


if __name__ == "__main__":
    main()