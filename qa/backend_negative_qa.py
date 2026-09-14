from __future__ import annotations

import hashlib
import shutil
import sys
from pathlib import Path

from converter.office_result_manager import (
    publish_office_result,
)
from office_app.conversion_runner import (
    OfficeGuiConversionResult,
    run_office_conversion,
)


PROJECT_ROOT = (
    Path(__file__)
    .resolve()
    .parent
    .parent
)

SOURCE_DOCX = (
    PROJECT_ROOT
    / "input"
    / "1293 Completo.docx"
)

QA_ROOT = (
    PROJECT_ROOT
    / "output"
    / "qa_backend_negative"
).resolve()

ARTICLE_NAME = "1293 Completo"


class QaFailure(
    AssertionError
):
    """
    Error controlado de QA.
    """


def check(
    condition: bool,
    message: str,
) -> None:
    """
    Falla cuando una condición no se cumple.
    """

    if not condition:
        raise QaFailure(
            message
        )


def print_ok(
    label: str,
    detail: str = "",
) -> None:
    """
    Imprime una prueba aprobada.
    """

    if detail:
        print(
            f"[OK] {label}: {detail}"
        )
    else:
        print(
            f"[OK] {label}"
        )


def sha256_file(
    path: Path,
) -> str:
    """
    Calcula SHA-256 de un archivo.
    """

    digest = hashlib.sha256()

    with path.open(
        "rb"
    ) as file:
        while True:
            block = file.read(
                1024 * 1024
            )

            if not block:
                break

            digest.update(
                block
            )

    return digest.hexdigest()


def cleanup_qa() -> None:
    """
    Limpia exclusivamente esta salida QA.
    """

    if QA_ROOT.exists():
        shutil.rmtree(
            QA_ROOT
        )

    QA_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )


def assert_failed_conversion(
    result: OfficeGuiConversionResult,
    label: str,
) -> None:
    """
    Comprueba que una conversión haya fallado
    de manera controlada.
    """

    check(
        not result.success,
        (
            f"{label}: el backend indicó éxito "
            "cuando debía rechazar la entrada."
        ),
    )

    check(
        bool(
            result.errors
        ),
        (
            f"{label}: el backend falló pero no "
            "devolvió ningún mensaje de error."
        ),
    )

    print_ok(
        label,
        result.errors[0],
    )


def test_missing_docx() -> None:
    """
    DOCX inexistente.
    """

    missing_path = (
        QA_ROOT
        / "no_existe.docx"
    )

    result = run_office_conversion(
        docx_path=missing_path,
        results_root=QA_ROOT,
    )

    assert_failed_conversion(
        result,
        "DOCX inexistente rechazado",
    )


def test_temporary_word_file() -> None:
    """
    Archivo temporal ~$.docx.
    """

    temporary_docx = (
        QA_ROOT
        / "~$articulo_temporal.docx"
    )

    temporary_docx.write_bytes(
        b"QA temporal"
    )

    try:
        result = run_office_conversion(
            docx_path=temporary_docx,
            results_root=QA_ROOT,
        )

        assert_failed_conversion(
            result,
            "Archivo temporal de Word rechazado",
        )

    finally:
        if temporary_docx.exists():
            temporary_docx.unlink()


def test_invalid_extension() -> None:
    """
    Archivo que no es DOCX.
    """

    invalid_file = (
        QA_ROOT
        / "articulo.txt"
    )

    invalid_file.write_text(
        "Archivo de QA",
        encoding="utf-8",
    )

    try:
        result = run_office_conversion(
            docx_path=invalid_file,
            results_root=QA_ROOT,
        )

        assert_failed_conversion(
            result,
            "Extensión incorrecta rechazada",
        )

    finally:
        if invalid_file.exists():
            invalid_file.unlink()


def create_good_baseline(
) -> OfficeGuiConversionResult:
    """
    Crea una conversión válida que utilizaremos
    para comprobar que posteriores fallos no la
    destruyan.
    """

    print()
    print(
        "-" * 70
    )

    print(
        "CREANDO RESULTADO BASE VÁLIDO"
    )

    print(
        "-" * 70
    )

    result = run_office_conversion(
        docx_path=SOURCE_DOCX,
        results_root=QA_ROOT,
        cleanup_run=False,
    )

    check(
        result.success,
        (
            "No fue posible crear el resultado "
            "base válido: "
            + " | ".join(
                result.errors
            )
        ),
    )

    check(
        result.result_directory is not None,
        "La conversión base no tiene carpeta final.",
    )

    check(
        result.xml_path is not None
        and result.xml_path.is_file(),
        "Falta XML en la conversión base.",
    )

    check(
        result.html_path is not None
        and result.html_path.is_file(),
        "Falta HTML en la conversión base.",
    )

    check(
        result.pdf_path is not None
        and result.pdf_path.is_file(),
        "Falta PDF en la conversión base.",
    )

    check(
        result.epub_path is not None
        and result.epub_path.is_file(),
        "Falta EPUB en la conversión base.",
    )

    check(
        result.media_directory is not None
        and result.media_directory.is_dir(),
        "Faltan imágenes en la conversión base.",
    )

    print_ok(
        "Resultado base",
        "XML + HTML + PDF + EPUB + imágenes",
    )

    return result


def snapshot_good_result(
    result: OfficeGuiConversionResult,
) -> dict[str, str]:
    """
    Guarda hashes de los cuatro formatos
    principales.
    """

    check(
        result.xml_path is not None,
        "Falta ruta XML.",
    )

    check(
        result.html_path is not None,
        "Falta ruta HTML.",
    )

    check(
        result.pdf_path is not None,
        "Falta ruta PDF.",
    )

    check(
        result.epub_path is not None,
        "Falta ruta EPUB.",
    )

    return {
        "xml": sha256_file(
            result.xml_path
        ),
        "html": sha256_file(
            result.html_path
        ),
        "pdf": sha256_file(
            result.pdf_path
        ),
        "epub": sha256_file(
            result.epub_path
        ),
    }


def verify_good_result_unchanged(
    result: OfficeGuiConversionResult,
    snapshot: dict[str, str],
    label: str,
) -> None:
    """
    Comprueba que una operación fallida no haya
    modificado el resultado válido.
    """

    paths = {
        "xml": result.xml_path,
        "html": result.html_path,
        "pdf": result.pdf_path,
        "epub": result.epub_path,
    }

    for (
        key,
        path,
    ) in paths.items():
        check(
            path is not None,
            (
                f"{label}: desapareció la ruta "
                f"{key.upper()}."
            ),
        )

        check(
            path.is_file(),
            (
                f"{label}: desapareció el archivo "
                f"{key.upper()}."
            ),
        )

        current_hash = sha256_file(
            path
        )

        check(
            current_hash
            == snapshot[
                key
            ],
            (
                f"{label}: el archivo "
                f"{key.upper()} fue modificado."
            ),
        )

    print_ok(
        label,
        "resultado anterior intacto",
    )


def test_invalid_metadata(
    baseline: OfficeGuiConversionResult,
    snapshot: dict[str, str],
) -> None:
    """
    Intenta convertir con un DOI claramente
    inválido.
    """

    result = run_office_conversion(
        docx_path=SOURCE_DOCX,
        article_overrides={
            "doi": "DOI_INVALIDO_QA",
        },
        results_root=QA_ROOT,
        cleanup_run=True,
    )

    assert_failed_conversion(
        result,
        "Metadatos inválidos rechazados",
    )

    verify_good_result_unchanged(
        baseline,
        snapshot,
        (
            "Resultado protegido tras "
            "metadatos inválidos"
        ),
    )


def test_missing_pdf_publication(
    baseline: OfficeGuiConversionResult,
    snapshot: dict[str, str],
) -> None:
    """
    Intenta publicar usando un PDF inexistente.
    """

    service_result = (
        baseline.service_result
    )

    check(
        service_result is not None,
        (
            "El resultado base no conserva "
            "service_result."
        ),
    )

    missing_pdf = (
        QA_ROOT
        / "pdf_que_no_existe.pdf"
    )

    published = publish_office_result(
        service_result=service_result,
        results_root=QA_ROOT,
        pdf_source_path=missing_pdf,
        cleanup_run=False,
    )

    check(
        not published.success,
        (
            "El administrador aceptó un PDF "
            "inexistente."
        ),
    )

    check(
        bool(
            published.errors
        ),
        (
            "El PDF inexistente fue rechazado "
            "sin devolver mensaje."
        ),
    )

    print_ok(
        "PDF inexistente rechazado",
        published.errors[0],
    )

    verify_good_result_unchanged(
        baseline,
        snapshot,
        (
            "Resultado protegido tras "
            "PDF inexistente"
        ),
    )


def test_corrupt_epub_publication(
    baseline: OfficeGuiConversionResult,
    snapshot: dict[str, str],
) -> None:
    """
    Intenta publicar un archivo .epub que no
    contiene un EPUB real.
    """

    service_result = (
        baseline.service_result
    )

    check(
        service_result is not None,
        (
            "El resultado base no conserva "
            "service_result."
        ),
    )

    corrupt_epub = (
        QA_ROOT
        / "epub_corrupto.epub"
    )

    corrupt_epub.write_bytes(
        b"ESTO NO ES UN EPUB VALIDO"
    )

    try:
        published = publish_office_result(
            service_result=service_result,
            results_root=QA_ROOT,
            epub_source_path=corrupt_epub,
            cleanup_run=False,
        )

        check(
            not published.success,
            (
                "El administrador aceptó un "
                "EPUB corrupto."
            ),
        )

        check(
            bool(
                published.errors
            ),
            (
                "El EPUB corrupto fue rechazado "
                "sin devolver mensaje."
            ),
        )

        print_ok(
            "EPUB corrupto rechazado",
            published.errors[0],
        )

        verify_good_result_unchanged(
            baseline,
            snapshot,
            (
                "Resultado protegido tras "
                "EPUB corrupto"
            ),
        )

    finally:
        if corrupt_epub.exists():
            corrupt_epub.unlink()


def test_no_temporary_leftovers() -> None:
    """
    Comprueba que no queden carpetas temporales
    de publicación.
    """

    leftovers = [
        path.name
        for path in QA_ROOT.iterdir()
        if (
            path.name.startswith(
                ".tmp_"
            )
            or path.name.startswith(
                ".backup_"
            )
        )
    ]

    check(
        not leftovers,
        (
            "Quedaron temporales de publicación: "
            + ", ".join(
                leftovers
            )
        ),
    )

    print_ok(
        "Sin temporales residuales"
    )


def cleanup_service_run(
    baseline: OfficeGuiConversionResult,
) -> None:
    """
    Elimina la carpeta técnica que mantuvimos
    solamente para las pruebas de publicación.
    """

    service_result = (
        baseline.service_result
    )

    if service_result is None:
        return

    run_directory = (
        service_result.run_directory
    )

    if (
        run_directory is not None
        and run_directory.exists()
    ):
        shutil.rmtree(
            run_directory,
            ignore_errors=True,
        )


def main() -> int:
    """
    Ejecuta el QA negativo.
    """

    print()
    print(
        "=" * 70
    )

    print(
        "QA DE ERRORES CONTROLADOS - BACKEND SCIELO FACPyA"
    )

    print(
        "=" * 70
    )

    baseline: OfficeGuiConversionResult | None = None

    try:
        check(
            SOURCE_DOCX.is_file(),
            (
                "No se encontró el documento "
                f"de prueba: {SOURCE_DOCX}"
            ),
        )

        cleanup_qa()

        test_missing_docx()

        test_temporary_word_file()

        test_invalid_extension()

        baseline = create_good_baseline()

        snapshot = snapshot_good_result(
            baseline
        )

        marker_path = (
            baseline.result_directory
            / "QA_PROTECTED_RESULT.txt"
        )

        marker_path.write_text(
            (
                "Este resultado debe sobrevivir "
                "a todas las pruebas negativas."
            ),
            encoding="utf-8",
        )

        test_invalid_metadata(
            baseline,
            snapshot,
        )

        check(
            marker_path.is_file(),
            (
                "Los metadatos inválidos destruyeron "
                "el resultado anterior."
            ),
        )

        test_missing_pdf_publication(
            baseline,
            snapshot,
        )

        check(
            marker_path.is_file(),
            (
                "El intento de PDF inválido destruyó "
                "el resultado anterior."
            ),
        )

        test_corrupt_epub_publication(
            baseline,
            snapshot,
        )

        check(
            marker_path.is_file(),
            (
                "El intento de EPUB corrupto destruyó "
                "el resultado anterior."
            ),
        )

        test_no_temporary_leftovers()

        print()
        print(
            "=" * 70
        )

        print(
            "RESULTADO QA NEGATIVO: APROBADO"
        )

        print(
            "=" * 70
        )

        print(
            "Entradas incorrectas rechazadas:"
        )

        print(
            "  DOCX inexistente      : OK"
        )

        print(
            "  Archivo temporal Word : OK"
        )

        print(
            "  Extensión incorrecta  : OK"
        )

        print(
            "  Metadatos inválidos   : OK"
        )

        print(
            "  PDF inexistente       : OK"
        )

        print(
            "  EPUB corrupto         : OK"
        )

        print()

        print(
            "Protección de resultados:"
        )

        print(
            "  Resultado previo intacto : OK"
        )

        print(
            "  Sin .tmp_ residuales     : OK"
        )

        print(
            "  Sin .backup_ residuales  : OK"
        )

        return 0

    except Exception as error:
        print()
        print(
            "=" * 70
        )

        print(
            "RESULTADO QA NEGATIVO: FALLÓ"
        )

        print(
            "=" * 70
        )

        print(
            f"{type(error).__name__}: {error}"
        )

        return 1

    finally:
        if baseline is not None:
            cleanup_service_run(
                baseline
            )


if __name__ == "__main__":
    sys.exit(
        main()
    )