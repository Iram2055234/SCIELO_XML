import sys
from pathlib import Path

from converter.defaults_validator import (
    DefaultsValidationResult,
    validate_journal_defaults_file,
)
from converter.pipeline import (
    ConversionResult,
    convert_article,
)
from converter.report_generator import (
    generate_batch_report,
)
from models import Article


def find_docx_files(
    input_directory: Path,
) -> list[Path]:
    """
    Localiza todos los archivos DOCX procesables.

    Ignora los archivos temporales de Microsoft Word,
    cuyos nombres comienzan con ~$.
    """

    if not input_directory.exists():
        input_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        raise FileNotFoundError(
            "La carpeta input no existía y fue creada. "
            "Agrega archivos .docx y sus archivos .json."
        )

    files = sorted(
        path
        for path in input_directory.glob("*.docx")
        if (
            path.is_file()
            and not path.name.startswith("~$")
        )
    )

    if not files:
        raise FileNotFoundError(
            "No se encontraron archivos .docx "
            "en la carpeta input."
        )

    return files


def count_subsections(
    article: Article,
) -> int:
    """
    Cuenta las subsecciones del artículo.
    """

    return sum(
        len(section.subsections)
        for section in article.sections
    )


def count_body_paragraphs(
    article: Article,
) -> int:
    """
    Cuenta los párrafos de las secciones
    y subsecciones.
    """

    total = 0

    for section in article.sections:
        total += len(
            section.paragraphs
        )

        for subsection in section.subsections:
            total += len(
                subsection.paragraphs
            )

    return total


def count_orcids(
    article: Article,
) -> int:
    """
    Cuenta los autores que tienen ORCID.
    """

    return sum(
        1
        for author in article.authors
        if author.orcid
    )


def print_defaults_validation_summary(
    validation: DefaultsValidationResult,
) -> None:
    """
    Muestra el resultado de validar
    journal_defaults.json.
    """

    print("\nVALIDACIÓN DE LA CONFIGURACIÓN GENERAL")
    print("=" * 70)

    if validation.valid:
        print(
            "✅ La configuración general "
            "de la revista es válida."
        )
    else:
        print(
            "❌ La configuración general "
            "de la revista contiene errores."
        )

    print(
        f"Errores: "
        f"{len(validation.errors)}"
    )

    print(
        f"Advertencias: "
        f"{len(validation.warnings)}"
    )

    if validation.errors:
        print("\nERRORES DE CONFIGURACIÓN GENERAL")
        print("-" * 70)

        for index, error in enumerate(
            validation.errors,
            start=1,
        ):
            print(
                f"{index}. {error}"
            )

    if validation.warnings:
        print("\nADVERTENCIAS DE CONFIGURACIÓN GENERAL")
        print("-" * 70)

        for index, warning in enumerate(
            validation.warnings,
            start=1,
        ):
            print(
                f"{index}. {warning}"
            )


def print_article_summary(
    article: Article,
) -> None:
    """
    Muestra los principales datos detectados.
    """

    print("\nRESUMEN DEL ARTÍCULO")
    print("-" * 70)

    print(
        f"ID: "
        f"{article.article_id or 'No detectado'}"
    )

    print(
        f"Título: "
        f"{article.title or 'No detectado'}"
    )

    print(
        f"DOI: "
        f"{article.doi or 'No detectado'}"
    )

    print(
        f"Revista: "
        f"{article.journal_title or 'No detectada'}"
    )

    print(
        f"ISSN: "
        f"{article.issn or 'No detectado'}"
    )

    print(
        f"Volumen: "
        f"{article.volume or 'No detectado'}"
    )

    print(
        f"Número: "
        f"{article.issue or 'No detectado'}"
    )

    print(
        f"Páginas: "
        f"{article.page_range or 'No detectadas'}"
    )

    print(
        f"Fecha de publicación: "
        f"{article.published_date or 'No detectada'}"
    )

    print(
        f"Autores: "
        f"{len(article.authors)}"
    )

    print(
        f"Autores con ORCID: "
        f"{count_orcids(article)}"
    )

    print(
        f"Afiliaciones: "
        f"{len(article.affiliations)}"
    )

    print(
        f"Secciones principales: "
        f"{len(article.sections)}"
    )

    print(
        f"Subsecciones: "
        f"{count_subsections(article)}"
    )

    print(
        f"Párrafos del cuerpo: "
        f"{count_body_paragraphs(article)}"
    )

    print(
        f"Tablas: "
        f"{len(article.tables)}"
    )

    print(
        f"Referencias: "
        f"{len(article.references)}"
    )


def print_config_validation_summary(
    result: ConversionResult,
) -> None:
    """
    Muestra el resultado de la validación
    de la configuración combinada.
    """

    if result.config_validation is None:
        return

    validation = result.config_validation

    print("\nVALIDACIÓN DEL JSON")
    print("-" * 70)

    if validation.valid:
        print(
            "✅ El archivo JSON y los valores generales "
            "forman una configuración válida."
        )
    else:
        print(
            "❌ La configuración combinada "
            "contiene errores."
        )

    print(
        f"Errores: "
        f"{len(validation.errors)}"
    )

    print(
        f"Advertencias: "
        f"{len(validation.warnings)}"
    )

    if validation.errors:
        print("\nERRORES DEL JSON")

        for index, error in enumerate(
            validation.errors,
            start=1,
        ):
            print(
                f"{index}. {error}"
            )

    if validation.warnings:
        print("\nADVERTENCIAS DEL JSON")

        for index, warning in enumerate(
            validation.warnings,
            start=1,
        ):
            print(
                f"{index}. {warning}"
            )


def print_consistency_validation_summary(
    result: ConversionResult,
) -> None:
    """
    Muestra el resultado de comparar los autores
    del Word con los autores del JSON.
    """

    if result.consistency_validation is None:
        return

    validation = result.consistency_validation

    print("\nCORRESPONDENCIA WORD–JSON")
    print("-" * 70)

    if validation.valid:
        print(
            "✅ Los autores del Word coinciden "
            "con los autores del JSON."
        )
    else:
        print(
            "❌ Se detectaron diferencias entre "
            "los autores del Word y del JSON."
        )

    print(
        f"Errores: "
        f"{len(validation.errors)}"
    )

    print(
        f"Advertencias: "
        f"{len(validation.warnings)}"
    )

    if validation.errors:
        print(
            "\nERRORES DE CORRESPONDENCIA"
        )

        for index, error in enumerate(
            validation.errors,
            start=1,
        ):
            print(
                f"{index}. {error}"
            )

    if validation.warnings:
        print(
            "\nADVERTENCIAS DE CORRESPONDENCIA"
        )

        for index, warning in enumerate(
            validation.warnings,
            start=1,
        ):
            print(
                f"{index}. {warning}"
            )


def print_xml_validation_summary(
    result: ConversionResult,
) -> None:
    """
    Muestra el resultado de la validación interna
    del XML generado.
    """

    if result.validation is None:
        return

    validation = result.validation

    print("\nVALIDACIÓN DEL XML")
    print("-" * 70)

    if validation.valid:
        print(
            "✅ El XML pasó la validación "
            "estructural básica."
        )
    else:
        print(
            "❌ El XML contiene errores."
        )

    print(
        f"Errores: "
        f"{len(validation.errors)}"
    )

    print(
        f"Advertencias: "
        f"{len(validation.warnings)}"
    )

    if validation.errors:
        print("\nERRORES DEL XML")

        for index, error in enumerate(
            validation.errors,
            start=1,
        ):
            print(
                f"{index}. {error}"
            )

    if validation.warnings:
        print("\nADVERTENCIAS DEL XML")

        for index, warning in enumerate(
            validation.warnings,
            start=1,
        ):
            print(
                f"{index}. {warning}"
            )


def print_success_result(
    source_path: Path,
    result: ConversionResult,
) -> None:
    """
    Muestra el resultado de una conversión exitosa.
    """

    print("\nCONVERSIÓN TERMINADA")
    print("=" * 70)

    print(
        f"Word: "
        f"{result.docx_path or source_path}"
    )

    print(
        f"Configuración: "
        f"{result.config_path or 'No disponible'}"
    )

    print_config_validation_summary(
        result
    )

    print_consistency_validation_summary(
        result
    )

    if result.article is not None:
        print_article_summary(
            result.article
        )

    print("\nARCHIVOS GENERADOS")
    print("-" * 70)

    print(
        f"XML: "
        f"{result.xml_path or 'No generado'}"
    )

    print(
        f"HTML: "
        f"{result.html_path or 'No generado'}"
    )

    print_xml_validation_summary(
        result
    )


def print_error_result(
    source_path: Path,
    result: ConversionResult,
) -> None:
    """
    Muestra los errores de una conversión fallida.
    """

    print("\nCONVERSIÓN FALLIDA")
    print("=" * 70)

    print(
        f"Word: {source_path}"
    )

    if result.config_path is not None:
        print(
            f"Configuración: {result.config_path}"
        )

    print_config_validation_summary(
        result
    )

    print_consistency_validation_summary(
        result
    )

    if not result.errors:
        print(
            "\nNo se recibió información "
            "sobre el error."
        )

        return

    print("\nERRORES DE CONVERSIÓN")
    print("-" * 70)

    for index, error in enumerate(
        result.errors,
        start=1,
    ):
        print(
            f"{index}. {error}"
        )


def print_batch_summary(
    processed_results: list[
        tuple[Path, ConversionResult]
    ],
) -> None:
    """
    Muestra el resumen general del lote.
    """

    successful_results = [
        item
        for item in processed_results
        if item[1].success
    ]

    failed_results = [
        item
        for item in processed_results
        if not item[1].success
    ]

    print("\n" + "=" * 70)
    print("RESUMEN GENERAL")
    print("=" * 70)

    print(
        f"Artículos detectados: "
        f"{len(processed_results)}"
    )

    print(
        f"Conversiones exitosas: "
        f"{len(successful_results)}"
    )

    print(
        f"Conversiones fallidas: "
        f"{len(failed_results)}"
    )

    if successful_results:
        print("\nARCHIVOS CONVERTIDOS")
        print("-" * 70)

        for source_path, _ in successful_results:
            print(
                f"✅ {source_path.name}"
            )

    if failed_results:
        print("\nARCHIVOS CON ERROR")
        print("-" * 70)

        for source_path, result in failed_results:
            first_error = (
                result.errors[0]
                if result.errors
                else "Error desconocido"
            )

            print(
                f"❌ {source_path.name}: "
                f"{first_error}"
            )


def print_report_summary(
    report_path: Path,
) -> None:
    """
    Muestra la ubicación del reporte generado.
    """

    print("\nREPORTE DE CONVERSIÓN")
    print("=" * 70)

    print(
        "✅ El reporte del lote fue generado."
    )

    print(
        f"Archivo: {report_path}"
    )


def main() -> int:
    """
    Valida la configuración general, procesa todos
    los artículos y genera un reporte del lote.
    """

    project_directory = (
        Path(__file__).resolve().parent
    )

    input_directory = (
        project_directory
        / "input"
    )

    defaults_path = (
        project_directory
        / "config"
        / "journal_defaults.json"
    )

    try:
        defaults_validation = (
            validate_journal_defaults_file(
                defaults_path
            )
        )

    except (
        FileNotFoundError,
        ValueError,
        RuntimeError,
    ) as error:
        print("\nERROR DE CONFIGURACIÓN GENERAL")
        print("=" * 70)
        print(error)

        return 1

    print_defaults_validation_summary(
        defaults_validation
    )

    if not defaults_validation.valid:
        print("\n" + "=" * 70)
        print(
            "PROCESAMIENTO DETENIDO"
        )
        print("=" * 70)

        print(
            "Corrige config\\journal_defaults.json "
            "antes de volver a ejecutar el programa."
        )

        return 1

    try:
        docx_files = find_docx_files(
            input_directory
        )

    except FileNotFoundError as error:
        print("\nERROR")
        print("=" * 70)
        print(error)

        return 1

    processed_results: list[
        tuple[Path, ConversionResult]
    ] = []

    total_files = len(
        docx_files
    )

    for index, docx_path in enumerate(
        docx_files,
        start=1,
    ):
        print("\n" + "#" * 70)

        print(
            f"PROCESANDO ARTÍCULO "
            f"{index} DE {total_files}"
        )

        print("#" * 70)

        print(
            f"Archivo: {docx_path.name}"
        )

        result = convert_article(
            docx_path=docx_path,
            project_directory=project_directory,
        )

        processed_results.append(
            (
                docx_path,
                result,
            )
        )

        if result.success:
            print_success_result(
                source_path=docx_path,
                result=result,
            )

        else:
            print_error_result(
                source_path=docx_path,
                result=result,
            )

    print_batch_summary(
        processed_results
    )

    try:
        report_path = generate_batch_report(
            processed_results=processed_results,
            project_directory=project_directory,
            defaults_validation=defaults_validation,
        )

    except (
        ValueError,
        RuntimeError,
        OSError,
    ) as error:
        print("\nERROR AL GENERAR EL REPORTE")
        print("=" * 70)
        print(error)

        return 1

    print_report_summary(
        report_path
    )

    failed_count = sum(
        1
        for _, result in processed_results
        if not result.success
    )

    if failed_count:
        print("\n" + "=" * 70)
        print(
            "PASO 19.2 TERMINÓ CON ERRORES"
        )
        print("=" * 70)

        return 1

    print("\n" + "=" * 70)
    print("PROCESAMIENTO GENERAL COMPLETADO")
    print("=" * 70)

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )