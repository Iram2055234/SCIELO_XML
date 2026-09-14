import shutil
from dataclasses import dataclass, field
from pathlib import Path

from converter.body_parser import parse_body
from converter.config_loader import (
    apply_publication_config,
    get_config_path_from_docx,
)
from converter.config_merger import (
    load_merged_publication_config,
)
from converter.config_template import (
    generate_config_template,
)
from converter.config_validator import (
    ConfigValidationResult,
    validate_publication_config,
)
from converter.consistency_validator import (
    ConsistencyValidationResult,
    validate_article_config_consistency,
)
from converter.custom_meta_normalizer import (
    normalize_custom_meta_groups,
)
from converter.docx_reader import read_docx
from converter.editorial_xml_generator import (
    integrate_editorial_metadata_into_xml,
)
from converter.figure_xml_generator import (
    integrate_figures_into_xml,
)
from converter.figure_position_integrator import (
    move_figures_to_original_positions,
)
from converter.html_generator import generate_html
from converter.scientific_figure_builder import (
    build_scientific_figures,
)
from converter.jats_generator import generate_jats_xml
from converter.metadata_parser import parse_metadata
from converter.sidebar_extractor import (
    SidebarMetadata,
    extract_sidebar_metadata,
)
from converter.table_extractor import extract_docx_tables
from converter.table_position_integrator import (
    move_tables_to_original_positions,
)
from converter.table_xml_generator import (
    integrate_tables_into_xml,
)
from converter.validator import (
    ValidationResult,
    validate_xml,
)
from models import Article


@dataclass
class ConversionResult:
    """
    Representa el resultado completo de una conversión.
    """

    success: bool = False

    docx_path: Path | None = None
    config_path: Path | None = None
    xml_path: Path | None = None
    html_path: Path | None = None
    media_directory: Path | None = None

    article: Article | None = None

    config_validation: (
        ConfigValidationResult | None
    ) = None

    consistency_validation: (
        ConsistencyValidationResult | None
    ) = None

    validation: ValidationResult | None = None

    errors: list[str] = field(
        default_factory=list
    )


def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que la entrada sea un archivo DOCX válido.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo Word: {path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta indicada no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo de entrada debe tener "
            "extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se pueden procesar archivos "
            "temporales de Microsoft Word."
        )

    return path


def prepare_output_directories(
    project_directory: str | Path,
    article_stem: str,
) -> tuple[Path, Path, Path]:
    """
    Prepara las carpetas XML, HTML y multimedia.

    La carpeta multimedia del artículo se limpia antes
    de cada conversión para evitar archivos obsoletos.
    """

    project_path = Path(
        project_directory
    ).resolve()

    xml_directory = (
        project_path
        / "output"
        / "xml"
    )

    html_directory = (
        project_path
        / "output"
        / "html"
    )

    media_directory = (
        project_path
        / "output"
        / "media"
        / article_stem
    )

    xml_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    html_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    if media_directory.exists():
        shutil.rmtree(
            media_directory
        )

    media_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    return (
        xml_directory,
        html_directory,
        media_directory,
    )


def get_xslt_path(
    project_directory: str | Path,
) -> Path:
    """
    Obtiene y valida la plantilla XSLT del proyecto.
    """

    path = (
        Path(project_directory).resolve()
        / "templates"
        / "article.xsl"
    )

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró la plantilla XSLT: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta de la plantilla XSLT "
            "no corresponde a un archivo."
        )

    return path


def build_article_from_docx(
    docx_path: Path,
) -> Article:
    """
    Lee el Word y construye el objeto Article
    sin aplicar todavía la configuración JSON.
    """

    document_data = read_docx(
        docx_path
    )

    article = parse_metadata(
        document_data
    )

    article = parse_body(
        article=article,
        document_data=document_data,
    )

    return article


def apply_sidebar_metadata(
    article: Article,
    sidebar_metadata: SidebarMetadata,
) -> Article:
    """
    Incorpora al artículo la información editorial
    extraída desde la columna lateral Cómo citar.
    """

    has_sidebar_content = any(
        [
            sidebar_metadata.citation_text,
            sidebar_metadata.peer_review_text,
            sidebar_metadata.received_text,
            sidebar_metadata.accepted_text,
            sidebar_metadata.published_text,
            sidebar_metadata.copyright_statement,
            sidebar_metadata.license_text,
            sidebar_metadata.license_url,
        ]
    )

    if not has_sidebar_content:
        return article

    if sidebar_metadata.citation_title:
        article.citation_title = (
            sidebar_metadata.citation_title
        )

    if sidebar_metadata.citation_text:
        article.citation_text = (
            sidebar_metadata.citation_text
        )

    if sidebar_metadata.peer_review_text:
        article.peer_review_text = (
            sidebar_metadata.peer_review_text
        )

    if sidebar_metadata.received_text:
        article.received_text = (
            sidebar_metadata.received_text
        )

    if sidebar_metadata.accepted_text:
        article.accepted_text = (
            sidebar_metadata.accepted_text
        )

    if sidebar_metadata.published_text:
        article.published_text = (
            sidebar_metadata.published_text
        )

    if sidebar_metadata.copyright_statement:
        article.copyright_statement = (
            sidebar_metadata.copyright_statement
        )

    if sidebar_metadata.license_text:
        article.license_text = (
            sidebar_metadata.license_text
        )

    if sidebar_metadata.license_url:
        article.license_url = (
            sidebar_metadata.license_url
        )

    return article


def create_missing_config_template(
    docx_path: Path,
    config_path: Path,
) -> ConversionResult:
    """
    Crea una plantilla JSON cuando el Word todavía
    no tiene archivo de configuración.
    """

    generated_config_path = generate_config_template(
        docx_path=docx_path,
        output_path=config_path,
        overwrite=False,
    )

    result = ConversionResult(
        success=False,
        docx_path=docx_path,
        config_path=generated_config_path,
    )

    result.errors.append(
        "No se encontró el JSON correspondiente. "
        "Se creó automáticamente una plantilla en: "
        f"{generated_config_path}. "
        "Completa los metadatos particulares del artículo "
        "y vuelve a ejecutar el programa."
    )

    return result


def add_config_validation_errors(
    result: ConversionResult,
    config_validation: ConfigValidationResult,
) -> None:
    """
    Agrega los errores encontrados en la configuración
    combinada.
    """

    for error in config_validation.errors:
        result.errors.append(
            f"Configuración JSON: {error}"
        )


def add_consistency_validation_errors(
    result: ConversionResult,
    consistency_validation: ConsistencyValidationResult,
) -> None:
    """
    Agrega los errores de correspondencia entre
    el Word y la configuración combinada.
    """

    for error in consistency_validation.errors:
        result.errors.append(
            f"Correspondencia Word–JSON: {error}"
        )


def convert_article(
    docx_path: str | Path,
    project_directory: str | Path,
) -> ConversionResult:
    """
    Convierte un artículo Word a XML JATS y HTML.

    Flujo:

    1. Valida el Word y el JSON.
    2. Construye el objeto Article.
    3. Aplica la configuración editorial.
    4. Extrae la columna Cómo citar.
    5. Extrae y normaliza las tablas científicas.
    6. Clasifica y extrae figuras científicas.
    7. Genera el XML JATS.
    8. Inserta los metadatos editoriales.
    9. Unifica los grupos de metadatos personalizados.
    10. Inserta las figuras científicas.
    11. Inserta las tablas científicas.
    12. Mueve las tablas a sus posiciones originales.
    13. Valida el XML.
    14. Genera el HTML.
    """

    result = ConversionResult()

    try:
        validated_docx_path = validate_docx_path(
            docx_path
        )

        project_path = Path(
            project_directory
        ).resolve()

        config_path = get_config_path_from_docx(
            validated_docx_path
        )

        result.docx_path = validated_docx_path
        result.config_path = config_path

        if not config_path.exists():
            return create_missing_config_template(
                docx_path=validated_docx_path,
                config_path=config_path,
            )

        merged_config = load_merged_publication_config(
            config_path=config_path
        )

        config_validation = validate_publication_config(
            merged_config
        )

        result.config_validation = (
            config_validation
        )

        if not config_validation.valid:
            add_config_validation_errors(
                result=result,
                config_validation=config_validation,
            )

            result.success = False
            return result

        article = build_article_from_docx(
            validated_docx_path
        )

        result.article = article

        consistency_validation = (
            validate_article_config_consistency(
                article=article,
                config=merged_config,
            )
        )

        result.consistency_validation = (
            consistency_validation
        )

        if not consistency_validation.valid:
            add_consistency_validation_errors(
                result=result,
                consistency_validation=consistency_validation,
            )

            result.success = False
            return result

        article = apply_publication_config(
            article=article,
            config=merged_config,
        )

        sidebar_metadata = extract_sidebar_metadata(
            validated_docx_path
        )

        article = apply_sidebar_metadata(
            article=article,
            sidebar_metadata=sidebar_metadata,
        )

        article.tables = extract_docx_tables(
            validated_docx_path
        )

        (
            xml_directory,
            html_directory,
            media_directory,
        ) = prepare_output_directories(
            project_directory=project_path,
            article_stem=validated_docx_path.stem,
        )

        article.figures = build_scientific_figures(
            docx_path=validated_docx_path,
            output_directory=media_directory,
        )

        result.article = article
        result.media_directory = media_directory

        xslt_path = get_xslt_path(
            project_path
        )

        xml_path = (
            xml_directory
            / f"{validated_docx_path.stem}.xml"
        )

        html_path = (
            html_directory
            / f"{validated_docx_path.stem}.html"
        )

        generated_xml_path = generate_jats_xml(
            article=article,
            output_path=xml_path,
        )

        generated_xml_path = (
            integrate_editorial_metadata_into_xml(
                xml_path=generated_xml_path,
                article=article,
            )
        )

        generated_xml_path = normalize_custom_meta_groups(
            xml_path=generated_xml_path,
        )

        generated_xml_path = integrate_figures_into_xml(
            xml_path=generated_xml_path,
            figures=article.figures,
            article_stem=validated_docx_path.stem,
        )

        generated_xml_path = integrate_tables_into_xml(
            xml_path=generated_xml_path,
            tables=article.tables,
        )

        generated_xml_path = (
            move_figures_to_original_positions(
                xml_path=generated_xml_path,
                docx_path=validated_docx_path,
            )
        )

        generated_xml_path = (
            move_tables_to_original_positions(
                xml_path=generated_xml_path,
                docx_path=validated_docx_path,
            )
        )

        validation_result = validate_xml(
            generated_xml_path
        )

        result.xml_path = generated_xml_path
        result.validation = validation_result

        if not validation_result.valid:
            result.errors.extend(
                validation_result.errors
            )

            result.success = False
            return result

        generated_html_path = generate_html(
            xml_path=generated_xml_path,
            xslt_path=xslt_path,
            output_path=html_path,
        )

        result.html_path = generated_html_path
        result.success = True

        return result

    except (
        FileExistsError,
        FileNotFoundError,
        ValueError,
        RuntimeError,
    ) as error:
        result.errors.append(
            str(error)
        )

        result.success = False
        return result

    except Exception as error:
        result.errors.append(
            f"{type(error).__name__}: {error}"
        )

        result.success = False
        return result