from pathlib import Path

from lxml import etree


def validate_input_file(
    file_path: str | Path,
    expected_suffix: str,
    description: str,
) -> Path:
    """
    Valida que una ruta corresponda a un archivo existente
    con la extensión esperada.
    """

    path = Path(file_path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"No se encontró {description}: {path}"
        )

    if not path.is_file():
        raise ValueError(
            f"La ruta no corresponde a un archivo: {path}"
        )

    if path.suffix.lower() != expected_suffix.lower():
        raise ValueError(
            f"{description} debe tener extensión "
            f"{expected_suffix}: {path.name}"
        )

    return path


def load_xml_document(
    xml_path: Path,
) -> etree._ElementTree:
    """
    Abre el XML JATS sin intentar descargar la DTD remota.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
    )

    try:
        return etree.parse(
            str(xml_path),
            parser,
        )

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "El XML no está bien formado. "
            f"Línea {error.lineno}: {error.msg}"
        ) from error

    except OSError as error:
        raise RuntimeError(
            f"No fue posible abrir el XML: {xml_path.name}"
        ) from error


def load_xslt_transformer(
    xslt_path: Path,
) -> etree.XSLT:
    """
    Abre y compila la plantilla XSLT.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
    )

    try:
        xslt_document = etree.parse(
            str(xslt_path),
            parser,
        )

        return etree.XSLT(
            xslt_document
        )

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "La plantilla XSLT no está bien formada. "
            f"Línea {error.lineno}: {error.msg}"
        ) from error

    except etree.XSLTParseError as error:
        raise RuntimeError(
            "No fue posible compilar la plantilla XSLT: "
            f"{error}"
        ) from error

    except OSError as error:
        raise RuntimeError(
            "No fue posible abrir la plantilla XSLT: "
            f"{xslt_path.name}"
        ) from error


def apply_xslt_transform(
    xml_document: etree._ElementTree,
    transformer: etree.XSLT,
) -> etree._XSLTResultTree:
    """
    Aplica la transformación XSLT al documento XML.
    """

    try:
        return transformer(
            xml_document
        )

    except etree.XSLTApplyError as error:
        raise RuntimeError(
            "Ocurrió un error durante la transformación "
            f"del XML a HTML: {error}"
        ) from error


def write_html_output(
    html_result: etree._XSLTResultTree,
    output_path: str | Path,
) -> Path:
    """
    Guarda el resultado de la transformación como HTML.
    """

    path = Path(output_path).resolve()

    if path.suffix.lower() not in {
        ".html",
        ".htm",
    }:
        raise ValueError(
            "El archivo de salida debe tener extensión "
            ".html o .htm"
        )

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        html_result.write_output(
            str(path)
        )

    except OSError as error:
        raise RuntimeError(
            f"No fue posible guardar el HTML: {path}"
        ) from error

    return path


def generate_html(
    xml_path: str | Path,
    xslt_path: str | Path,
    output_path: str | Path,
) -> Path:
    """
    Genera un archivo HTML a partir de un XML JATS
    y una plantilla XSLT.

    Flujo:
    1. Valida las rutas.
    2. Abre el XML.
    3. Compila el XSLT.
    4. Aplica la transformación.
    5. Guarda el HTML.
    """

    validated_xml_path = validate_input_file(
        file_path=xml_path,
        expected_suffix=".xml",
        description="el archivo XML",
    )

    validated_xslt_path = validate_input_file(
        file_path=xslt_path,
        expected_suffix=".xsl",
        description="la plantilla XSLT",
    )

    xml_document = load_xml_document(
        validated_xml_path
    )

    transformer = load_xslt_transformer(
        validated_xslt_path
    )

    html_result = apply_xslt_transform(
        xml_document=xml_document,
        transformer=transformer,
    )

    return write_html_output(
        html_result=html_result,
        output_path=output_path,
    )