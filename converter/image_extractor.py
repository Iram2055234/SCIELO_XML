import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from converter.image_classifier import (
    get_included_relationship_ids,
)
from models import Figure


MEDIA_TYPE_EXTENSIONS = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/gif": ".gif",
    "image/tiff": ".tiff",
    "image/bmp": ".bmp",
    "image/x-emf": ".emf",
    "image/x-wmf": ".wmf",
    "image/svg+xml": ".svg",
}


def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda a un archivo DOCX.
    """

    path = Path(docx_path).resolve()

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


def relationship_sort_key(
    relationship_item: tuple[str, Any],
) -> tuple[int, str]:
    """
    Ordena las relaciones utilizando el número de rId.

    Ejemplo:
    rId2 aparece antes que rId10.
    """

    relationship_id = relationship_item[0]

    match = re.search(
        r"(\d+)$",
        relationship_id,
    )

    if match:
        return (
            int(match.group(1)),
            relationship_id,
        )

    return (
        0,
        relationship_id,
    )


def get_source_filename(
    image_part: Any,
) -> str:
    """
    Obtiene el nombre original de la imagen
    dentro del paquete DOCX.
    """

    part_name = str(
        getattr(
            image_part,
            "partname",
            "",
        )
    )

    if not part_name:
        return ""

    return Path(
        part_name
    ).name


def get_media_type(
    image_part: Any,
) -> str:
    """
    Obtiene el tipo MIME de la imagen.
    """

    return str(
        getattr(
            image_part,
            "content_type",
            "",
        )
        or ""
    ).strip()


def get_image_extension(
    source_filename: str,
    media_type: str,
) -> str:
    """
    Determina la extensión del archivo extraído.

    Primero utiliza el tipo MIME. Si no existe una
    correspondencia conocida, conserva la extensión
    original del archivo.
    """

    normalized_media_type = (
        media_type.strip().lower()
    )

    known_extension = (
        MEDIA_TYPE_EXTENSIONS.get(
            normalized_media_type
        )
    )

    if known_extension:
        return known_extension

    source_extension = Path(
        source_filename
    ).suffix.lower()

    if source_extension:
        return source_extension

    return ".bin"


def get_image_blob(
    image_part: Any,
    relationship_id: str,
) -> bytes:
    """
    Obtiene los datos binarios de una imagen.
    """

    blob = getattr(
        image_part,
        "blob",
        None,
    )

    if blob is None:
        raise RuntimeError(
            "No fue posible obtener los datos "
            "de la imagen relacionada con "
            f"{relationship_id}."
        )

    return bytes(blob)


def build_figure(
    index: int,
    relationship_id: str,
    image_part: Any,
) -> tuple[Figure, bytes]:
    """
    Construye un objeto Figure y devuelve también
    los datos binarios que deben guardarse.
    """

    source_filename = get_source_filename(
        image_part
    )

    media_type = get_media_type(
        image_part
    )

    extension = get_image_extension(
        source_filename=source_filename,
        media_type=media_type,
    )

    output_filename = (
        f"fig{index}{extension}"
    )

    figure = Figure(
        figure_id=f"fig{index}",
        label=f"Figura {index}",
        caption="",
        alt_text="",
        relationship_id=relationship_id,
        source_filename=source_filename,
        output_filename=output_filename,
        media_type=media_type,
        width=0,
        height=0,
    )

    image_blob = get_image_blob(
        image_part=image_part,
        relationship_id=relationship_id,
    )

    return figure, image_blob


def write_image_file(
    image_blob: bytes,
    output_path: str | Path,
) -> Path:
    """
    Guarda una imagen extraída del DOCX.
    """

    path = Path(
        output_path
    ).resolve()

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    try:
        path.write_bytes(
            image_blob
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible guardar la imagen: "
            f"{path}"
        ) from error

    return path


def extract_docx_images(
    docx_path: str | Path,
    output_directory: str | Path,
) -> list[Figure]:
    """
    Extrae únicamente las imágenes clasificadas
    como figuras científicas.

    Se excluyen los elementos editoriales, como:

    - iconos ORCID;
    - logotipos;
    - imágenes de licencia;
    - elementos decorativos identificados.

    Devuelve una lista de objetos Figure.
    """

    validated_docx_path = validate_docx_path(
        docx_path
    )

    target_directory = Path(
        output_directory
    ).resolve()

    target_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    included_relationship_ids = (
        get_included_relationship_ids(
            validated_docx_path
        )
    )

    if not included_relationship_ids:
        return []

    try:
        document = Document(
            str(validated_docx_path)
        )

    except Exception as error:
        raise RuntimeError(
            "No fue posible abrir el archivo Word "
            f"para extraer sus imágenes: "
            f"{validated_docx_path.name}"
        ) from error

    relationships = sorted(
        document.part.rels.items(),
        key=relationship_sort_key,
    )

    figures: list[Figure] = []

    for relationship_id, relationship in relationships:
        if relationship.reltype != RT.IMAGE:
            continue

        if relationship_id not in included_relationship_ids:
            continue

        image_part = relationship.target_part

        figure, image_blob = build_figure(
            index=len(figures) + 1,
            relationship_id=relationship_id,
            image_part=image_part,
        )

        image_output_path = (
            target_directory
            / figure.output_filename
        )

        write_image_file(
            image_blob=image_blob,
            output_path=image_output_path,
        )

        figures.append(
            figure
        )

    return figures