from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class WordPdfExportResult:
    """
    Resultado de una exportación DOCX -> PDF
    mediante Microsoft Word.
    """

    success: bool = False

    source_docx: Path | None = None
    pdf_path: Path | None = None

    errors: list[str] = field(
        default_factory=list
    )


def validate_docx(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que el archivo de entrada exista
    y sea un documento DOCX.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.is_file():
        raise FileNotFoundError(
            f"No se encontró el documento Word: {path}"
        )

    if path.suffix.casefold() != ".docx":
        raise ValueError(
            "El archivo de entrada debe tener "
            "extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede exportar un archivo "
            "temporal de Microsoft Word."
        )

    return path


def prepare_pdf_path(
    docx_path: Path,
    pdf_path: str | Path | None,
) -> Path:
    """
    Determina la ruta PDF de salida.
    """

    if pdf_path is None:
        target = (
            docx_path.parent
            / f"{docx_path.stem}.pdf"
        )

    else:
        target = Path(
            pdf_path
        ).resolve()

    if target.suffix.casefold() != ".pdf":
        raise ValueError(
            "La ruta de salida debe terminar "
            "con extensión .pdf."
        )

    target.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    return target


def export_docx_to_pdf(
    docx_path: str | Path,
    pdf_path: str | Path | None = None,
    overwrite: bool = False,
) -> WordPdfExportResult:
    """
    Exporta un DOCX como PDF utilizando
    Microsoft Word.

    La exportación equivale conceptualmente a:

    Word
    -> Archivo
    -> Exportar / Guardar como PDF

    Word se ejecuta de forma invisible y el DOCX
    se abre únicamente para lectura.
    """

    result = WordPdfExportResult()

    word = None
    document = None
    pythoncom = None
    com_initialized = False

    try:
        source = validate_docx(
            docx_path
        )

        target = prepare_pdf_path(
            source,
            pdf_path,
        )

        result.source_docx = source
        result.pdf_path = target

        if target.exists():
            if not overwrite:
                raise FileExistsError(
                    "El PDF ya existe y no se autorizó "
                    f"reemplazarlo: {target}"
                )

            target.unlink()

        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        com_initialized = True

        word = win32com.client.DispatchEx(
            "Word.Application"
        )

        word.Visible = False

        # 0 = wdAlertsNone
        word.DisplayAlerts = 0

        document = word.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=False,
        )

        # ExportAsFixedFormat:
        #
        # ExportFormat = 17 -> wdExportFormatPDF
        # OptimizeFor = 0   -> impresión estándar
        # Range = 0         -> documento completo
        # Item = 0          -> contenido del documento
        # CreateBookmarks = 1 -> encabezados de Word
        document.ExportAsFixedFormat(
            OutputFileName=str(target),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            From=1,
            To=1,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )

        if not target.is_file():
            raise RuntimeError(
                "Microsoft Word terminó la operación, "
                "pero no se encontró el PDF generado."
            )

        if target.stat().st_size <= 0:
            raise RuntimeError(
                "El PDF generado está vacío."
            )

        result.success = True

        return result

    except Exception as error:
        result.success = False

        result.errors.append(
            f"{type(error).__name__}: {error}"
        )

        return result

    finally:
        if document is not None:
            try:
                # False = no guardar cambios.
                document.Close(
                    SaveChanges=False
                )
            except Exception:
                pass

        if word is not None:
            try:
                word.Quit(
                    SaveChanges=False
                )
            except Exception:
                try:
                    word.Quit()
                except Exception:
                    pass

        if (
            pythoncom is not None
            and com_initialized
        ):
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass