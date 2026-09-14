from __future__ import annotations

import py_compile
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

MAIN = (
    ROOT
    / "office_app"
    / "main.py"
)

BUILDER = (
    ROOT
    / "converter"
    / "office_config_builder.py"
)

LOADER = (
    ROOT
    / "converter"
    / "config_loader.py"
)

EDITABLE_FORM = (
    ROOT
    / "office_app"
    / "editable_authors_form.py"
)

BACKUP_MAIN = (
    ROOT
    / "office_app"
    / "main_antes_ui_37_3.py"
)

BACKUP_BUILDER = (
    ROOT
    / "converter"
    / "office_config_builder_antes_ui_37_3.py"
)

BACKUP_LOADER = (
    ROOT
    / "converter"
    / "config_loader_antes_ui_37_3.py"
)


BUILDER_AUTHOR_FIELDS = '''AUTHOR_FIELDS = {
    "email",
    "orcid",
    "full_name",
    "source_name",
}
'''


BUILDER_BLOCK = '''def get_author_names(
    metadata: OfficeMetadata,
) -> list[str]:
    """
    Obtiene y valida los nombres originales.
    """

    names = [
        clean(author.full_name)
        for author in metadata.authors
    ]

    if any(
        not name
        for name in names
    ):
        raise ValueError(
            "Existe al menos un autor sin nombre."
        )

    normalized = [
        name.casefold()
        for name in names
    ]

    if (
        len(set(normalized))
        != len(normalized)
    ):
        raise ValueError(
            "Existen autores con nombres "
            "duplicados."
        )

    return names


def normalize_author_name(
    value: object,
) -> str:
    return clean(value).casefold()


def validate_author_overrides(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None,
) -> None:
    """
    Valida ediciones y autores manuales.
    """

    if overrides is None:
        return

    if not isinstance(
        overrides,
        dict,
    ):
        raise ValueError(
            "Los cambios de autores deben "
            "ser un diccionario."
        )

    detected_names = get_author_names(
        metadata
    )

    detected_by_normalized = {
        normalize_author_name(name): name
        for name in detected_names
    }

    used_sources: set[str] = set()
    used_final_names: set[str] = set()

    for (
        override_key,
        values,
    ) in overrides.items():
        if not isinstance(
            values,
            dict,
        ):
            raise ValueError(
                f"Los datos de "
                f"'{clean(override_key)}' "
                "deben ser un diccionario."
            )

        unknown_fields = (
            set(values)
            - AUTHOR_FIELDS
        )

        if unknown_fields:
            fields = ", ".join(
                sorted(
                    unknown_fields
                )
            )

            raise ValueError(
                f"El autor "
                f"'{clean(override_key)}' "
                "contiene campos no "
                f"reconocidos: {fields}"
            )

        source_name = clean(
            values.get(
                "source_name",
                "",
            )
        )

        if not source_name:
            normalized_key = (
                normalize_author_name(
                    override_key
                )
            )

            if (
                normalized_key
                in detected_by_normalized
            ):
                source_name = (
                    detected_by_normalized[
                        normalized_key
                    ]
                )

        final_name = clean(
            values.get(
                "full_name",
                override_key,
            )
        )

        if not final_name:
            raise ValueError(
                "Existe un autor sin "
                "nombre completo."
            )

        normalized_final = (
            normalize_author_name(
                final_name
            )
        )

        if (
            normalized_final
            in used_final_names
        ):
            raise ValueError(
                f"El autor '{final_name}' "
                "está duplicado."
            )

        used_final_names.add(
            normalized_final
        )

        if source_name:
            normalized_source = (
                normalize_author_name(
                    source_name
                )
            )

            if (
                normalized_source
                not in detected_by_normalized
            ):
                raise ValueError(
                    "El autor original indicado "
                    "no fue encontrado en el Word: "
                    f"{source_name}"
                )

            if (
                normalized_source
                in used_sources
            ):
                raise ValueError(
                    "Un mismo autor detectado "
                    "fue editado más de una vez: "
                    f"{source_name}"
                )

            used_sources.add(
                normalized_source
            )


def build_authors_section(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None = None,
) -> dict[str, dict[str, str]]:
    """
    Conserva la identidad del autor original
    aunque el nombre sea corregido.
    """

    validate_author_overrides(
        metadata=metadata,
        overrides=overrides,
    )

    detected_names = get_author_names(
        metadata
    )

    detected_by_normalized = {
        normalize_author_name(name): name
        for name in detected_names
    }

    overrides_by_source: dict[
        str,
        dict[str, Any],
    ] = {}

    manual_overrides: list[
        tuple[
            str,
            dict[str, Any],
        ]
    ] = []

    if overrides:
        for (
            override_key,
            values,
        ) in overrides.items():
            source_name = clean(
                values.get(
                    "source_name",
                    "",
                )
            )

            if not source_name:
                normalized_key = (
                    normalize_author_name(
                        override_key
                    )
                )

                source_name = (
                    detected_by_normalized.get(
                        normalized_key,
                        "",
                    )
                )

            if source_name:
                overrides_by_source[
                    normalize_author_name(
                        source_name
                    )
                ] = values

            else:
                manual_overrides.append(
                    (
                        override_key,
                        values,
                    )
                )

    authors: dict[
        str,
        dict[str, str],
    ] = {}

    for author in metadata.authors:
        source_name = clean(
            author.full_name
        )

        values = (
            overrides_by_source.get(
                normalize_author_name(
                    source_name
                ),
                {},
            )
        )

        final_name = clean(
            values.get(
                "full_name",
                source_name,
            )
        ) or source_name

        authors[
            final_name
        ] = {
            "email": clean(
                values.get(
                    "email",
                    author.email,
                )
            ),
            "orcid": clean(
                values.get(
                    "orcid",
                    author.orcid,
                )
            ),
            "source_name": source_name,
        }

    for (
        override_key,
        values,
    ) in manual_overrides:
        final_name = clean(
            values.get(
                "full_name",
                override_key,
            )
        )

        authors[
            final_name
        ] = {
            "email": clean(
                values.get(
                    "email",
                    "",
                )
            ),
            "orcid": clean(
                values.get(
                    "orcid",
                    "",
                )
            ),
            "source_name": "",
        }

    return authors
'''


LOADER_BLOCK = '''def clean_author_name(
    value: object,
) -> str:
    return " ".join(
        str(value or "").split()
    )


def normalize_author_name(
    value: object,
) -> str:
    return clean_author_name(
        value
    ).casefold()


def split_manual_author_name(
    full_name: str,
) -> tuple[str, str]:
    cleaned = clean_author_name(
        full_name
    )

    parts = cleaned.split()

    if not parts:
        return "", ""

    if len(parts) == 1:
        return "", parts[0]

    return (
        " ".join(parts[:-1]),
        parts[-1],
    )


def apply_author_config(
    article: Article,
    authors_config: dict[str, Any],
) -> None:
    """
    Edita el mismo autor cuando existe
    source_name y conserva sus afiliaciones.
    """

    existing_by_name: dict[
        str,
        Author,
    ] = {
        normalize_author_name(
            author.full_name
        ): author
        for author in article.authors
        if clean_author_name(
            author.full_name
        )
    }

    for (
        configured_name,
        author_config,
    ) in authors_config.items():
        if not isinstance(
            author_config,
            dict,
        ):
            continue

        final_name = clean_author_name(
            configured_name
        )

        if not final_name:
            continue

        source_name = clean_author_name(
            author_config.get(
                "source_name",
                "",
            )
        )

        author: Author | None = None

        if source_name:
            author = existing_by_name.get(
                normalize_author_name(
                    source_name
                )
            )

        if author is None:
            author = existing_by_name.get(
                normalize_author_name(
                    final_name
                )
            )

        if author is None:
            (
                given_names,
                surname,
            ) = split_manual_author_name(
                final_name
            )

            author = Author(
                given_names=given_names,
                surname=surname,
                full_name=final_name,
            )

            article.authors.append(
                author
            )

        else:
            old_normalized = (
                normalize_author_name(
                    author.full_name
                )
            )

            (
                given_names,
                surname,
            ) = split_manual_author_name(
                final_name
            )

            author.full_name = (
                final_name
            )

            author.given_names = (
                given_names
            )

            author.surname = surname

            existing_by_name.pop(
                old_normalized,
                None,
            )

        existing_by_name[
            normalize_author_name(
                final_name
            )
        ] = author

        author.orcid = str(
            author_config.get(
                "orcid",
                "",
            )
        ).strip()

        configured_email = str(
            author_config.get(
                "email",
                "",
            )
        ).strip()

        if configured_email:
            author.email = (
                configured_email
            )
'''


MAIN_THEME_BLOCK = '''    def _apply_windows_titlebar_theme_without_withdraw(
        self,
        dark: bool,
    ) -> None:
        """
        Cambia la barra de título de Windows
        sin ocultar la ventana.
        """

        if os.name != "nt":
            return

        try:
            import ctypes

            self.update_idletasks()

            hwnd = (
                ctypes.windll.user32
                .GetParent(
                    self.winfo_id()
                )
            )

            value = ctypes.c_int(
                1 if dark else 0
            )

            for attribute in (
                20,
                19,
            ):
                result = (
                    ctypes.windll.dwmapi
                    .DwmSetWindowAttribute(
                        hwnd,
                        attribute,
                        ctypes.byref(
                            value
                        ),
                        ctypes.sizeof(
                            value
                        ),
                    )
                )

                if result == 0:
                    break

        except Exception:
            pass

    def _restore_window_geometry_after_theme(
        self,
        state_before: str,
        geometry_before: str,
    ) -> None:
        """
        Restaura maximizado o geometría
        después del cambio de apariencia.
        """

        try:
            if state_before == "zoomed":
                self.state(
                    "zoomed"
                )
                return

            if (
                state_before == "normal"
                and geometry_before
            ):
                self.geometry(
                    geometry_before
                )

        except tk.TclError:
            pass

    def toggle_theme(self) -> None:
        """
        Alterna claro/oscuro sin perder
        tamaño, posición ni Windows Snap.
        """

        current_mode = (
            ctk.get_appearance_mode()
        )

        target_dark = (
            current_mode == "Light"
        )

        try:
            self.update_idletasks()

            state_before = (
                self.state()
            )

            geometry_before = (
                self.geometry()
            )

        except tk.TclError:
            state_before = "normal"
            geometry_before = ""

        previous_header_setting = getattr(
            ctk.CTk,
            (
                "_deactivate_windows_"
                "window_header_manipulation"
            ),
            False,
        )

        ctk.CTk._deactivate_windows_window_header_manipulation = True

        try:
            if target_dark:
                set_dark_mode()
            else:
                set_light_mode()

        finally:
            ctk.CTk._deactivate_windows_window_header_manipulation = (
                previous_header_setting
            )

        if self.theme_button is not None:
            self.theme_button.configure(
                text=(
                    "☀"
                    if target_dark
                    else "☾"
                )
            )

        self.status_variable.set(
            (
                "Modo oscuro activado."
                if target_dark
                else "Modo claro activado."
            )
        )

        self._apply_windows_titlebar_theme_without_withdraw(
            dark=target_dark
        )

        self._restore_window_geometry_after_theme(
            state_before=state_before,
            geometry_before=geometry_before,
        )

        self.after(
            40,
            lambda: (
                self._restore_window_geometry_after_theme(
                    state_before,
                    geometry_before,
                )
            ),
        )
'''


NEW_FORM_IMPORT = '''from office_app.editable_authors_form import (
    EditableAuthorsMetadataForm as MetadataForm,
)'''


MANUAL_FORM_IMPORT = '''from office_app.manual_authors_form import (
    ManualAuthorsMetadataForm as MetadataForm,
)'''


ORIGINAL_FORM_IMPORT = (
    "from office_app.metadata_form "
    "import MetadataForm"
)


def backup_once(
    source: Path,
    backup: Path,
) -> None:
    if not backup.exists():
        shutil.copy2(
            source,
            backup,
        )


def replace_between(
    text: str,
    start_marker: str,
    end_marker: str,
    replacement: str,
    label: str,
) -> str:
    start = text.find(
        start_marker
    )

    if start < 0:
        raise RuntimeError(
            f"No se encontró el inicio "
            f"de {label}."
        )

    end = text.find(
        end_marker,
        start,
    )

    if end < 0:
        raise RuntimeError(
            f"No se encontró el final "
            f"de {label}."
        )

    return (
        text[:start]
        + replacement
        + "\n\n"
        + text[end:]
    )


def patch_builder() -> None:
    backup_once(
        BUILDER,
        BACKUP_BUILDER,
    )

    text = BUILDER.read_text(
        encoding="utf-8"
    )

    start = text.find(
        "AUTHOR_FIELDS = {"
    )

    if start < 0:
        raise RuntimeError(
            "No se encontró AUTHOR_FIELDS."
        )

    end = text.find(
        "}\n",
        start,
    )

    if end < 0:
        raise RuntimeError(
            "No terminó AUTHOR_FIELDS."
        )

    end += 2

    text = (
        text[:start]
        + BUILDER_AUTHOR_FIELDS
        + text[end:]
    )

    text = replace_between(
        text=text,
        start_marker=(
            "def get_author_names("
        ),
        end_marker=(
            "def build_office_article_config("
        ),
        replacement=(
            BUILDER_BLOCK.rstrip()
        ),
        label=(
            "autores de "
            "office_config_builder.py"
        ),
    )

    BUILDER.write_text(
        text,
        encoding="utf-8",
    )


def patch_loader() -> None:
    backup_once(
        LOADER,
        BACKUP_LOADER,
    )

    text = LOADER.read_text(
        encoding="utf-8"
    )

    if (
        "from models import Article, Author"
        not in text
    ):
        if (
            "from models import Article\n"
            in text
        ):
            text = text.replace(
                "from models import Article\n",
                (
                    "from models import "
                    "Article, Author\n"
                ),
                1,
            )
        else:
            raise RuntimeError(
                "No se encontró import "
                "de Article."
            )

    if (
        "def clean_author_name("
        in text
    ):
        start_marker = (
            "def clean_author_name("
        )
    else:
        start_marker = (
            "def apply_author_config("
        )

    text = replace_between(
        text=text,
        start_marker=start_marker,
        end_marker=(
            "def apply_publication_config("
        ),
        replacement=(
            LOADER_BLOCK.rstrip()
        ),
        label=(
            "autores de "
            "config_loader.py"
        ),
    )

    LOADER.write_text(
        text,
        encoding="utf-8",
    )


def patch_main() -> None:
    backup_once(
        MAIN,
        BACKUP_MAIN,
    )

    text = MAIN.read_text(
        encoding="utf-8"
    )

    if NEW_FORM_IMPORT not in text:
        if MANUAL_FORM_IMPORT in text:
            text = text.replace(
                MANUAL_FORM_IMPORT,
                NEW_FORM_IMPORT,
                1,
            )

        elif (
            ORIGINAL_FORM_IMPORT
            in text
        ):
            text = text.replace(
                ORIGINAL_FORM_IMPORT,
                NEW_FORM_IMPORT,
                1,
            )

        else:
            raise RuntimeError(
                "No se encontró la "
                "importación actual "
                "de MetadataForm."
            )

    text = replace_between(
        text=text,
        start_marker=(
            "    def toggle_theme("
        ),
        end_marker=(
            "    def _build_stepper("
        ),
        replacement=(
            MAIN_THEME_BLOCK.rstrip()
        ),
        label="toggle_theme de main.py",
    )

    MAIN.write_text(
        text,
        encoding="utf-8",
    )


def compile_files() -> None:
    for path in (
        EDITABLE_FORM,
        BUILDER,
        LOADER,
        MAIN,
    ):
        py_compile.compile(
            str(path),
            doraise=True,
        )


def main() -> None:
    for path in (
        MAIN,
        BUILDER,
        LOADER,
        EDITABLE_FORM,
    ):
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

    patch_builder()
    patch_loader()
    patch_main()
    compile_files()

    print(
        "AUTOR DETECTADO EDITABLE: ACTIVADO"
    )

    print(
        "IDENTIDAD source_name: OK"
    )

    print(
        "ALTA MANUAL: CONSERVADA"
    )

    print(
        "TEMA SIN WITHDRAW: ACTIVADO"
    )

    print(
        "RESTAURACION SNAP/GEOMETRIA: ACTIVADA"
    )

    print(
        "COMPILACION: OK"
    )

    print(
        "PASO 37.3 CORRECCION TECNICA: APROBADA"
    )

    print(
        "BACKUP MAIN   :",
        BACKUP_MAIN,
    )

    print(
        "BACKUP BUILDER:",
        BACKUP_BUILDER,
    )

    print(
        "BACKUP LOADER :",
        BACKUP_LOADER,
    )


if __name__ == "__main__":
    main()