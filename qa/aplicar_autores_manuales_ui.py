from __future__ import annotations

import py_compile
import shutil
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

MAIN_PATH = (
    PROJECT_ROOT
    / "office_app"
    / "main.py"
)

FORM_PATH = (
    PROJECT_ROOT
    / "office_app"
    / "manual_authors_form.py"
)

MAIN_BACKUP = (
    PROJECT_ROOT
    / "office_app"
    / "main_antes_autores_manual_ui.py"
)


NEW_IMPORT = """from office_app.manual_authors_form import (
    ManualAuthorsMetadataForm as MetadataForm,
)"""


BIND_BLOCK = """    def _bind_metadata_changes(self) -> None:
        \\"\\"\\"
        Detecta cambios en el formulario.
        \\"\\"\\"

        if self.metadata_form is None:
            return

        variables: list[
            tk.StringVar
        ] = []

        variables.extend(
            self.metadata_form
            .article_variables
            .values()
        )

        variables.extend(
            self.metadata_form
            .license_variables
            .values()
        )

        for author_fields in (
            self.metadata_form
            .author_variables
            .values()
        ):
            variables.extend(
                author_fields.values()
            )

        for variable in variables:
            variable.trace_add(
                "write",
                self._on_metadata_changed,
            )

        self.metadata_form.set_change_callback(
            self._on_metadata_changed
        )

"""


VALIDATION_BLOCK = """            manual_author_errors = (
                self.metadata_form
                .validate_manual_authors()
            )

            if manual_author_errors:
                self.current_config_result = None

                self.validation_variable.set(
                    "Hay datos que deben corregirse."
                )

                if self.validation_label is not None:
                    self.validation_label.configure(
                        text_color=ERROR
                    )

                if self.convert_button is not None:
                    self.convert_button.configure(
                        state="disabled",
                        fg_color=BUTTON_DISABLED,
                        text_color=(
                            BUTTON_DISABLED_TEXT
                        ),
                    )

                error_text = "\\n".join(
                    f"• {message}"
                    for message in dict.fromkeys(
                        manual_author_errors
                    )
                )

                ModernDialog.show(
                    self,
                    "Revise los autores",
                    (
                        "Hay información de autores que "
                        "debe completarse antes de "
                        "continuar:\\n\\n"
                        f"{error_text}"
                    ),
                    kind="warning",
                )

                self.status_variable.set(
                    "Revise los autores agregados."
                )

                return None

"""


def replace_between(
    text: str,
    start_marker: str,
    end_marker: str,
    replacement: str,
) -> str:
    start = text.find(
        start_marker
    )

    if start < 0:
        raise RuntimeError(
            "No se encontró el inicio del bloque "
            "que se iba a actualizar."
        )

    end = text.find(
        end_marker,
        start,
    )

    if end < 0:
        raise RuntimeError(
            "No se encontró el final del bloque "
            "que se iba a actualizar."
        )

    return (
        text[:start]
        + replacement
        + text[end:]
    )


def main() -> None:
    if not MAIN_PATH.is_file():
        raise FileNotFoundError(
            MAIN_PATH
        )

    if not FORM_PATH.is_file():
        raise FileNotFoundError(
            "Primero cree manual_authors_form.py: "
            f"{FORM_PATH}"
        )

    if not MAIN_BACKUP.exists():
        shutil.copy2(
            MAIN_PATH,
            MAIN_BACKUP,
        )

    text = MAIN_PATH.read_text(
        encoding="utf-8"
    )

    old_import = (
        "from office_app.metadata_form "
        "import MetadataForm"
    )

    if NEW_IMPORT not in text:
        if old_import not in text:
            raise RuntimeError(
                "No se encontró la importación "
                "esperada de MetadataForm."
            )

        text = text.replace(
            old_import,
            NEW_IMPORT,
            1,
        )

    text = replace_between(
        text=text,
        start_marker=(
            "    def _bind_metadata_changes("
        ),
        end_marker=(
            "    def _on_metadata_changed("
        ),
        replacement=BIND_BLOCK,
    )

    validation_marker = """            ) = self._get_form_overrides()

            result = (
"""

    if (
        "validate_manual_authors()"
        not in text
    ):
        if validation_marker not in text:
            raise RuntimeError(
                "No se encontró el punto esperado "
                "de validación de metadatos."
            )

        replacement = (
            """            ) = self._get_form_overrides()

"""
            + VALIDATION_BLOCK
            + """            result = (
"""
        )

        text = text.replace(
            validation_marker,
            replacement,
            1,
        )

    MAIN_PATH.write_text(
        text,
        encoding="utf-8",
    )

    py_compile.compile(
        str(FORM_PATH),
        doraise=True,
    )

    py_compile.compile(
        str(MAIN_PATH),
        doraise=True,
    )

    final_text = MAIN_PATH.read_text(
        encoding="utf-8"
    )

    checks = {
        "IMPORT NUEVO": (
            NEW_IMPORT
            in final_text
        ),
        "CALLBACK DINAMICO": (
            "set_change_callback("
            in final_text
        ),
        "VALIDACION MANUAL": (
            "validate_manual_authors()"
            in final_text
        ),
    }

    if not all(
        checks.values()
    ):
        raise RuntimeError(
            f"Verificación incompleta: {checks}"
        )

    print(
        "UI AUTORES MANUALES: ACTUALIZADA"
    )
    print(
        "IMPORT NUEVO: OK"
    )
    print(
        "CALLBACK DINAMICO: OK"
    )
    print(
        "VALIDACION MANUAL: OK"
    )
    print(
        "COMPILACION: OK"
    )
    print(
        "BACKUP MAIN:",
        MAIN_BACKUP,
    )


if __name__ == "__main__":
    main()