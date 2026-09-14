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

FORM = (
    ROOT
    / "office_app"
    / "manual_authors_form.py"
)

BACKUP = (
    ROOT
    / "office_app"
    / "main_antes_autores_manual_ui.py"
)


NEW_IMPORT = '''from office_app.manual_authors_form import (
    ManualAuthorsMetadataForm as MetadataForm,
)'''


BIND_BLOCK = '''    def _bind_metadata_changes(self) -> None:
        """
        Detecta cambios en el formulario.
        """

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

'''


VALIDATION_BLOCK = '''            manual_author_errors = (
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

'''


def replace_between(
    text: str,
    start: str,
    end: str,
    replacement: str,
) -> str:
    start_index = text.find(
        start
    )

    if start_index < 0:
        raise RuntimeError(
            f"No se encontró: {start}"
        )

    end_index = text.find(
        end,
        start_index,
    )

    if end_index < 0:
        raise RuntimeError(
            f"No se encontró: {end}"
        )

    return (
        text[:start_index]
        + replacement
        + text[end_index:]
    )


def main() -> None:
    if not BACKUP.is_file():
        raise FileNotFoundError(
            "No existe el respaldo seguro: "
            f"{BACKUP}"
        )

    if not FORM.is_file():
        raise FileNotFoundError(
            f"No existe: {FORM}"
        )

    # Regresar al main sano anterior al cambio UI.
    shutil.copy2(
        BACKUP,
        MAIN,
    )

    text = MAIN.read_text(
        encoding="utf-8"
    )

    old_import = (
        "from office_app.metadata_form "
        "import MetadataForm"
    )

    if old_import not in text:
        raise RuntimeError(
            "No se encontró la importación "
            "original de MetadataForm."
        )

    text = text.replace(
        old_import,
        NEW_IMPORT,
        1,
    )

    text = replace_between(
        text=text,
        start=(
            "    def _bind_metadata_changes("
        ),
        end=(
            "    def _on_metadata_changed("
        ),
        replacement=BIND_BLOCK,
    )

    marker = '''            ) = self._get_form_overrides()

            result = (
'''

    if marker not in text:
        raise RuntimeError(
            "No se encontró el punto de "
            "validación esperado."
        )

    replacement = (
        '''            ) = self._get_form_overrides()

'''
        + VALIDATION_BLOCK
        + '''            result = (
'''
    )

    text = text.replace(
        marker,
        replacement,
        1,
    )

    MAIN.write_text(
        text,
        encoding="utf-8",
    )

    py_compile.compile(
        str(FORM),
        doraise=True,
    )

    py_compile.compile(
        str(MAIN),
        doraise=True,
    )

    final_text = MAIN.read_text(
        encoding="utf-8"
    )

    required = (
        NEW_IMPORT,
        "self.metadata_form.set_change_callback(",
        ".validate_manual_authors()",
        (
            '"""\n'
            "        Detecta cambios en el formulario.\n"
            '        """'
        ),
    )

    if not all(
        item in final_text
        for item in required
    ):
        raise RuntimeError(
            "La verificación final de main.py "
            "no fue completa."
        )

    print(
        "MAIN RESTAURADO DESDE BACKUP: OK"
    )
    print(
        "UI AUTORES MANUALES: ACTUALIZADA"
    )
    print(
        "MANUAL AUTHORS FORM: OK"
    )
    print(
        "COMPILACION MAIN: OK"
    )
    print(
        "REPARACION UI: APROBADA"
    )


if __name__ == "__main__":
    main()