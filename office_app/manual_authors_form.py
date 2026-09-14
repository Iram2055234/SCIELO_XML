from __future__ import annotations

import tkinter as tk
from typing import Any

import customtkinter as ctk

from office_app.metadata_form import MetadataForm
from office_app.theme import (
    BORDER,
    ERROR,
    FACPYA_RED,
    FACPYA_RED_SOFT,
    FONT_FAMILY,
    INPUT_BG,
    SOFT_BG,
    TEXT_MUTED,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    WARNING,
    WARNING_SOFT,
)


class ManualAuthorsMetadataForm(MetadataForm):
    """
    Agrega captura manual de autores sin cambiar
    el comportamiento de los autores detectados.
    """

    def __init__(
        self,
        parent: tk.Misc,
        metadata: Any,
    ) -> None:
        self.manual_author_variables: list[
            dict[str, Any]
        ] = []
        self.manual_author_counter = 0
        self.authors_container: (
            ctk.CTkFrame | None
        ) = None
        self.authors_empty_box: (
            ctk.CTkFrame | None
        ) = None
        self.author_change_callback: Any = None

        super().__init__(
            parent=parent,
            metadata=metadata,
        )

    def set_change_callback(
        self,
        callback: Any,
    ) -> None:
        self.author_change_callback = callback

    def _notify_author_change(
        self,
    ) -> None:
        if self.author_change_callback is not None:
            self.author_change_callback()

    def _build_authors_section(
        self,
        row: int,
    ) -> None:
        section = self.create_section(
            title="Autores",
            subtitle=(
                "Revise los autores detectados. "
                "Si falta alguno, agréguelo manualmente. "
                "El ORCID es opcional."
            ),
            row=row,
        )

        authors = list(
            getattr(
                self.metadata,
                "authors",
                [],
            )
        )

        if not authors:
            self.authors_empty_box = ctk.CTkFrame(
                section,
                fg_color=WARNING_SOFT,
                corner_radius=10,
                border_width=1,
                border_color=(
                    "#FEDF89",
                    "#5E4725",
                ),
            )

            self.authors_empty_box.grid(
                row=0,
                column=0,
                columnspan=2,
                sticky="ew",
                pady=(0, 10),
            )

            ctk.CTkLabel(
                self.authors_empty_box,
                text=(
                    "No se detectaron autores "
                    "automáticamente. Puede agregarlos "
                    "manualmente antes de validar."
                ),
                text_color=WARNING,
                anchor="w",
                justify="left",
                wraplength=900,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).pack(
                fill="x",
                padx=16,
                pady=14,
            )

        self.authors_container = ctk.CTkFrame(
            section,
            fg_color="transparent",
        )

        self.authors_container.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="ew",
        )

        self.authors_container.grid_columnconfigure(
            0,
            weight=1,
        )

        for index, author in enumerate(
            authors,
            start=1,
        ):
            super()._build_author_row(
                parent=self.authors_container,
                row=index - 1,
                number=index,
                author=author,
            )

        ctk.CTkButton(
            section,
            text="+ Agregar autor manualmente",
            command=self._add_manual_author,
            height=38,
            corner_radius=9,
            fg_color=FACPYA_RED,
            hover_color=FACPYA_RED,
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
        ).grid(
            row=2,
            column=1,
            sticky="e",
            pady=(4, 0),
        )

    def _create_manual_field(
        self,
        parent: ctk.CTkFrame,
        column: int,
        label: str,
        field_key: str,
        required: bool = False,
        placeholder: str = "",
    ) -> tuple[
        tk.StringVar,
        ctk.CTkEntry,
        ctk.CTkLabel,
    ]:
        frame = ctk.CTkFrame(
            parent,
            fg_color="transparent",
        )

        frame.grid(
            row=0,
            column=column,
            sticky="ew",
            padx=(
                0 if column == 0 else 6,
                0 if column == 2 else 6,
            ),
        )

        frame.grid_columnconfigure(
            0,
            weight=1,
        )

        label_row = ctk.CTkFrame(
            frame,
            fg_color="transparent",
        )

        label_row.grid(
            row=0,
            column=0,
            sticky="w",
            pady=(0, 4),
        )

        ctk.CTkLabel(
            label_row,
            text=label,
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        if required:
            ctk.CTkLabel(
                label_row,
                text=" *",
                text_color=ERROR,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=1,
                sticky="w",
            )

        variable = tk.StringVar(
            master=self,
            value="",
        )

        entry = ctk.CTkEntry(
            frame,
            textvariable=variable,
            height=38,
            corner_radius=8,
            border_width=1,
            border_color=BORDER,
            fg_color=INPUT_BG,
            text_color=TEXT_PRIMARY,
            placeholder_text=placeholder,
            placeholder_text_color=TEXT_MUTED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        )

        entry.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        error_label = ctk.CTkLabel(
            frame,
            text="",
            text_color=ERROR,
            anchor="w",
            justify="left",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        )

        error_label.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(3, 0),
        )
        error_label.grid_remove()

        self.field_entries[
            field_key
        ] = entry

        self.field_error_labels[
            field_key
        ] = error_label

        def on_change(
            *_args: object,
        ) -> None:
            self._clear_field_validation(
                field_key
            )
            self._notify_author_change()

        variable.trace_add(
            "write",
            on_change,
        )

        return (
            variable,
            entry,
            error_label,
        )

    def _add_manual_author(
        self,
    ) -> None:
        if self.authors_container is None:
            return

        self.manual_author_counter += 1
        manual_id = self.manual_author_counter

        row = (
            len(self.author_variables)
            + len(self.manual_author_variables)
        )

        card = ctk.CTkFrame(
            self.authors_container,
            fg_color=SOFT_BG,
            corner_radius=10,
            border_width=1,
            border_color=FACPYA_RED_SOFT,
        )

        card.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(0, 10),
        )

        card.grid_columnconfigure(
            0,
            weight=1,
        )

        header = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        header.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=14,
            pady=(12, 8),
        )

        header.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            header,
            text="Autor agregado manualmente",
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        fields = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        fields.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=14,
            pady=(0, 14),
        )

        for column in range(3):
            fields.grid_columnconfigure(
                column,
                weight=1,
            )

        name_key = (
            f"manual_author.{manual_id}.name"
        )
        email_key = (
            f"manual_author.{manual_id}.email"
        )
        orcid_key = (
            f"manual_author.{manual_id}.orcid"
        )

        (
            name_variable,
            name_entry,
            _,
        ) = self._create_manual_field(
            parent=fields,
            column=0,
            label="Nombre completo",
            field_key=name_key,
            required=True,
            placeholder="Nombre completo",
        )

        (
            email_variable,
            _,
            _,
        ) = self._create_manual_field(
            parent=fields,
            column=1,
            label="Correo (opcional)",
            field_key=email_key,
        )

        (
            orcid_variable,
            _,
            _,
        ) = self._create_manual_field(
            parent=fields,
            column=2,
            label="ORCID (opcional)",
            field_key=orcid_key,
            placeholder="0000-0000-0000-0000",
        )

        record: dict[str, Any] = {
            "id": manual_id,
            "card": card,
            "name": name_variable,
            "email": email_variable,
            "orcid": orcid_variable,
            "name_key": name_key,
            "email_key": email_key,
            "orcid_key": orcid_key,
        }

        ctk.CTkButton(
            header,
            text="Eliminar",
            width=78,
            height=30,
            corner_radius=8,
            fg_color="transparent",
            hover_color=SOFT_BG,
            border_width=1,
            border_color=ERROR,
            text_color=ERROR,
            command=lambda item=record: (
                self._remove_manual_author(item)
            ),
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=1,
            sticky="e",
        )

        self.manual_author_variables.append(
            record
        )

        if self.authors_empty_box is not None:
            self.authors_empty_box.grid_remove()

        name_entry.focus_set()
        self._notify_author_change()

    def _remove_manual_author(
        self,
        record: dict[str, Any],
    ) -> None:
        if record not in self.manual_author_variables:
            return

        for key_name in (
            "name_key",
            "email_key",
            "orcid_key",
        ):
            field_key = str(
                record[key_name]
            )

            self.field_entries.pop(
                field_key,
                None,
            )
            self.field_error_labels.pop(
                field_key,
                None,
            )

        card = record.get(
            "card"
        )

        if card is not None:
            try:
                card.destroy()
            except tk.TclError:
                pass

        self.manual_author_variables.remove(
            record
        )

        if (
            not self.author_variables
            and not self.manual_author_variables
            and self.authors_empty_box is not None
        ):
            self.authors_empty_box.grid()

        self._notify_author_change()

    def validate_manual_authors(
        self,
    ) -> list[str]:
        errors: list[str] = []

        seen_names = {
            self.clean(name).casefold()
            for name in self.author_variables
            if self.clean(name)
        }

        valid_manual_count = 0

        for record in self.manual_author_variables:
            name = self.clean(
                record["name"].get()
            )

            name_key = str(
                record["name_key"]
            )

            self._clear_field_validation(
                name_key
            )

            if not name:
                self._mark_field_invalid(
                    name_key,
                    "Capture el nombre completo.",
                )
                errors.append(
                    "Capture el nombre completo de cada "
                    "autor agregado manualmente."
                )
                continue

            normalized_name = (
                name.casefold()
            )

            if normalized_name in seen_names:
                self._mark_field_invalid(
                    name_key,
                    "Este autor ya está registrado.",
                )
                errors.append(
                    f"El autor '{name}' está duplicado."
                )
                continue

            seen_names.add(
                normalized_name
            )
            valid_manual_count += 1

        if (
            not self.author_variables
            and valid_manual_count == 0
            and not self.manual_author_variables
        ):
            errors.append(
                "Agregue al menos un autor al artículo."
            )

        return errors

    def get_author_overrides(
        self,
    ) -> dict[
        str,
        dict[str, str],
    ]:
        result = (
            super().get_author_overrides()
        )

        for record in self.manual_author_variables:
            author_name = self.clean(
                record["name"].get()
            )

            if not author_name:
                continue

            result[author_name] = {
                "email": self.clean(
                    record["email"].get()
                ),
                "orcid": self.clean(
                    record["orcid"].get()
                ),
            }

        return result