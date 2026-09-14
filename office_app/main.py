from __future__ import annotations

import os
import subprocess
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog
from typing import Callable

import customtkinter as ctk

from converter.conversion_history import (
    STATUS_CONVERTED,
    STATUS_FAILED,
    STATUS_IN_PROGRESS,
    list_recent_history,
)
from converter.office_config_builder import (
    OfficeConfigBuildResult,
    build_and_validate_office_config,
)
from converter.office_metadata_enricher import (
    extract_complete_office_metadata,
)
from converter.office_metadata_extractor import OfficeMetadata
from converter.metadata_detection_telemetry import (
    MetadataDetectionEvent,
    compare_metadata_detection,
)
from converter.office_result_manager import get_results_root
from converter.workstation_identity import (
    load_station_identity,
    save_station_alias,
)
from office_app.conversion_runner import (
    OfficeGuiConversionResult,
    run_office_conversion,
)
from office_app.file_opener import (
    open_directory,
    open_html,
    open_xml_in_notepad,
)
from office_app.editable_authors_form import (
    EditableAuthorsMetadataForm as MetadataForm,
)
from office_app.theme import (
    APP_BG,
    BORDER,
    BUTTON_DISABLED,
    BUTTON_DISABLED_TEXT,
    CARD_BG,
    CHIP_BG,
    ERROR,
    ERROR_HOVER,
    ERROR_SOFT,
    FACPYA_RED,
    FACPYA_RED_HOVER,
    FACPYA_RED_SOFT,
    FONT_FAMILY,
    INFO,
    INFO_SOFT,
    SCROLLBAR,
    SCROLLBAR_HOVER,
    SECONDARY_HOVER,
    SOFT_BG,
    SUCCESS,
    SUCCESS_SOFT,
    TEXT_MUTED,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    WARNING,
    WARNING_SOFT,
    configure_appearance,
    set_dark_mode,
    set_light_mode,
)

configure_appearance()


class ModernDialog(ctk.CTkToplevel):
    """
    Diálogo visual propio de la aplicación.
    """

    def __init__(
        self,
        parent: tk.Misc,
        title: str,
        message: str,
        kind: str = "info",
        confirm_text: str = "Aceptar",
        cancel_text: str | None = None,
        destructive: bool = False,
    ) -> None:
        super().__init__(parent)

        self.withdraw()

        self.result = False

        self.title(title)
        self.geometry("560x390")
        self.resizable(False, False)
        self.configure(fg_color=APP_BG)
        self.transient(parent)

        self.protocol(
            "WM_DELETE_WINDOW",
            self._cancel,
        )

        self.grid_columnconfigure(
            0,
            weight=1,
        )

        self.grid_rowconfigure(
            0,
            weight=1,
        )

        if kind == "success":
            icon_text = "✓"
            icon_color = SUCCESS
            icon_background = SUCCESS_SOFT

        elif kind == "warning":
            icon_text = "!"
            icon_color = WARNING
            icon_background = WARNING_SOFT

        elif kind == "error":
            icon_text = "×"
            icon_color = ERROR
            icon_background = ERROR_SOFT

        else:
            icon_text = "i"
            icon_color = INFO
            icon_background = INFO_SOFT

        card = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            corner_radius=18,
            border_width=1,
            border_color=BORDER,
        )

        card.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=20,
            pady=20,
        )

        card.grid_columnconfigure(
            0,
            weight=1,
        )

        icon_container = ctk.CTkFrame(
            card,
            width=60,
            height=60,
            fg_color=icon_background,
            corner_radius=30,
        )

        icon_container.grid(
            row=0,
            column=0,
            pady=(28, 12),
        )

        icon_container.grid_propagate(False)

        ctk.CTkLabel(
            icon_container,
            text=icon_text,
            text_color=icon_color,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=27,
                weight="bold",
            ),
        ).place(
            relx=0.5,
            rely=0.5,
            anchor="center",
        )

        ctk.CTkLabel(
            card,
            text=title,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=19,
                weight="bold",
            ),
        ).grid(
            row=1,
            column=0,
            padx=30,
            sticky="ew",
        )

        ctk.CTkLabel(
            card,
            text=message,
            text_color=TEXT_SECONDARY,
            justify="center",
            anchor="center",
            wraplength=460,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
            ),
        ).grid(
            row=2,
            column=0,
            sticky="ew",
            padx=30,
            pady=(10, 24),
        )

        actions = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        actions.grid(
            row=3,
            column=0,
            pady=(0, 28),
        )

        if cancel_text:
            ctk.CTkButton(
                actions,
                text=cancel_text,
                width=130,
                height=42,
                corner_radius=9,
                fg_color=CARD_BG,
                hover_color=SECONDARY_HOVER,
                border_width=1,
                border_color=BORDER,
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
                command=self._cancel,
            ).grid(
                row=0,
                column=0,
                padx=(0, 10),
            )

        confirm_color = (
            ERROR
            if destructive
            else FACPYA_RED
        )

        confirm_hover = (
            ERROR_HOVER
            if destructive
            else FACPYA_RED_HOVER
        )

        confirm_button = ctk.CTkButton(
            actions,
            text=confirm_text,
            width=(
                190
                if cancel_text
                else 135
            ),
            height=42,
            corner_radius=9,
            fg_color=confirm_color,
            hover_color=confirm_hover,
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
            command=self._confirm,
        )

        confirm_button.grid(
            row=0,
            column=(
                1
                if cancel_text
                else 0
            ),
        )

        self.after(
            30,
            lambda: self._prepare_dialog(
                parent,
                confirm_button,
            ),
        )

    def _prepare_dialog(
        self,
        parent: tk.Misc,
        confirm_button: ctk.CTkButton,
    ) -> None:
        """
        Centra el diálogo sobre la ventana principal
        manteniendo su tamaño actual.
        """

        self.update_idletasks()

        try:
            parent.update_idletasks()

            dialog_width = max(
                1,
                self.winfo_width(),
            )

            dialog_height = max(
                1,
                self.winfo_height(),
            )

            parent_width = max(
                1,
                parent.winfo_width(),
            )

            parent_height = max(
                1,
                parent.winfo_height(),
            )

            x = (
                parent.winfo_rootx()
                + (
                    parent_width
                    - dialog_width
                )
                // 2
            )

            y = (
                parent.winfo_rooty()
                + (
                    parent_height
                    - dialog_height
                )
                // 2
            )

            screen_width = (
                self.winfo_screenwidth()
            )

            screen_height = (
                self.winfo_screenheight()
            )

            x = max(
                0,
                min(
                    x,
                    screen_width
                    - dialog_width,
                ),
            )

            y = max(
                0,
                min(
                    y,
                    screen_height
                    - dialog_height,
                ),
            )

            self.geometry(
                (
                    f"{dialog_width}x{dialog_height}"
                    f"+{x}+{y}"
                )
            )

        except tk.TclError:
            pass

        self.deiconify()
        self.lift()
        self.grab_set()
        self.focus_force()
        confirm_button.focus_set()

    def _confirm(self) -> None:
        self.result = True
        self._close()

    def _cancel(self) -> None:
        self.result = False
        self._close()

    def _close(self) -> None:
        try:
            self.grab_release()

        except tk.TclError:
            pass

        self.destroy()

    @classmethod
    def show(
        cls,
        parent: tk.Misc,
        title: str,
        message: str,
        kind: str = "info",
    ) -> None:
        """
        Muestra un diálogo informativo.
        """

        dialog = cls(
            parent=parent,
            title=title,
            message=message,
            kind=kind,
        )

        parent.wait_window(
            dialog
        )

    @classmethod
    def ask(
        cls,
        parent: tk.Misc,
        title: str,
        message: str,
        confirm_text: str,
        cancel_text: str = "Cancelar",
        kind: str = "warning",
        destructive: bool = False,
    ) -> bool:
        """
        Muestra un diálogo de confirmación.
        """

        dialog = cls(
            parent=parent,
            title=title,
            message=message,
            kind=kind,
            confirm_text=confirm_text,
            cancel_text=cancel_text,
            destructive=destructive,
        )

        parent.wait_window(
            dialog
        )

        return dialog.result


class ScieloOfficeApp(ctk.CTk):
    """
    Convertidor SCIELO de escritorio.
    """

    def __init__(self) -> None:
        super().__init__()

        self.selected_docx: Path | None = None
        self.current_metadata: OfficeMetadata | None = None

        self.metadata_detection_events: list[
            MetadataDetectionEvent
        ] = []

        self.metadata_form: MetadataForm | None = None

        self.current_config_result: (
            OfficeConfigBuildResult | None
        ) = None

        self.last_conversion_result: (
            OfficeGuiConversionResult | None
        ) = None

        self.validate_button: (
            ctk.CTkButton | None
        ) = None

        self.convert_button: (
            ctk.CTkButton | None
        ) = None

        self.progress_bar: (
            ctk.CTkProgressBar | None
        ) = None

        self.validation_label: (
            ctk.CTkLabel | None
        ) = None

        self.theme_button: (
            ctk.CTkButton | None
        ) = None

        self.station_button: (
            ctk.CTkButton | None
        ) = None

        self.dashboard_button: (
            ctk.CTkButton | None
        ) = None

        self.dashboard_view = None

        self.station_identity = (
            load_station_identity()
        )

        self.is_converting = False

        self.validation_variable = tk.StringVar(
            master=self,
            value="",
        )

        self.status_variable = tk.StringVar(
            master=self,
            value="Listo",
        )

        self.conversion_status_variable = tk.StringVar(
            master=self,
            value="",
        )

        self.current_step = 1

        self.title(
            "Convertidor SCIELO - FACPyA"
        )

        self.configure(
            fg_color=APP_BG
        )

        self.geometry(
            "1180x800"
        )

        self.minsize(
            940,
            650,
        )

        self.protocol(
            "WM_DELETE_WINDOW",
            self._on_close,
        )

        self.grid_columnconfigure(
            0,
            weight=1,
        )

        self.grid_rowconfigure(
            2,
            weight=1,
        )

        self._build_header()
        self._build_stepper()
        self._build_content_host()
        self._build_status_bar()
        self._build_persistent_signature()

        self.show_selection_page()

        self.after_idle(
            self._raise_persistent_signature
        )

        self.after(
            0,
            self.maximize_window,
        )

    def _build_persistent_signature(
        self,
    ) -> None:
        """
        Firma institucional persistente de la aplicación.

        Se crea directamente sobre la ventana raíz para que
        permanezca visible al cambiar de página y al mostrar
        vistas embebidas como el Dashboard.
        """

        self.signature_label = ctk.CTkLabel(
            self,
            text="© Iram Rodríguez López",
            fg_color="transparent",
            text_color=(
                "#6B7280",
                "#9CA3AF",
            ),
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
            height=18,
        )

        self.signature_label.place(
            relx=1.0,
            rely=1.0,
            x=-14,
            y=-8,
            anchor="se",
        )

        self._raise_persistent_signature()


    def _raise_persistent_signature(
        self,
    ) -> None:
        """
        Mantiene la firma por encima de las vistas embebidas
        sin interferir con su geometría.
        """

        signature = getattr(
            self,
            "signature_label",
            None,
        )

        if signature is None:
            return

        try:
            if signature.winfo_exists():
                signature.lift()

        except Exception:
            pass


    def _build_header(self) -> None:
        """
        Encabezado superior.
        """

        header = ctk.CTkFrame(
            self,
            height=86,
            corner_radius=0,
            fg_color=CARD_BG,
        )

        header.grid(
            row=0,
            column=0,
            sticky="ew",
        )

        header.grid_propagate(False)

        header.grid_columnconfigure(
            0,
            weight=1,
        )

        brand = ctk.CTkFrame(
            header,
            fg_color="transparent",
        )

        brand.grid(
            row=0,
            column=0,
            sticky="w",
            padx=28,
            pady=14,
        )

        ctk.CTkLabel(
            brand,
            text="Convertidor SCIELO",
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=21,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ctk.CTkLabel(
            brand,
            text=(
                "Preparación de artículos "
                "para publicación científica"
            ),
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(2, 0),
        )

        right_actions = ctk.CTkFrame(
            header,
            fg_color="transparent",
        )

        right_actions.grid(
            row=0,
            column=1,
            sticky="e",
            padx=28,
            pady=14,
        )

        self.station_button = ctk.CTkButton(
            right_actions,
            text=(
                "Equipo · "
                + self.station_identity.alias
            ),
            height=42,
            corner_radius=9,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
            command=(
                self.configure_station_alias
            ),
        )

        self.station_button.grid(
            row=0,
            column=0,
            padx=(0, 10),
        )

        self.dashboard_button = ctk.CTkButton(
            right_actions,
            text="Dashboard",
            width=108,
            height=42,
            corner_radius=9,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
            command=self.open_dashboard_view,
        )

        self.dashboard_button.grid(
            row=0,
            column=1,
            padx=(0, 10),
        )

        self.theme_button = ctk.CTkButton(
            right_actions,
            text="☾",
            width=42,
            height=42,
            corner_radius=21,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family="Segoe UI Symbol",
                size=19,
                weight="bold",
            ),
            command=self.toggle_theme,
        )

        self.theme_button.grid(
            row=0,
            column=2,
            padx=(0, 10),
        )

        institution = ctk.CTkFrame(
            right_actions,
            fg_color=FACPYA_RED_SOFT,
            corner_radius=9,
        )

        institution.grid(
            row=0,
            column=3,
        )

        ctk.CTkLabel(
            institution,
            text="FACPyA · UANL",
            text_color=FACPYA_RED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
        ).pack(
            padx=15,
            pady=10,
        )

        accent = ctk.CTkFrame(
            self,
            height=3,
            corner_radius=0,
            fg_color=FACPYA_RED,
        )

        accent.place(
            x=0,
            y=83,
            relwidth=1,
        )

    def configure_station_alias(self) -> None:
        """
        Permite asignar un nombre amigable al equipo.

        La identificación se conserva localmente y no
        sustituye el hostname real de Windows, que también
        se registra en el historial.
        """

        dialog = ctk.CTkInputDialog(
            text=(
                "Nombre visible de esta computadora:\n\n"
                f"Hostname: {self.station_identity.hostname}"
            ),
            title="Identificar computadora",
        )

        value = dialog.get_input()

        if value is None:
            return

        try:
            self.station_identity = (
                save_station_alias(
                    value
                )
            )

        except Exception as error:
            ModernDialog.show(
                self,
                "Nombre no guardado",
                str(error),
                kind="error",
            )
            return

        self._refresh_station_button()

        self.status_variable.set(
            "Identificación del equipo actualizada."
        )

        if self.current_step == 1:
            self.show_selection_page()

    def _refresh_station_button(self) -> None:
        """Actualiza el alias mostrado en el encabezado."""

        if self.station_button is None:
            return

        self.station_button.configure(
            text=(
                "Equipo · "
                + self.station_identity.alias
            )
        )

    def open_dashboard_view(self) -> None:
        """
        Muestra el Dashboard dentro de la misma ventana
        del Convertidor, sin crear una ventana secundaria.
        """

        if self.is_converting:
            ModernDialog.show(
                self,
                "Conversión en proceso",
                (
                    "Espere a que termine la conversión "
                    "antes de abrir el Dashboard."
                ),
                kind="warning",
            )
            return

        if self.dashboard_view is not None:
            try:
                if self.dashboard_view.winfo_exists():
                    self.dashboard_view.lift()
                    self.status_variable.set(
                        "Dashboard abierto."
                    )
                    return
            except Exception:
                self.dashboard_view = None

        try:
            from office_app.dashboard import open_dashboard

            self.dashboard_view = open_dashboard(
                master=self,
                results_root=get_results_root(),
                on_close=self._return_from_dashboard,
            )

            self.status_variable.set(
                "Dashboard abierto."
            )

        except Exception as error:
            self.dashboard_view = None

            ModernDialog.show(
                self,
                "Dashboard no disponible",
                (
                    "No fue posible abrir el Dashboard.\n\n"
                    f"{type(error).__name__}: {error}"
                ),
                kind="error",
            )

    def _return_from_dashboard(self) -> None:
        """
        Regresa al flujo que ya estaba visible en el
        Convertidor. No reinicia ni reconstruye la app.
        """

        self.dashboard_view = None
        self.status_variable.set(
            "Regresó al Convertidor SCIELO."
        )

    def _apply_windows_titlebar_theme_without_withdraw(
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
        if self.dashboard_view is not None:
            try:
                if self.dashboard_view.winfo_exists():
                    self.after(
                        10,
                        self.dashboard_view.refresh_theme,
                    )
            except Exception:
                self.dashboard_view = None


    def _build_stepper(self) -> None:
        """
        Indicador de cuatro etapas.
        """

        self.stepper = ctk.CTkFrame(
            self,
            height=66,
            corner_radius=0,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER,
        )

        self.stepper.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        self.stepper.grid_propagate(False)

        self.stepper.grid_rowconfigure(
            0,
            weight=1,
        )

        for column in range(4):
            self.stepper.grid_columnconfigure(
                column,
                weight=1,
                uniform="steps",
            )

        self.step_widgets: list[
            tuple[
                ctk.CTkFrame,
                ctk.CTkLabel,
                ctk.CTkLabel,
            ]
        ] = []

        step_names = [
            "Seleccionar",
            "Revisar",
            "Convertir",
            "Resultados",
        ]

        for index, name in enumerate(
            step_names,
            start=1,
        ):
            step_frame = ctk.CTkFrame(
                self.stepper,
                fg_color="transparent",
            )

            step_frame.grid(
                row=0,
                column=index - 1,
                sticky="nsew",
                padx=6,
                pady=8,
            )

            content = ctk.CTkFrame(
                step_frame,
                fg_color="transparent",
            )

            content.pack(
                expand=True
            )

            number_label = ctk.CTkLabel(
                content,
                text=str(index),
                width=30,
                height=30,
                corner_radius=15,
                fg_color=CHIP_BG,
                text_color=TEXT_MUTED,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                    weight="bold",
                ),
            )

            number_label.grid(
                row=0,
                column=0,
                padx=(0, 8),
            )

            name_label = ctk.CTkLabel(
                content,
                text=name,
                text_color=TEXT_MUTED,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                    weight="bold",
                ),
            )

            name_label.grid(
                row=0,
                column=1,
            )

            self.step_widgets.append(
                (
                    step_frame,
                    number_label,
                    name_label,
                )
            )

        self._set_step(1)

    def _set_step(
        self,
        step: int,
    ) -> None:
        """
        Actualiza el indicador.
        """

        self.current_step = step

        for index, (
            _frame,
            number_label,
            name_label,
        ) in enumerate(
            self.step_widgets,
            start=1,
        ):
            if index < step:
                number_label.configure(
                    text="✓",
                    fg_color=SUCCESS,
                    text_color="#FFFFFF",
                )

                name_label.configure(
                    text_color=SUCCESS
                )

            elif index == step:
                number_label.configure(
                    text=str(index),
                    fg_color=FACPYA_RED,
                    text_color="#FFFFFF",
                )

                name_label.configure(
                    text_color=FACPYA_RED
                )

            else:
                number_label.configure(
                    text=str(index),
                    fg_color=CHIP_BG,
                    text_color=TEXT_MUTED,
                )

                name_label.configure(
                    text_color=TEXT_MUTED
                )

    def _build_content_host(self) -> None:
        """
        Contenedor de páginas.
        """

        self.content_host = ctk.CTkFrame(
            self,
            fg_color=APP_BG,
            corner_radius=0,
        )

        self.content_host.grid(
            row=2,
            column=0,
            sticky="nsew",
        )

        self.content_host.grid_columnconfigure(
            0,
            weight=1,
        )

        self.content_host.grid_rowconfigure(
            0,
            weight=1,
        )

    def _build_status_bar(self) -> None:
        """
        Barra de estado.
        """

        status_bar = ctk.CTkFrame(
            self,
            height=34,
            fg_color=CARD_BG,
            corner_radius=0,
            border_width=1,
            border_color=BORDER,
        )

        status_bar.grid(
            row=3,
            column=0,
            sticky="ew",
        )

        status_bar.grid_propagate(False)

        ctk.CTkLabel(
            status_bar,
            textvariable=self.status_variable,
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            fill="x",
            padx=18,
            pady=7,
        )

    def clear_content(self) -> None:
        """
        Limpia la página actual.
        """

        for child in (
            self.content_host.winfo_children()
        ):
            child.destroy()

        self.metadata_form = None
        self.validate_button = None
        self.convert_button = None
        self.progress_bar = None
        self.validation_label = None

    def reset_validation(self) -> None:
        """
        Restablece validación.
        """

        self.current_config_result = None
        self.validation_variable.set("")

    def primary_button(
        self,
        parent: tk.Misc,
        text: str,
        command: Callable[[], None],
        width: int = 170,
    ) -> ctk.CTkButton:
        """
        Botón principal.
        """

        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=42,
            corner_radius=9,
            fg_color=FACPYA_RED,
            hover_color=FACPYA_RED_HOVER,
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        )

    def secondary_button(
        self,
        parent: tk.Misc,
        text: str,
        command: Callable[[], None],
        width: int = 150,
    ) -> ctk.CTkButton:
        """
        Botón secundario.
        """

        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=42,
            corner_radius=9,
            fg_color=CARD_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        )

    def show_selection_page(self) -> None:
        """
        Pantalla inicial.
        """

        self.clear_content()
        self._set_step(1)

        self.status_variable.set(
            (
                "Artículo seleccionado. "
                "Listo para revisar metadatos."
                if self.selected_docx
                else "Listo"
            )
        )

        page = ctk.CTkScrollableFrame(
            self.content_host,
            fg_color=APP_BG,
            corner_radius=0,
            scrollbar_button_color=SCROLLBAR,
            scrollbar_button_hover_color=(
                SCROLLBAR_HOVER
            ),
        )

        page.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        page.grid_columnconfigure(
            0,
            weight=1,
        )

        card = ctk.CTkFrame(
            page,
            width=880,
            fg_color=CARD_BG,
            corner_radius=18,
            border_width=1,
            border_color=BORDER,
        )

        card.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=48,
            pady=(24, 12),
        )

        card.grid_columnconfigure(
            0,
            weight=1,
        )

        top_accent = ctk.CTkFrame(
            card,
            height=6,
            fg_color=FACPYA_RED,
            corner_radius=3,
        )

        top_accent.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=26,
            pady=(22, 24),
        )

        ctk.CTkLabel(
            card,
            text="Seleccione el artículo Word",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=22,
                weight="bold",
            ),
        ).grid(
            row=1,
            column=0,
            padx=40,
        )

        ctk.CTkLabel(
            card,
            text=(
                "Seleccione el documento DOCX que desea "
                "preparar para publicación."
            ),
            text_color=TEXT_SECONDARY,
            wraplength=720,
            justify="center",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
            ),
        ).grid(
            row=2,
            column=0,
            padx=40,
            pady=(5, 22),
        )

        file_box = ctk.CTkFrame(
            card,
            fg_color=SOFT_BG,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
        )

        file_box.grid(
            row=3,
            column=0,
            sticky="ew",
            padx=52,
        )

        file_box.grid_columnconfigure(
            0,
            weight=1,
        )

        if self.selected_docx:
            ctk.CTkLabel(
                file_box,
                text="✓",
                width=44,
                height=44,
                corner_radius=22,
                fg_color=SUCCESS_SOFT,
                text_color=SUCCESS,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=20,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=0,
                pady=(22, 8),
            )

            ctk.CTkLabel(
                file_box,
                text=self.selected_docx.name,
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=14,
                    weight="bold",
                ),
            ).grid(
                row=1,
                column=0,
                padx=22,
            )

            ctk.CTkLabel(
                file_box,
                text=str(
                    self.selected_docx
                ),
                text_color=TEXT_SECONDARY,
                wraplength=680,
                justify="center",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).grid(
                row=2,
                column=0,
                padx=22,
                pady=(3, 13),
            )

            self.secondary_button(
                file_box,
                "Cambiar archivo",
                self.select_docx,
                width=145,
            ).grid(
                row=3,
                column=0,
                pady=(0, 22),
            )

        else:
            ctk.CTkLabel(
                file_box,
                text="DOCX",
                width=60,
                height=60,
                corner_radius=14,
                fg_color=FACPYA_RED_SOFT,
                text_color=FACPYA_RED,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=12,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=0,
                pady=(24, 9),
            )

            ctk.CTkLabel(
                file_box,
                text="Ningún artículo seleccionado",
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=13,
                    weight="bold",
                ),
            ).grid(
                row=1,
                column=0,
            )

            ctk.CTkLabel(
                file_box,
                text=(
                    "Documento de Microsoft Word (.docx)"
                ),
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).grid(
                row=2,
                column=0,
                pady=(4, 14),
            )

            self.primary_button(
                file_box,
                "Seleccionar artículo",
                self.select_docx,
                width=220,
            ).grid(
                row=3,
                column=0,
                pady=(0, 24),
            )

        formats = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        formats.grid(
            row=4,
            column=0,
            pady=(20, 18),
        )

        ctk.CTkLabel(
            formats,
            text="Archivos de salida",
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            padx=(0, 8),
        )

        for column, format_name in enumerate(
            [
                "XML JATS",
                "HTML",
                "PDF",
                "EPUB",
                "Imágenes",
            ],
            start=1,
        ):
            ctk.CTkLabel(
                formats,
                text=format_name,
                fg_color=CHIP_BG,
                corner_radius=7,
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=column,
                padx=3,
                ipadx=9,
                ipady=4,
            )

        actions = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        actions.grid(
            row=5,
            column=0,
            sticky="ew",
            padx=52,
            pady=(0, 28),
        )

        actions.grid_columnconfigure(
            0,
            weight=1,
        )

        if self.selected_docx:
            self.secondary_button(
                actions,
                "Seleccionar otro",
                self.select_docx,
                width=145,
            ).grid(
                row=0,
                column=0,
                sticky="w",
            )

        review_button = self.primary_button(
            actions,
            "Revisar metadatos  →",
            self.review_metadata,
            width=200,
        )

        review_button.grid(
            row=0,
            column=1,
            sticky="e",
        )

        if not self.selected_docx:
            review_button.configure(
                state="disabled",
                fg_color=BUTTON_DISABLED,
                text_color=BUTTON_DISABLED_TEXT,
            )

        self._build_history_card(
            page
        )

    def _build_history_card(
        self,
        parent: tk.Misc,
    ) -> None:
        """Construye el historial reciente en la portada."""

        card = ctk.CTkFrame(
            parent,
            fg_color=CARD_BG,
            corner_radius=16,
            border_width=1,
            border_color=BORDER,
        )

        card.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=48,
            pady=(0, 24),
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
            padx=22,
            pady=(18, 10),
        )

        header.grid_columnconfigure(
            0,
            weight=1,
        )

        history_title = ctk.CTkFrame(
            header,
            fg_color="transparent",
        )
        history_title.grid(
            row=0,
            column=0,
            sticky="w",
        )

        ctk.CTkLabel(
            history_title,
            text="Historial reciente",
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=15,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ctk.CTkLabel(
            history_title,
            text=(
                "Última actualización: "
                + datetime.now().strftime("%H:%M:%S")
            ),
            text_color=TEXT_MUTED,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(2, 0),
        )

        self.secondary_button(
            header,
            "Resultados",
            self.open_results_root,
            width=105,
        ).grid(
            row=0,
            column=1,
            padx=(8, 6),
        )

        self.secondary_button(
            header,
            "Actualizar",
            self.show_selection_page,
            width=100,
        ).grid(
            row=0,
            column=2,
        )

        try:
            records = list_recent_history(
                limit=5
            )
            history_error = ""

        except Exception as error:
            records = []
            history_error = str(error)

        if history_error:
            ctk.CTkLabel(
                card,
                text=(
                    "Historial no disponible: "
                    + history_error
                ),
                text_color=WARNING,
                wraplength=760,
                justify="left",
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).grid(
                row=1,
                column=0,
                sticky="ew",
                padx=22,
                pady=(0, 18),
            )
            return

        if not records:
            ctk.CTkLabel(
                card,
                text="Todavía no hay conversiones registradas.",
                text_color=TEXT_SECONDARY,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                ),
            ).grid(
                row=1,
                column=0,
                sticky="ew",
                padx=22,
                pady=(0, 18),
            )
            return

        rows = ctk.CTkFrame(
            card,
            fg_color=SOFT_BG,
            corner_radius=10,
        )

        rows.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=22,
            pady=(0, 18),
        )

        rows.grid_columnconfigure(
            0,
            weight=3,
        )
        rows.grid_columnconfigure(
            1,
            weight=2,
        )
        rows.grid_columnconfigure(
            2,
            weight=2,
        )

        for row_number, record in enumerate(
            records
        ):
            article_name = str(
                record.get(
                    "article_name",
                    "Artículo",
                )
            )

            station = str(
                record.get(
                    "workstation_alias",
                    "",
                )
                or record.get(
                    "hostname",
                    "",
                )
            )

            timestamp = str(
                record.get(
                    "finished_at",
                    "",
                )
                or record.get(
                    "started_at",
                    "",
                )
            )

            visible_time = (
                timestamp[5:16]
                .replace("T", " · ")
                if len(timestamp) >= 16
                else timestamp
            )

            status = str(
                record.get(
                    "status",
                    "",
                )
            ).upper()

            if status == STATUS_CONVERTED:
                status_text = "Convertido"
                status_color = SUCCESS
            elif status == STATUS_FAILED:
                status_text = "Fallido"
                status_color = ERROR
            else:
                status_text = "En proceso"
                status_color = INFO

            ctk.CTkLabel(
                rows,
                text=article_name,
                text_color=TEXT_PRIMARY,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                    weight="bold",
                ),
            ).grid(
                row=row_number,
                column=0,
                sticky="ew",
                padx=(14, 8),
                pady=7,
            )

            ctk.CTkLabel(
                rows,
                text=(
                    station
                    + (
                        " · " + visible_time
                        if visible_time
                        else ""
                    )
                ),
                text_color=TEXT_SECONDARY,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).grid(
                row=row_number,
                column=1,
                sticky="ew",
                padx=8,
            )

            ctk.CTkLabel(
                rows,
                text=status_text,
                text_color=status_color,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
            ).grid(
                row=row_number,
                column=2,
                sticky="w",
                padx=8,
            )

            action_parent = ctk.CTkFrame(
                rows,
                fg_color="transparent",
            )
            action_parent.grid(
                row=row_number,
                column=3,
                padx=(8, 12),
                pady=4,
            )

            if (
                status == STATUS_CONVERTED
                and record.get(
                    "result_directory"
                )
            ):
                self.secondary_button(
                    action_parent,
                    "Abrir",
                    lambda value=str(
                        record.get(
                            "result_directory",
                            "",
                        )
                    ): self.open_history_result(
                        value
                    ),
                    width=74,
                ).pack()

            elif status == STATUS_FAILED:
                self.secondary_button(
                    action_parent,
                    "Error",
                    lambda item=dict(record): (
                        self.show_history_error(
                            item
                        )
                    ),
                    width=74,
                ).pack()

    def open_results_root(self) -> None:
        """Abre la carpeta general de resultados."""

        try:
            open_directory(
                get_results_root()
            )
        except Exception as error:
            ModernDialog.show(
                self,
                "Resultados no disponibles",
                str(error),
                kind="error",
            )

    def open_history_result(
        self,
        directory: str,
    ) -> None:
        """Abre la carpeta de un registro histórico."""

        try:
            open_directory(
                Path(directory)
            )
        except Exception as error:
            ModernDialog.show(
                self,
                "Resultado no disponible",
                str(error),
                kind="error",
            )

    def show_history_error(
        self,
        record: dict,
    ) -> None:
        """Muestra el error asociado a una conversión."""

        error = str(
            record.get(
                "error",
                "",
            )
            or "No se registró un mensaje de error."
        )

        ModernDialog.show(
            self,
            "Conversión fallida",
            error,
            kind="error",
        )

    def select_docx(self) -> None:
        """
        Selecciona documento Word.
        """

        selected_file = (
            filedialog.askopenfilename(
                parent=self,
                title="Seleccionar artículo completo",
                filetypes=[
                    (
                        "Documentos de Word",
                        "*.docx",
                    ),
                    (
                        "Todos los archivos",
                        "*.*",
                    ),
                ],
            )
        )

        if not selected_file:
            return

        selected_path = Path(
            selected_file
        ).resolve()

        if (
            selected_path.suffix.casefold()
            != ".docx"
        ):
            ModernDialog.show(
                self,
                "Archivo no válido",
                (
                    "Seleccione un documento de "
                    "Microsoft Word con extensión .docx."
                ),
                kind="warning",
            )

            return

        if selected_path.name.startswith(
            "~$"
        ):
            ModernDialog.show(
                self,
                "Archivo temporal",
                (
                    "Ese archivo es un documento temporal "
                    "creado por Microsoft Word.\n\n"
                    "Seleccione el archivo original."
                ),
                kind="warning",
            )

            return

        if not selected_path.is_file():
            ModernDialog.show(
                self,
                "Archivo no encontrado",
                (
                    "El archivo seleccionado ya "
                    "no se encuentra disponible."
                ),
                kind="error",
            )

            return

        self.selected_docx = selected_path
        self.current_metadata = None
        self.metadata_detection_events = []
        self.last_conversion_result = None

        self.reset_validation()
        self.show_selection_page()

    def review_metadata(self) -> None:
        """
        Extrae metadatos.
        """

        if self.selected_docx is None:
            return

        self.status_variable.set(
            "Extrayendo metadatos del artículo..."
        )

        self.update_idletasks()

        try:
            metadata = (
                extract_complete_office_metadata(
                    self.selected_docx
                )
            )

        except Exception as error:
            self.status_variable.set(
                "No fue posible extraer los metadatos."
            )

            ModernDialog.show(
                self,
                "No fue posible leer el artículo",
                (
                    "Ocurrió un problema al extraer "
                    "los metadatos del documento.\n\n"
                    f"{type(error).__name__}: {error}"
                ),
                kind="error",
            )

            return

        self.current_metadata = metadata

        self.reset_validation()
        self.show_metadata_page()

        self.status_variable.set(
            "Metadatos extraídos. Revise la información."
        )

    def show_metadata_page(self) -> None:
        """
        Revisión de metadatos.
        """

        if (
            self.selected_docx is None
            or self.current_metadata is None
        ):
            return

        self.clear_content()
        self._set_step(2)

        page = ctk.CTkFrame(
            self.content_host,
            fg_color=APP_BG,
        )

        page.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=24,
            pady=16,
        )

        page.grid_columnconfigure(
            0,
            weight=1,
        )

        page.grid_rowconfigure(
            1,
            weight=1,
        )

        header_card = ctk.CTkFrame(
            page,
            fg_color=CARD_BG,
            corner_radius=13,
            border_width=1,
            border_color=BORDER,
        )

        header_card.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 10),
        )

        header_card.grid_columnconfigure(
            0,
            weight=1,
        )

        title_area = ctk.CTkFrame(
            header_card,
            fg_color="transparent",
        )

        title_area.grid(
            row=0,
            column=0,
            sticky="w",
            padx=20,
            pady=14,
        )

        ctk.CTkLabel(
            title_area,
            text="Revisión de metadatos",
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=17,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ctk.CTkLabel(
            title_area,
            text=self.selected_docx.name,
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).grid(
            row=1,
            column=0,
            sticky="w",
            pady=(2, 0),
        )

        self.secondary_button(
            header_card,
            "Cambiar artículo",
            self.show_selection_page,
            width=145,
        ).grid(
            row=0,
            column=1,
            padx=20,
            pady=14,
        )

        scroll = ctk.CTkScrollableFrame(
            page,
            fg_color=APP_BG,
            corner_radius=0,
            scrollbar_button_color=SCROLLBAR,
            scrollbar_button_hover_color=(
                SCROLLBAR_HOVER
            ),
        )

        scroll.grid(
            row=1,
            column=0,
            sticky="nsew",
        )

        scroll.grid_columnconfigure(
            0,
            weight=1,
        )

        self.metadata_form = MetadataForm(
            scroll,
            self.current_metadata,
        )

        self.metadata_form.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=3,
            pady=(3, 8),
        )

        self._bind_metadata_changes()

        footer = ctk.CTkFrame(
            page,
            fg_color=CARD_BG,
            corner_radius=13,
            border_width=1,
            border_color=BORDER,
        )

        footer.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(10, 0),
        )

        footer.grid_columnconfigure(
            1,
            weight=1,
        )

        self.secondary_button(
            footer,
            "← Volver",
            self.show_selection_page,
            width=110,
        ).grid(
            row=0,
            column=0,
            padx=16,
            pady=12,
        )

        status_frame = ctk.CTkFrame(
            footer,
            fg_color="transparent",
        )

        status_frame.grid(
            row=0,
            column=1,
            sticky="ew",
            padx=10,
        )

        self.validation_label = ctk.CTkLabel(
            status_frame,
            textvariable=self.validation_variable,
            text_color=TEXT_SECONDARY,
            anchor="w",
            justify="left",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
        )

        self.validation_label.pack(
            fill="x"
        )

        actions = ctk.CTkFrame(
            footer,
            fg_color="transparent",
        )

        actions.grid(
            row=0,
            column=2,
            padx=16,
            pady=12,
        )

        self.validate_button = (
            self.secondary_button(
                actions,
                "Validar metadatos",
                self.validate_metadata_preview,
                width=160,
            )
        )

        self.validate_button.grid(
            row=0,
            column=0,
            padx=(0, 8),
        )

        self.convert_button = (
            self.primary_button(
                actions,
                "Convertir artículo  →",
                self.start_conversion,
                width=180,
            )
        )

        self.convert_button.grid(
            row=0,
            column=1,
        )

        self.convert_button.configure(
            state="disabled",
            fg_color=BUTTON_DISABLED,
            text_color=BUTTON_DISABLED_TEXT,
        )

        self.validation_variable.set(
            (
                "Revise los datos y pulse "
                "Validar metadatos."
            )
        )

        self.status_variable.set(
            (
                "Revise los metadatos antes "
                "de convertir."
            )
        )

    def _bind_metadata_changes(self) -> None:
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

    def _on_metadata_changed(
        self,
        *_args: object,
    ) -> None:
        """
        Invalida una validación anterior.
        """

        if self.current_config_result is None:
            return

        self.current_config_result = None

        self.validation_variable.set(
            "Hay cambios sin validar."
        )

        if self.validation_label is not None:
            self.validation_label.configure(
                text_color=WARNING
            )

        if self.convert_button is not None:
            self.convert_button.configure(
                state="disabled",
                fg_color=BUTTON_DISABLED,
                text_color=BUTTON_DISABLED_TEXT,
            )

    def _get_form_overrides(
        self,
    ) -> tuple[
        dict[str, str],
        dict[str, str],
        dict[str, dict[str, str]],
    ]:
        """
        Obtiene valores del formulario.
        """

        if self.metadata_form is None:
            raise RuntimeError(
                (
                    "El formulario de metadatos "
                    "no está disponible."
                )
            )

        return (
            self.metadata_form
            .get_article_overrides(),
            self.metadata_form
            .get_license_overrides(),
            self.metadata_form
            .get_author_overrides(),
        )

    def _friendly_validation_errors(
        self,
        errors: list[str],
    ) -> list[str]:
        """
        Convierte mensajes técnicos de validación
        en instrucciones claras para el usuario.
        """

        friendly: list[str] = []

        def add(
            message: str,
        ) -> None:
            if message not in friendly:
                friendly.append(
                    message
                )

        for error in errors:
            normalized = error.casefold()

            if any(
                field in normalized
                for field in (
                    "article.first_page",
                    "article.last_page",
                    "article.page_range",
                    "article.elocation_id",
                )
            ):
                add(
                    (
                        "Capture la paginación del artículo "
                        "(página inicial, página final "
                        "y rango)."
                    )
                )

                continue

            field_messages = (
                (
                    "article.doi",
                    "Capture el DOI del artículo.",
                ),
                (
                    "article.category",
                    (
                        "Seleccione la categoría "
                        "del artículo."
                    ),
                ),
                (
                    "article.volume",
                    (
                        "Capture el volumen "
                        "de la revista."
                    ),
                ),
                (
                    "article.issue",
                    (
                        "Capture el número "
                        "de la revista."
                    ),
                ),
                (
                    "article.published_date",
                    (
                        "Revise la fecha "
                        "de publicación."
                    ),
                ),
                (
                    "article.collection_date",
                    (
                        "Revise el año "
                        "de publicación."
                    ),
                ),
                (
                    "article.received_date",
                    (
                        "Revise la fecha "
                        "de recepción."
                    ),
                ),
                (
                    "article.accepted_date",
                    (
                        "Revise la fecha "
                        "de aceptación."
                    ),
                ),
                (
                    "journal.title",
                    (
                        "Revise el título "
                        "de la revista."
                    ),
                ),
                (
                    "journal.abbreviated_title",
                    (
                        "Revise el título abreviado "
                        "de la revista."
                    ),
                ),
                (
                    "journal.issn",
                    (
                        "Revise el ISSN "
                        "de la revista."
                    ),
                ),
                (
                    "journal.publisher_name",
                    (
                        "Revise el nombre de la "
                        "institución editora."
                    ),
                ),
                (
                    "journal.publisher_location",
                    (
                        "Revise la ubicación de la "
                        "institución editora."
                    ),
                ),
                (
                    "license.url",
                    (
                        "Revise el enlace de la "
                        "licencia Creative Commons."
                    ),
                ),
                (
                    "license.text",
                    (
                        "Revise el texto de la "
                        "licencia del artículo."
                    ),
                ),
                (
                    "license.copyright_year",
                    (
                        "Revise el año "
                        "de copyright."
                    ),
                ),
                (
                    "license.copyright_holder",
                    (
                        "Revise el titular de los "
                        "derechos de autor."
                    ),
                ),
            )

            matched = False

            for (
                technical_name,
                message,
            ) in field_messages:
                if (
                    technical_name
                    in normalized
                ):
                    add(
                        message
                    )

                    matched = True
                    break

            if matched:
                continue

            if (
                "doi no tiene un formato válido"
                in normalized
            ):
                add(
                    (
                        "Revise el formato "
                        "del DOI del artículo."
                    )
                )

                continue

            if "issn" in normalized:
                add(
                    (
                        "Revise el ISSN "
                        "de la revista."
                    )
                )

                continue

            if "orcid" in normalized:
                author_name = ""

                parts = error.split(
                    "'"
                )

                if len(parts) >= 3:
                    author_name = (
                        parts[1].strip()
                    )

                if author_name:
                    add(
                        (
                            "Revise el ORCID de "
                            f"{author_name}."
                        )
                    )
                else:
                    add(
                        (
                            "Revise el ORCID "
                            "de los autores."
                        )
                    )

                continue

            if (
                "correo de '"
                in normalized
            ):
                author_name = ""

                parts = error.split(
                    "'"
                )

                if len(parts) >= 3:
                    author_name = (
                        parts[1].strip()
                    )

                if author_name:
                    add(
                        (
                            "Revise el correo electrónico "
                            f"de {author_name}."
                        )
                    )
                else:
                    add(
                        (
                            "Revise los correos electrónicos "
                            "de los autores."
                        )
                    )

                continue

            if (
                "sección 'authors' no contiene autores"
                in normalized
            ):
                add(
                    (
                        "Agregue al menos un autor "
                        "al artículo."
                    )
                )

                continue

            if (
                "autor sin nombre"
                in normalized
            ):
                add(
                    (
                        "Revise el nombre "
                        "de los autores."
                    )
                )

                continue

            if (
                "url válida"
                in normalized
            ):
                add(
                    (
                        "Revise los enlaces "
                        "de los metadatos."
                    )
                )

                continue

            add(
                (
                    "Revise la información editorial "
                    "marcada como incompleta "
                    "en el formulario."
                )
            )

        return friendly

    def _validate_current_form(
        self,
        show_success_message: bool,
    ) -> OfficeConfigBuildResult | None:
        """
        Valida metadatos.
        """

        if (
            self.metadata_form is None
            or self.current_metadata is None
        ):
            return None

        try:
            (
                article_overrides,
                license_overrides,
                author_overrides,
            ) = self._get_form_overrides()

            manual_author_errors = (
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

                error_text = "\n".join(
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
                        "continuar:\n\n"
                        f"{error_text}"
                    ),
                    kind="warning",
                )

                self.status_variable.set(
                    "Revise los autores agregados."
                )

                return None

            result = (
                build_and_validate_office_config(
                    metadata=self.current_metadata,
                    article_overrides=(
                        article_overrides
                    ),
                    license_overrides=(
                        license_overrides
                    ),
                    author_overrides=(
                        author_overrides
                    ),
                )
            )

        except Exception as error:
            self.validation_variable.set(
                (
                    "No fue posible validar "
                    "los metadatos."
                )
            )

            if (
                self.validation_label
                is not None
            ):
                self.validation_label.configure(
                    text_color=ERROR
                )

            if (
                self.convert_button
                is not None
            ):
                self.convert_button.configure(
                    state="disabled",
                    fg_color=BUTTON_DISABLED,
                    text_color=(
                        BUTTON_DISABLED_TEXT
                    ),
                )

            ModernDialog.show(
                self,
                "Error de validación",
                (
                    "Ocurrió un problema al validar "
                    "la información.\n\n"
                    f"{type(error).__name__}: "
                    f"{error}"
                ),
                kind="error",
            )

            return None

        self.current_config_result = (
            result
        )

        if result.valid:
            try:
                self.metadata_detection_events = (
                    compare_metadata_detection(
                        metadata=self.current_metadata,
                        article_overrides=article_overrides,
                        license_overrides=license_overrides,
                        author_overrides=author_overrides,
                    )
                )

            except Exception:
                # La telemetría nunca debe impedir una
                # validación o conversión correcta.
                self.metadata_detection_events = []

            self.validation_variable.set(
                (
                    "✓ Metadatos validados. "
                    "El artículo está listo."
                )
            )

            if (
                self.validation_label
                is not None
            ):
                self.validation_label.configure(
                    text_color=SUCCESS
                )

            if (
                self.convert_button
                is not None
            ):
                self.convert_button.configure(
                    state="normal",
                    fg_color=FACPYA_RED,
                    text_color="#FFFFFF",
                )

            self.status_variable.set(
                (
                    "Metadatos válidos. "
                    "Artículo listo para convertir."
                )
            )

            if show_success_message:
                ModernDialog.show(
                    self,
                    "Metadatos validados",
                    (
                        "La información editorial "
                        "es válida.\n\n"
                        "El artículo está listo para "
                        "generar XML JATS, HTML, PDF, "
                        "EPUB e imágenes."
                    ),
                    kind="success",
                )

            return result

        errors: list[str] = []

        if (
            result.validation
            is not None
        ):
            errors = (
                self._friendly_validation_errors(
                    list(
                        result.validation.errors
                    )
                )
            )

        self.validation_variable.set(
            "Hay datos que deben corregirse."
        )

        if (
            self.validation_label
            is not None
        ):
            self.validation_label.configure(
                text_color=ERROR
            )

        if (
            self.convert_button
            is not None
        ):
            self.convert_button.configure(
                state="disabled",
                fg_color=BUTTON_DISABLED,
                text_color=(
                    BUTTON_DISABLED_TEXT
                ),
            )

        error_text = (
            "\n".join(
                f"• {error}"
                for error in errors
            )
            if errors
            else (
                "La configuración no pasó "
                "la validación."
            )
        )

        ModernDialog.show(
            self,
            "Revise los metadatos",
            (
                "Hay información que debe revisarse "
                "antes de convertir:\n\n"
                f"{error_text}\n\n"
                "Complete estos datos en el formulario "
                "y vuelva a validar."
            ),
            kind="warning",
        )

        self.status_variable.set(
            (
                "La validación encontró "
                "datos por corregir."
            )
        )

        return result

    def validate_metadata_preview(
        self,
    ) -> None:
        """
        Botón de validación.
        """

        if (
            self.validate_button
            is not None
        ):
            self.validate_button.configure(
                state="disabled"
            )

        self.status_variable.set(
            "Validando metadatos..."
        )

        self.update_idletasks()

        try:
            self._validate_current_form(
                show_success_message=True
            )

        finally:
            if (
                self.validate_button
                is not None
            ):
                self.validate_button.configure(
                    state="normal"
                )

    def confirm_conversion(
        self,
    ) -> bool | None:
        """
        Confirma la conversión.

        Devuelve True si el usuario confirma,
        False si cancela y None si no fue posible
        acceder a la carpeta compartida.
        """

        if self.selected_docx is None:
            return False

        article_name = (
            self.selected_docx.stem
        )

        try:
            result_directory = (
                get_results_root()
                / article_name
            )

            result_exists = (
                result_directory.exists()
            )

        except OSError:
            self.status_variable.set(
                (
                    "No fue posible acceder a la "
                    "carpeta compartida de resultados."
                )
            )

            ModernDialog.show(
                self,
                "Carpeta compartida no disponible",
                (
                    "No fue posible acceder a la carpeta "
                    "compartida de resultados.\n\n"
                    "Verifique que el equipo esté "
                    "conectado a la red institucional "
                    "y que tenga permisos de lectura "
                    "y escritura.\n\n"
                    "No se inició ninguna conversión."
                ),
                kind="error",
            )

            return None

        if result_exists:
            return ModernDialog.ask(
                parent=self,
                title="Reemplazar resultados",
                message=(
                    "Ya existen archivos generados "
                    "para:\n\n"
                    f"{article_name}\n\n"
                    "La nueva conversión reemplazará "
                    "los resultados anteriores."
                ),
                confirm_text=(
                    "Reemplazar y convertir"
                ),
                cancel_text="Cancelar",
                kind="warning",
                destructive=True,
            )

        return ModernDialog.ask(
            parent=self,
            title="Confirmar conversión",
            message=(
                "Se generarán los archivos de:\n\n"
                f"{article_name}\n\n"
                "XML JATS · HTML · PDF · EPUB · "
                "Imágenes\n\n"
                "Espere hasta que aparezca "
                "la pantalla de resultados."
            ),
            confirm_text="Generar archivos",
            cancel_text="Cancelar",
            kind="info",
        )

    def start_conversion(
        self,
    ) -> None:
        """
        Inicia conversión.
        """

        if (
            self.selected_docx is None
            or self.current_metadata is None
            or self.metadata_form is None
        ):
            return

        validation_result = (
            self._validate_current_form(
                show_success_message=False
            )
        )

        if (
            validation_result is None
            or not validation_result.valid
        ):
            return

        confirmation_result = (
            self.confirm_conversion()
        )

        if confirmation_result is None:
            return

        if not confirmation_result:
            self.status_variable.set(
                "Conversión cancelada."
            )

            return

        (
            article_overrides,
            license_overrides,
            author_overrides,
        ) = self._get_form_overrides()

        docx_path = self.selected_docx

        self.is_converting = True

        self.show_conversion_progress_page()

        worker = threading.Thread(
            target=self._conversion_worker,
            args=(
                docx_path,
                article_overrides,
                license_overrides,
                author_overrides,
                list(
                    self.metadata_detection_events
                ),
            ),
            daemon=True,
        )

        worker.start()

    def show_conversion_progress_page(
        self,
    ) -> None:
        """
        Pantalla de progreso.
        """

        self.clear_content()
        self._set_step(3)

        self.status_variable.set(
            "Conversión en proceso..."
        )

        page = ctk.CTkFrame(
            self.content_host,
            fg_color=APP_BG,
        )

        page.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        page.grid_columnconfigure(
            0,
            weight=1,
        )

        page.grid_rowconfigure(
            0,
            weight=1,
        )

        card = ctk.CTkFrame(
            page,
            width=700,
            height=380,
            fg_color=CARD_BG,
            corner_radius=20,
            border_width=1,
            border_color=BORDER,
        )

        card.place(
            relx=0.5,
            rely=0.46,
            anchor="center",
        )

        card.grid_propagate(False)

        card.grid_columnconfigure(
            0,
            weight=1,
        )

        icon = ctk.CTkFrame(
            card,
            width=64,
            height=64,
            corner_radius=32,
            fg_color=FACPYA_RED_SOFT,
        )

        icon.grid(
            row=0,
            column=0,
            pady=(38, 14),
        )

        icon.grid_propagate(False)

        ctk.CTkLabel(
            icon,
            text="↻",
            text_color=FACPYA_RED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=28,
                weight="bold",
            ),
        ).place(
            relx=0.5,
            rely=0.5,
            anchor="center",
        )

        ctk.CTkLabel(
            card,
            text=(
                "Generando archivos del artículo"
            ),
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=20,
                weight="bold",
            ),
        ).grid(
            row=1,
            column=0,
        )

        article_name = (
            self.selected_docx.name
            if self.selected_docx
            else ""
        )

        ctk.CTkLabel(
            card,
            text=(
                f"Artículo en proceso: {article_name}"
                if article_name
                else "Artículo en proceso"
            ),
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=12,
                weight="bold",
            ),
        ).grid(
            row=2,
            column=0,
            pady=(6, 24),
        )

        self.progress_bar = (
            ctk.CTkProgressBar(
                card,
                width=530,
                height=10,
                corner_radius=5,
                mode="indeterminate",
                progress_color=FACPYA_RED,
                fg_color=CHIP_BG,
            )
        )

        self.progress_bar.grid(
            row=3,
            column=0,
            padx=75,
            sticky="ew",
        )

        self.progress_bar.start()

        self.conversion_status_variable.set(
            (
                "Procesando Word, figuras, tablas, "
                "XML JATS, HTML, PDF y EPUB..."
            )
        )

        ctk.CTkLabel(
            card,
            textvariable=(
                self.conversion_status_variable
            ),
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
        ).grid(
            row=4,
            column=0,
            pady=(20, 5),
        )

        ctk.CTkLabel(
            card,
            text=(
                "No cierre la aplicación ni abra "
                "el documento Word durante el proceso."
            ),
            text_color=TEXT_SECONDARY,
            wraplength=560,
            justify="center",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).grid(
            row=5,
            column=0,
            padx=50,
        )

    def _conversion_worker(
        self,
        docx_path: Path,
        article_overrides: dict[str, str],
        license_overrides: dict[str, str],
        author_overrides: dict[
            str,
            dict[str, str],
        ],
        metadata_detection_events: list[MetadataDetectionEvent],
    ) -> None:
        """
        Ejecuta backend fuera del hilo gráfico.
        """

        pythoncom_module = None

        try:
            import pythoncom

            pythoncom_module = pythoncom
            pythoncom.CoInitialize()

            result = run_office_conversion(
                docx_path=docx_path,
                article_overrides=(
                    article_overrides
                ),
                license_overrides=(
                    license_overrides
                ),
                author_overrides=(
                    author_overrides
                ),
                cleanup_run=True,
            metadata_detection_events=metadata_detection_events,
            )

        except Exception as error:
            result = OfficeGuiConversionResult(
                success=False,
                errors=[
                    (
                        f"{type(error).__name__}: "
                        f"{error}"
                    )
                ],
            )

        finally:
            if (
                pythoncom_module
                is not None
            ):
                try:
                    pythoncom_module.CoUninitialize()

                except Exception:
                    pass

        self.after(
            0,
            self._finish_conversion,
            result,
        )

    def _finish_conversion(
        self,
        result: OfficeGuiConversionResult,
    ) -> None:
        """
        Recibe resultado del backend.
        """

        if self.progress_bar is not None:
            try:
                self.progress_bar.stop()

            except tk.TclError:
                pass

        self.is_converting = False
        self.last_conversion_result = (
            result
        )

        if result.success:
            self.status_variable.set(
                (
                    "Conversión completada "
                    "correctamente."
                )
            )

            self.show_conversion_result_page(
                result
            )

            return

        self.status_variable.set(
            (
                "La conversión terminó "
                "con errores."
            )
        )

        self.show_conversion_error_page(
            result
        )

    def show_conversion_result_page(
        self,
        result: OfficeGuiConversionResult,
    ) -> None:
        """
        Dashboard final.
        """

        self.clear_content()
        self._set_step(4)

        page = ctk.CTkScrollableFrame(
            self.content_host,
            fg_color=APP_BG,
            corner_radius=0,
            scrollbar_button_color=SCROLLBAR,
            scrollbar_button_hover_color=(
                SCROLLBAR_HOVER
            ),
        )

        page.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=24,
            pady=16,
        )

        page.grid_columnconfigure(
            0,
            weight=1,
        )

        success_card = ctk.CTkFrame(
            page,
            fg_color=SUCCESS_SOFT,
            corner_radius=15,
            border_width=1,
            border_color=(
                "#ABEFC6",
                "#285842",
            ),
        )

        success_card.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(0, 14),
        )

        success_card.grid_columnconfigure(
            1,
            weight=1,
        )

        ctk.CTkLabel(
            success_card,
            text="✓",
            width=48,
            height=48,
            corner_radius=24,
            fg_color=SUCCESS,
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=21,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            rowspan=2,
            padx=20,
            pady=18,
        )

        ctk.CTkLabel(
            success_card,
            text="Conversión completada",
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=19,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=1,
            sticky="sw",
            pady=(17, 0),
        )

        article_name = (
            self.selected_docx.stem
            if self.selected_docx
            else "Artículo"
        )

        ctk.CTkLabel(
            success_card,
            text=(
                f"{article_name} · "
                "Los archivos fueron "
                "generados correctamente."
            ),
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).grid(
            row=1,
            column=1,
            sticky="nw",
            pady=(2, 17),
        )

        cards = ctk.CTkFrame(
            page,
            fg_color="transparent",
        )

        cards.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        for column in range(3):
            cards.grid_columnconfigure(
                column,
                weight=1,
                uniform="results",
            )

        self._create_result_card(
            cards,
            row=0,
            column=0,
            format_name="XML JATS",
            description=(
                "Documento estructurado "
                "para publicación."
            ),
            available=bool(
                result.xml_path
                and result.xml_path.is_file()
            ),
            action_text="Ver XML",
            command=self.open_result_xml,
        )

        self._create_result_card(
            cards,
            row=0,
            column=1,
            format_name="HTML",
            description=(
                "Versión web del "
                "artículo científico."
            ),
            available=bool(
                result.html_path
                and result.html_path.is_file()
            ),
            action_text="Abrir HTML",
            command=self.open_result_html,
        )

        self._create_result_card(
            cards,
            row=0,
            column=2,
            format_name="PDF",
            description=(
                "PDF exportado directamente "
                "desde Microsoft Word."
            ),
            available=bool(
                result.pdf_path
                and result.pdf_path.is_file()
            ),
            action_text="Abrir PDF",
            command=self.open_result_pdf,
        )

        self._create_result_card(
            cards,
            row=1,
            column=0,
            format_name="EPUB",
            description=(
                "Publicación electrónica "
                "en formato EPUB 3."
            ),
            available=bool(
                result.epub_path
                and result.epub_path.is_file()
            ),
            action_text="Mostrar EPUB",
            command=self.open_result_epub,
        )

        self._create_result_card(
            cards,
            row=1,
            column=1,
            format_name="IMÁGENES",
            description=(
                f"{result.media_files} archivos "
                "extraídos del documento."
            ),
            available=bool(
                result.media_directory
                and result.media_directory.is_dir()
            ),
            action_text="Ver imágenes",
            command=self.open_result_images,
        )

        summary_card = ctk.CTkFrame(
            cards,
            fg_color=CARD_BG,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
        )

        summary_card.grid(
            row=1,
            column=2,
            sticky="nsew",
            padx=6,
            pady=6,
        )

        ctk.CTkLabel(
            summary_card,
            text="RESUMEN",
            text_color=TEXT_MUTED,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        ).pack(
            fill="x",
            padx=20,
            pady=(18, 8),
        )

        ctk.CTkLabel(
            summary_card,
            text=(
                f"Errores        "
                f"{len(result.errors)}\n"
                f"Advertencias  "
                f"{len(result.warnings)}\n"
                f"Imágenes      "
                f"{result.media_files}"
            ),
            text_color=TEXT_PRIMARY,
            justify="left",
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).pack(
            fill="x",
            padx=20,
        )

        if result.warnings:
            self.secondary_button(
                summary_card,
                "Ver advertencias",
                self.show_result_warnings,
                width=145,
            ).pack(
                anchor="w",
                padx=20,
                pady=(12, 18),
            )

        path_card = ctk.CTkFrame(
            page,
            fg_color=CARD_BG,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
        )

        path_card.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(14, 0),
        )

        path_card.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            path_card,
            text=(
                "Archivos guardados correctamente"
            ),
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=13,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=20,
            pady=(17, 2),
        )

        result_path_text = (
            str(
                result.result_directory
            )
            if result.result_directory
            else "Ruta no disponible"
        )

        ctk.CTkLabel(
            path_card,
            text=result_path_text,
            text_color=TEXT_SECONDARY,
            anchor="w",
            justify="left",
            wraplength=760,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).grid(
            row=1,
            column=0,
            sticky="ew",
            padx=20,
            pady=(2, 17),
        )

        self.primary_button(
            path_card,
            "Abrir carpeta de resultados",
            self.open_result_directory,
            width=220,
        ).grid(
            row=0,
            column=1,
            rowspan=2,
            padx=20,
            pady=17,
        )

        bottom = ctk.CTkFrame(
            page,
            fg_color="transparent",
        )

        bottom.grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(16, 8),
        )

        bottom.grid_columnconfigure(
            0,
            weight=1,
        )

        self.secondary_button(
            bottom,
            "Procesar otro artículo",
            self.start_new_article,
            width=180,
        ).grid(
            row=0,
            column=1,
            sticky="e",
        )

    def _create_result_card(
        self,
        parent: tk.Misc,
        row: int,
        column: int,
        format_name: str,
        description: str,
        available: bool,
        action_text: str,
        command: Callable[[], None],
    ) -> None:
        """
        Tarjeta de resultado.
        """

        card = ctk.CTkFrame(
            parent,
            fg_color=CARD_BG,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
        )

        card.grid(
            row=row,
            column=column,
            sticky="nsew",
            padx=6,
            pady=6,
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
            padx=18,
            pady=(18, 6),
        )

        header.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            header,
            text=format_name,
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=14,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        ctk.CTkLabel(
            header,
            text=(
                "✓"
                if available
                else "×"
            ),
            width=28,
            height=28,
            corner_radius=14,
            fg_color=(
                SUCCESS_SOFT
                if available
                else ERROR_SOFT
            ),
            text_color=(
                SUCCESS
                if available
                else ERROR
            ),
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=12,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=1,
        )

        ctk.CTkLabel(
            card,
            text=description,
            text_color=TEXT_SECONDARY,
            anchor="nw",
            justify="left",
            wraplength=280,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).grid(
            row=1,
            column=0,
            sticky="ew",
            padx=18,
            pady=(0, 14),
        )

        action_button = (
            self.secondary_button(
                card,
                action_text,
                command,
                width=125,
            )
        )

        action_button.grid(
            row=2,
            column=0,
            sticky="w",
            padx=18,
            pady=(0, 18),
        )

        if not available:
            action_button.configure(
                state="disabled",
                fg_color=BUTTON_DISABLED,
                text_color=BUTTON_DISABLED_TEXT,
            )

    def show_result_warnings(
        self,
    ) -> None:
        """
        Muestra las advertencias no bloqueantes
        registradas durante la conversión.
        """

        result = self.last_conversion_result

        if result is None:
            ModernDialog.show(
                self,
                "Advertencias no disponibles",
                (
                    "No hay un resultado de conversión "
                    "disponible para consultar."
                ),
                kind="info",
            )

            return

        warnings: list[str] = []

        for warning in result.warnings:
            normalized = " ".join(
                str(
                    warning or ""
                ).split()
            )

            if (
                normalized
                and normalized not in warnings
            ):
                warnings.append(
                    normalized
                )

        if not warnings:
            ModernDialog.show(
                self,
                "Sin advertencias",
                (
                    "La conversión no registró "
                    "advertencias."
                ),
                kind="success",
            )

            return

        warning_text = "\n".join(
            f"• {warning}"
            for warning in warnings
        )

        ModernDialog.show(
            self,
            "Advertencias de la conversión",
            (
                "La conversión terminó correctamente, "
                "pero se registraron estas "
                "observaciones:\n\n"
                f"{warning_text}\n\n"
                "Estas advertencias no impidieron "
                "generar los archivos."
            ),
            kind="warning",
        )

    def open_result_html(
        self,
    ) -> None:
        """
        Abre HTML.
        """

        result = (
            self.last_conversion_result
        )

        if (
            result is None
            or result.html_path is None
            or not result.html_path.is_file()
        ):
            ModernDialog.show(
                self,
                "HTML no disponible",
                (
                    "No se encontró el archivo "
                    "HTML generado."
                ),
                kind="error",
            )

            return

        try:
            open_html(
                result.html_path
            )

        except Exception as error:
            ModernDialog.show(
                self,
                (
                    "No fue posible abrir "
                    "el HTML"
                ),
                str(error),
                kind="error",
            )

    def open_result_xml(
        self,
    ) -> None:
        """
        Abre XML.
        """

        result = (
            self.last_conversion_result
        )

        if (
            result is None
            or result.xml_path is None
            or not result.xml_path.is_file()
        ):
            ModernDialog.show(
                self,
                "XML no disponible",
                (
                    "No se encontró el archivo "
                    "XML generado."
                ),
                kind="error",
            )

            return

        try:
            open_xml_in_notepad(
                result.xml_path
            )

        except Exception as error:
            ModernDialog.show(
                self,
                (
                    "No fue posible abrir "
                    "el XML"
                ),
                str(error),
                kind="error",
            )

    def open_result_pdf(
        self,
    ) -> None:
        """
        Abre PDF.
        """

        result = (
            self.last_conversion_result
        )

        if (
            result is None
            or result.pdf_path is None
            or not result.pdf_path.is_file()
        ):
            ModernDialog.show(
                self,
                "PDF no disponible",
                (
                    "No se encontró el archivo "
                    "PDF generado."
                ),
                kind="error",
            )

            return

        try:
            os.startfile(
                str(
                    result.pdf_path
                )
            )

        except Exception as error:
            ModernDialog.show(
                self,
                (
                    "No fue posible abrir "
                    "el PDF"
                ),
                str(error),
                kind="error",
            )

    def open_result_epub(
        self,
    ) -> None:
        """
        Muestra el EPUB en el Explorador
        de Windows.
        """

        result = (
            self.last_conversion_result
        )

        if (
            result is None
            or result.epub_path is None
            or not result.epub_path.is_file()
        ):
            ModernDialog.show(
                self,
                "EPUB no disponible",
                (
                    "No se encontró el archivo "
                    "EPUB generado."
                ),
                kind="error",
            )

            return

        try:
            subprocess.Popen(
                [
                    "explorer.exe",
                    "/select,",
                    str(
                        result.epub_path
                    ),
                ]
            )

        except Exception as error:
            ModernDialog.show(
                self,
                (
                    "No fue posible mostrar "
                    "el EPUB"
                ),
                str(error),
                kind="error",
            )

    def open_result_images(
        self,
    ) -> None:
        """
        Abre carpeta de imágenes.
        """

        result = (
            self.last_conversion_result
        )

        if (
            result is None
            or result.media_directory is None
            or not result.media_directory.is_dir()
        ):
            ModernDialog.show(
                self,
                "Imágenes no disponibles",
                (
                    "No se encontró la carpeta "
                    "de imágenes."
                ),
                kind="error",
            )

            return

        try:
            open_directory(
                result.media_directory
            )

        except Exception as error:
            ModernDialog.show(
                self,
                (
                    "No fue posible abrir "
                    "las imágenes"
                ),
                str(error),
                kind="error",
            )

    def open_result_directory(
        self,
    ) -> None:
        """
        Abre carpeta principal.
        """

        result = (
            self.last_conversion_result
        )

        if (
            result is None
            or result.result_directory is None
            or not result.result_directory.is_dir()
        ):
            ModernDialog.show(
                self,
                "Carpeta no disponible",
                (
                    "No se encontró la carpeta "
                    "de resultados."
                ),
                kind="error",
            )

            return

        try:
            open_directory(
                result.result_directory
            )

        except Exception as error:
            ModernDialog.show(
                self,
                (
                    "No fue posible abrir "
                    "la carpeta"
                ),
                str(error),
                kind="error",
            )

    def show_conversion_error_page(
        self,
        result: OfficeGuiConversionResult,
    ) -> None:
        """
        Pantalla de error.
        """

        self.clear_content()
        self._set_step(3)

        page = ctk.CTkFrame(
            self.content_host,
            fg_color=APP_BG,
        )

        page.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        page.grid_columnconfigure(
            0,
            weight=1,
        )

        page.grid_rowconfigure(
            0,
            weight=1,
        )

        card = ctk.CTkFrame(
            page,
            width=740,
            fg_color=CARD_BG,
            corner_radius=18,
            border_width=1,
            border_color=BORDER,
        )

        card.place(
            relx=0.5,
            rely=0.44,
            anchor="center",
        )

        card.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            card,
            text="×",
            width=56,
            height=56,
            corner_radius=28,
            fg_color=ERROR_SOFT,
            text_color=ERROR,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=25,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            pady=(28, 12),
        )

        ctk.CTkLabel(
            card,
            text=(
                "No fue posible completar "
                "la conversión"
            ),
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=19,
                weight="bold",
            ),
        ).grid(
            row=1,
            column=0,
        )

        article_name = (
            self.selected_docx.stem
            if self.selected_docx
            else "Artículo"
        )

        ctk.CTkLabel(
            card,
            text=f"Artículo: {article_name}",
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=2,
            column=0,
            pady=(6, 0),
        )

        errors = (
            result.errors
            if result.errors
            else [
                (
                    "La conversión no terminó "
                    "correctamente."
                )
            ]
        )

        error_text = "\n".join(
            f"• {error}"
            for error in errors
        )

        error_box = ctk.CTkFrame(
            card,
            fg_color=ERROR_SOFT,
            corner_radius=10,
            border_width=1,
            border_color=(
                "#FECDCA",
                "#65302D",
            ),
        )

        error_box.grid(
            row=3,
            column=0,
            sticky="ew",
            padx=40,
            pady=20,
        )

        ctk.CTkLabel(
            error_box,
            text=error_text,
            text_color=ERROR,
            justify="left",
            anchor="w",
            wraplength=610,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).pack(
            fill="x",
            padx=16,
            pady=14,
        )

        actions = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        actions.grid(
            row=4,
            column=0,
            pady=(0, 28),
        )

        self.secondary_button(
            actions,
            "Volver a metadatos",
            self.show_metadata_page,
            width=165,
        ).grid(
            row=0,
            column=0,
            padx=5,
        )

        self.primary_button(
            actions,
            "Procesar otro artículo",
            self.start_new_article,
            width=180,
        ).grid(
            row=0,
            column=1,
            padx=5,
        )

    def start_new_article(
        self,
    ) -> None:
        """
        Reinicia el flujo.
        """

        self.selected_docx = None
        self.current_metadata = None
        self.metadata_detection_events = []
        self.current_config_result = None
        self.last_conversion_result = None

        self.validation_variable.set("")

        self.status_variable.set(
            (
                "Listo para seleccionar "
                "otro artículo."
            )
        )

        self.show_selection_page()

    def _on_close(
        self,
    ) -> None:
        """
        Impide cerrar durante conversión.
        """

        if self.is_converting:
            ModernDialog.show(
                self,
                "Conversión en proceso",
                (
                    "Espere a que finalice "
                    "la conversión antes de "
                    "cerrar la aplicación."
                ),
                kind="warning",
            )

            return

        self.destroy()

    def maximize_window(
        self,
    ) -> None:
        """
        Abre la ventana principal maximizada.

        En Windows se utiliza el estado "zoomed".
        Si el sistema no lo admite, conserva como
        respaldo el centrado tradicional.
        """

        try:
            self.state(
                "zoomed"
            )

        except tk.TclError:
            try:
                self.attributes(
                    "-zoomed",
                    True,
                )

            except tk.TclError:
                self.center_window()

    def center_window(
        self,
    ) -> None:
        """
        Centra la ventana.
        """

        self.update_idletasks()

        width = self.winfo_width()
        height = self.winfo_height()

        screen_width = (
            self.winfo_screenwidth()
        )

        screen_height = (
            self.winfo_screenheight()
        )

        x = max(
            0,
            (
                screen_width
                - width
            )
            // 2,
        )

        y = max(
            0,
            (
                screen_height
                - height
            )
            // 2,
        )

        self.geometry(
            f"{width}x{height}+{x}+{y}"
        )


def main() -> None:
    """
    Punto de entrada.
    """

    app = ScieloOfficeApp()
    app.mainloop()


if __name__ == "__main__":
    main()