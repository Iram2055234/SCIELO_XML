from __future__ import annotations

import customtkinter as ctk


# ============================================================
# TIPOGRAFÍA
# ============================================================

FONT_FAMILY = "Segoe UI"


# ============================================================
# COLORES ADAPTATIVOS
#
# Cada tupla contiene:
#   (modo claro, modo oscuro)
# ============================================================

APP_BG = (
    "#EEF1F5",
    "#101318",
)

CARD_BG = (
    "#FFFFFF",
    "#1B1F24",
)

SOFT_BG = (
    "#F7F9FC",
    "#232830",
)

ELEVATED_BG = (
    "#FFFFFF",
    "#20252C",
)

TEXT_PRIMARY = (
    "#182230",
    "#F3F5F7",
)

TEXT_SECONDARY = (
    "#667085",
    "#B2BAC5",
)

TEXT_MUTED = (
    "#98A2B3",
    "#7F8995",
)

BORDER = (
    "#DDE2E8",
    "#353C46",
)

BORDER_STRONG = (
    "#C9D0D8",
    "#454E59",
)


# ============================================================
# FACPYA
# ============================================================

FACPYA_RED = (
    "#941D2E",
    "#B52B40",
)

FACPYA_RED_HOVER = (
    "#781725",
    "#CA344B",
)

FACPYA_RED_SOFT = (
    "#FAECEF",
    "#351D22",
)


# ============================================================
# ESTADOS
# ============================================================

SUCCESS = (
    "#18794E",
    "#42B883",
)

SUCCESS_HOVER = (
    "#11643F",
    "#35966B",
)

SUCCESS_SOFT = (
    "#ECFDF3",
    "#173328",
)

WARNING = (
    "#B54708",
    "#F2A544",
)

WARNING_SOFT = (
    "#FFFAEB",
    "#382A17",
)

ERROR = (
    "#B42318",
    "#F97066",
)

ERROR_HOVER = (
    "#912018",
    "#E45D55",
)

ERROR_SOFT = (
    "#FEF3F2",
    "#3A2020",
)

INFO = (
    "#175CD3",
    "#6EA8FE",
)

INFO_SOFT = (
    "#EFF8FF",
    "#172B45",
)


# ============================================================
# CONTROLES
# ============================================================

INPUT_BG = (
    "#FFFFFF",
    "#171A1F",
)

BUTTON_DISABLED = (
    "#D0D5DD",
    "#343A43",
)

BUTTON_DISABLED_TEXT = (
    "#667085",
    "#777F89",
)

SECONDARY_HOVER = (
    "#F2F4F7",
    "#2B3139",
)

CHIP_BG = (
    "#F2F4F7",
    "#292F37",
)

SCROLLBAR = (
    "#C7CDD4",
    "#414954",
)

SCROLLBAR_HOVER = (
    "#98A2B3",
    "#59636F",
)


def configure_appearance() -> None:
    """
    Configuración inicial de CustomTkinter.
    """

    ctk.set_appearance_mode(
        "Light"
    )


def set_light_mode() -> None:
    """
    Activa modo claro.
    """

    ctk.set_appearance_mode(
        "Light"
    )


def set_dark_mode() -> None:
    """
    Activa modo oscuro.
    """

    ctk.set_appearance_mode(
        "Dark"
    )