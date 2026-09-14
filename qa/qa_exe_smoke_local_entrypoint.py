from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


def get_application_directory() -> Path:
    """
    En EXE:
        carpeta donde vive el ejecutable.

    En desarrollo:
        raíz del proyecto.
    """

    if getattr(
        sys,
        "frozen",
        False,
    ):
        return Path(
            sys.executable
        ).resolve().parent

    return (
        Path(
            __file__
        )
        .resolve()
        .parents[1]
    )


APPLICATION_DIRECTORY = (
    get_application_directory()
)

RESULTS_ROOT = (
    APPLICATION_DIRECTORY
    / "_SMOKE_RESULTADOS"
)

TECHNICAL_ROOT = (
    APPLICATION_DIRECTORY
    / "_SMOKE_TECNICO"
)


def assert_local_path(
    path: Path,
    label: str,
) -> None:
    """
    Defensa adicional:
    esta variante QA nunca debe apuntar
    a una ruta UNC o de producción.
    """

    resolved_text = str(
        path.resolve()
    )

    if resolved_text.startswith(
        "\\\\"
    ):
        raise RuntimeError(
            f"{label} apunta a una ruta UNC: "
            f"{resolved_text}"
        )

    forbidden = (
        "148.234.45.124",
        "Secretaria de Investigacion",
    )

    lowered = (
        resolved_text.casefold()
    )

    for value in forbidden:
        if (
            value.casefold()
            in lowered
        ):
            raise RuntimeError(
                f"{label} contiene una ruta "
                "de producción prohibida: "
                f"{resolved_text}"
            )


def prepare_local_directories() -> None:
    assert_local_path(
        RESULTS_ROOT,
        "RESULTS_ROOT",
    )

    assert_local_path(
        TECHNICAL_ROOT,
        "TECHNICAL_ROOT",
    )

    RESULTS_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    TECHNICAL_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    marker_path = (
        APPLICATION_DIRECTORY
        / "SMOKE_LOCAL_PATHS.txt"
    )

    marker_path.write_text(
        "\n".join(
            [
                "SCIELO FACPyA - SMOKE LOCAL",
                "",
                (
                    "RESULTS_ROOT="
                    + str(
                        RESULTS_ROOT
                    )
                ),
                (
                    "TECHNICAL_ROOT="
                    + str(
                        TECHNICAL_ROOT
                    )
                ),
                "",
                (
                    "UNC PRODUCCION: "
                    "BLOQUEADO POR ESTA VARIANTE QA"
                ),
            ]
        )
        + "\n",
        encoding="utf-8",
    )


def configure_smoke_guard() -> None:
    """
    Aplica tres defensas:

    1. Cambia SHARED_RESULTS_ROOT en el módulo
       de publicación.

    2. Sustituye get_results_root() usado
       directamente por la GUI.

    3. Envuelve run_office_conversion() para
       imponer results_root y technical_root
       locales aunque el caller los omita.
    """

    prepare_local_directories()

    import converter.office_result_manager as result_manager
    import office_app.main as app_module

    # Evita envolver dos veces las funciones si
    # configure_smoke_guard() se llama nuevamente
    # dentro del mismo proceso.
    if getattr(
        app_module,
        "_smoke_local_guard_installed",
        False,
    ):
        return

    real_run_office_conversion = (
        app_module.run_office_conversion
    )

    # ========================================================
    # DEFENSA 1
    # ========================================================

    result_manager.SHARED_RESULTS_ROOT = (
        RESULTS_ROOT
    )

    # ========================================================
    # DEFENSA 2
    # ========================================================

    def safe_get_results_root(
        results_root: str | Path | None = None,
    ) -> Path:
        """
        En esta variante QA se ignora cualquier
        destino proporcionado y se usa
        exclusivamente la carpeta local.
        """

        del results_root

        prepare_local_directories()

        return RESULTS_ROOT

    # ========================================================
    # DEFENSA 3
    # ========================================================

    def safe_run_office_conversion(
        *args: Any,
        **kwargs: Any,
    ) -> Any:
        """
        Fuerza publicación y reportes técnicos
        a las carpetas QA locales.
        """

        prepare_local_directories()

        kwargs[
            "results_root"
        ] = RESULTS_ROOT

        kwargs[
            "technical_root"
        ] = TECHNICAL_ROOT

        return real_run_office_conversion(
            *args,
            **kwargs,
        )

    app_module.get_results_root = (
        safe_get_results_root
    )

    app_module.run_office_conversion = (
        safe_run_office_conversion
    )

    app_module._smoke_local_guard_installed = (
        True
    )


def main() -> None:
    configure_smoke_guard()

    import office_app.main as app_module

    app = (
        app_module.ScieloOfficeApp()
    )

    app.title(
        "Convertidor SCIELO FACPyA "
        "- SMOKE LOCAL"
    )

    app.mainloop()


if __name__ == "__main__":
    main()