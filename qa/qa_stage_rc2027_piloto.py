from __future__ import annotations

import hashlib
import shutil
import sys
from pathlib import Path


ROOT = Path.cwd()

RC_ROOT = (
    ROOT
    / "dist_rc2027"
    / "Convertidor SCIELO FACPyA"
).resolve()

PILOT_PARENT = Path(
    r"\\148.234.45.124\Secretaria de Investigacion"
)

PILOT_ROOT = (
    PILOT_PARENT
    / "Convertidor SCIELO FACPyA - PILOTO"
)

STAGE_ROOT = (
    PILOT_PARENT
    / "Convertidor SCIELO FACPyA - PILOTO.__RC2027_STAGE__"
)

EXPECTED_ARTICLE_XSL_SHA256 = (
    "5903c4c652f3a79d5f4746a851c43ded"
    "b4d8771422b015184ea640ad829211b0"
)


def separator(title: str) -> None:
    print()
    print("=" * 108)
    print(title)
    print("=" * 108)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()

    with path.open("rb") as stream:
        while True:
            chunk = stream.read(
                1024 * 1024
            )

            if not chunk:
                break

            digest.update(chunk)

    return digest.hexdigest()


def collect_files(root: Path) -> list[Path]:
    return sorted(
        (
            path
            for path in root.rglob("*")
            if path.is_file()
        ),
        key=lambda path: (
            path.relative_to(root)
            .as_posix()
            .casefold()
        ),
    )


def validate_paths() -> None:
    if not RC_ROOT.is_dir():
        raise FileNotFoundError(
            f"No se encontró el RC local: {RC_ROOT}"
        )

    if not PILOT_ROOT.is_dir():
        raise FileNotFoundError(
            f"No se encontró el piloto actual: {PILOT_ROOT}"
        )

    if STAGE_ROOT.exists():
        raise FileExistsError(
            "Ya existe una carpeta de staging. "
            "No se eliminará automáticamente por seguridad: "
            f"{STAGE_ROOT}"
        )

    if (
        "SCIELO_RESULTADOS"
        in str(STAGE_ROOT).casefold()
    ):
        raise RuntimeError(
            "La ruta de staging no puede estar "
            "dentro de SCIELO_RESULTADOS."
        )


def validate_rc_identity() -> None:
    separator(
        "1. IDENTIDAD DEL RC LOCAL"
    )

    exe_path = (
        RC_ROOT
        / "Convertidor SCIELO FACPyA.exe"
    )

    xsl_path = (
        RC_ROOT
        / "_internal"
        / "templates"
        / "article.xsl"
    )

    config_path = (
        RC_ROOT
        / "_internal"
        / "config"
        / "journal_defaults.json"
    )

    for path in (
        exe_path,
        xsl_path,
        config_path,
    ):
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

    xsl_hash = sha256_file(
        xsl_path
    )

    print(
        "RC:",
        RC_ROOT,
    )

    print(
        "EXE:",
        exe_path.name,
    )

    print(
        "ARTICLE.XSL SHA256:",
        xsl_hash,
    )

    if (
        xsl_hash.casefold()
        != EXPECTED_ARTICLE_XSL_SHA256.casefold()
    ):
        raise RuntimeError(
            "article.xsl no coincide con "
            "el RC aprobado."
        )

    print(
        "IDENTIDAD ARTICLE.XSL: APROBADA"
    )


def copy_and_verify() -> None:
    separator(
        "2. INVENTARIO RC"
    )

    source_files = collect_files(
        RC_ROOT
    )

    source_bytes = sum(
        path.stat().st_size
        for path in source_files
    )

    print(
        "ARCHIVOS:",
        len(source_files),
    )

    print(
        "BYTES:",
        source_bytes,
    )

    if not source_files:
        raise RuntimeError(
            "El RC está vacío."
        )

    separator(
        "3. COPIA A STAGING UNC"
    )

    STAGE_ROOT.mkdir(
        parents=False,
        exist_ok=False,
    )

    for index, source_path in enumerate(
        source_files,
        start=1,
    ):
        relative = source_path.relative_to(
            RC_ROOT
        )

        target_path = (
            STAGE_ROOT
            / relative
        )

        target_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        shutil.copy2(
            source_path,
            target_path,
        )

        if (
            index % 100 == 0
            or index == len(source_files)
        ):
            print(
                f"{index} / "
                f"{len(source_files)} copiados"
            )

    separator(
        "4. VERIFICACION SHA-256"
    )

    staged_files = collect_files(
        STAGE_ROOT
    )

    staged_bytes = sum(
        path.stat().st_size
        for path in staged_files
    )

    print(
        "RC ARCHIVOS:",
        len(source_files),
    )

    print(
        "STAGE ARCHIVOS:",
        len(staged_files),
    )

    print(
        "RC BYTES:",
        source_bytes,
    )

    print(
        "STAGE BYTES:",
        staged_bytes,
    )

    if (
        len(source_files)
        != len(staged_files)
    ):
        raise RuntimeError(
            "Cantidad de archivos distinta "
            "entre RC y staging."
        )

    if (
        source_bytes
        != staged_bytes
    ):
        raise RuntimeError(
            "Total de bytes distinto "
            "entre RC y staging."
        )

    for index, source_path in enumerate(
        source_files,
        start=1,
    ):
        relative = source_path.relative_to(
            RC_ROOT
        )

        staged_path = (
            STAGE_ROOT
            / relative
        )

        if not staged_path.is_file():
            raise FileNotFoundError(
                staged_path
            )

        source_hash = sha256_file(
            source_path
        )

        staged_hash = sha256_file(
            staged_path
        )

        if (
            source_hash
            != staged_hash
        ):
            raise RuntimeError(
                "SHA-256 diferente: "
                f"{relative}"
            )

        if (
            index % 100 == 0
            or index == len(source_files)
        ):
            print(
                f"{index} / "
                f"{len(source_files)} verificados"
            )

    separator(
        "5. RECURSOS CRITICOS EN STAGING"
    )

    critical = [
        (
            "EXE",
            STAGE_ROOT
            / "Convertidor SCIELO FACPyA.exe",
        ),
        (
            "article.xsl",
            STAGE_ROOT
            / "_internal"
            / "templates"
            / "article.xsl",
        ),
        (
            "journal_defaults.json",
            STAGE_ROOT
            / "_internal"
            / "config"
            / "journal_defaults.json",
        ),
    ]

    for name, path in critical:
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

        print(
            name,
            "| PRESENTE |",
            sha256_file(path),
        )

    separator(
        "STAGING RC 2027 — APROBADO"
    )

    print(
        "RC LOCAL:",
        RC_ROOT,
    )

    print(
        "STAGING UNC:",
        STAGE_ROOT,
    )

    print(
        "ARCHIVOS:",
        len(source_files),
    )

    print(
        "BYTES:",
        source_bytes,
    )

    print(
        "SHA-256 RC -> STAGING: APROBADO"
    )

    print(
        "ARTICLE.XSL RC: APROBADO"
    )

    print()
    print(
        "PILOTO ACTUAL: SIN CAMBIOS"
    )

    print(
        "BACKUP PILOTO: SIN CAMBIOS"
    )

    print(
        "SCIELO_RESULTADOS: SIN CAMBIOS"
    )

    print()
    print(
        "RC LISTO PARA SWAP ATOMICO"
    )

    print(
        "=" * 108
    )


def main() -> None:
    separator(
        "STAGING DEL RC 2027 PARA PILOTO"
    )

    validate_paths()
    validate_rc_identity()

    try:
        copy_and_verify()

    except Exception:
        print()
        print(
            "ERROR DURANTE STAGING."
        )

        print(
            "El piloto actual NO fue modificado."
        )

        print(
            "La carpeta de staging se conserva "
            "para diagnóstico y NO se elimina "
            "automáticamente."
        )

        raise


if __name__ == "__main__":
    main()