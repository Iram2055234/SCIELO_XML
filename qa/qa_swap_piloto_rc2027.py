from __future__ import annotations

import hashlib
import json
import sys
from datetime import datetime
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

DEPLOYMENTS_ROOT = (
    ROOT
    / "qa"
    / "_pilot_deployments"
)

EXPECTED_ARTICLE_XSL_SHA256 = (
    "5903c4c652f3a79d5f4746a851c43ded"
    "b4d8771422b015184ea640ad829211b0"
)


def separator(
    title: str,
) -> None:
    print()
    print("=" * 108)
    print(title)
    print("=" * 108)


def sha256_file(
    path: Path,
) -> str:
    digest = hashlib.sha256()

    with path.open(
        "rb"
    ) as stream:
        while True:
            chunk = stream.read(
                1024 * 1024
            )

            if not chunk:
                break

            digest.update(
                chunk
            )

    return digest.hexdigest()


def collect_files(
    root: Path,
) -> list[Path]:
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


def get_inventory(
    root: Path,
) -> tuple[
    list[Path],
    int,
]:
    files = collect_files(
        root
    )

    total_bytes = sum(
        path.stat().st_size
        for path in files
    )

    return (
        files,
        total_bytes,
    )


def latest_local_backup() -> Path | None:
    backup_root = (
        ROOT
        / "qa"
        / "_pilot_backups"
    )

    if not backup_root.is_dir():
        return None

    candidates = sorted(
        (
            path
            for path in backup_root.iterdir()
            if (
                path.is_dir()
                and path.name.startswith(
                    "PILOTO_ANTES_RC2027_"
                )
            )
        ),
        key=lambda path: (
            path.stat().st_mtime
        ),
        reverse=True,
    )

    if not candidates:
        return None

    return candidates[0]


def validate_preflight() -> tuple[
    list[Path],
    int,
]:
    separator(
        "1. PREFLIGHT"
    )

    if not RC_ROOT.is_dir():
        raise FileNotFoundError(
            f"RC local no encontrado: {RC_ROOT}"
        )

    if not PILOT_ROOT.is_dir():
        raise FileNotFoundError(
            f"Piloto actual no encontrado: {PILOT_ROOT}"
        )

    if not STAGE_ROOT.is_dir():
        raise FileNotFoundError(
            f"Staging no encontrado: {STAGE_ROOT}"
        )

    rc_files, rc_bytes = (
        get_inventory(
            RC_ROOT
        )
    )

    stage_files, stage_bytes = (
        get_inventory(
            STAGE_ROOT
        )
    )

    print(
        "RC ARCHIVOS:",
        len(rc_files),
    )

    print(
        "RC BYTES:",
        rc_bytes,
    )

    print(
        "STAGE ARCHIVOS:",
        len(stage_files),
    )

    print(
        "STAGE BYTES:",
        stage_bytes,
    )

    if (
        len(rc_files)
        != len(stage_files)
    ):
        raise RuntimeError(
            "RC y staging tienen distinta "
            "cantidad de archivos."
        )

    if (
        rc_bytes
        != stage_bytes
    ):
        raise RuntimeError(
            "RC y staging tienen distinto "
            "total de bytes."
        )

    critical_relatives = [
        Path(
            "Convertidor SCIELO FACPyA.exe"
        ),
        (
            Path("_internal")
            / "templates"
            / "article.xsl"
        ),
        (
            Path("_internal")
            / "config"
            / "journal_defaults.json"
        ),
    ]

    separator(
        "2. RECURSOS CRITICOS"
    )

    for relative in critical_relatives:
        rc_path = (
            RC_ROOT
            / relative
        )

        stage_path = (
            STAGE_ROOT
            / relative
        )

        if not rc_path.is_file():
            raise FileNotFoundError(
                rc_path
            )

        if not stage_path.is_file():
            raise FileNotFoundError(
                stage_path
            )

        rc_hash = sha256_file(
            rc_path
        )

        stage_hash = sha256_file(
            stage_path
        )

        print(
            relative.as_posix()
        )

        print(
            "  RC:   ",
            rc_hash,
        )

        print(
            "  STAGE:",
            stage_hash,
        )

        if (
            rc_hash
            != stage_hash
        ):
            raise RuntimeError(
                "Hash crítico distinto: "
                f"{relative}"
            )

    xsl_path = (
        RC_ROOT
        / "_internal"
        / "templates"
        / "article.xsl"
    )

    xsl_hash = sha256_file(
        xsl_path
    )

    if (
        xsl_hash.casefold()
        != EXPECTED_ARTICLE_XSL_SHA256.casefold()
    ):
        raise RuntimeError(
            "article.xsl no corresponde "
            "al RC aprobado."
        )

    print()
    print(
        "PREFLIGHT: APROBADO"
    )

    return (
        rc_files,
        rc_bytes,
    )


def verify_deployed_rc(
    rc_files: list[Path],
    rc_bytes: int,
) -> None:
    separator(
        "4. VERIFICACION POST-SWAP"
    )

    if not PILOT_ROOT.is_dir():
        raise FileNotFoundError(
            "El nuevo piloto no existe."
        )

    deployed_files, deployed_bytes = (
        get_inventory(
            PILOT_ROOT
        )
    )

    print(
        "RC ARCHIVOS:",
        len(rc_files),
    )

    print(
        "PILOTO NUEVO ARCHIVOS:",
        len(deployed_files),
    )

    print(
        "RC BYTES:",
        rc_bytes,
    )

    print(
        "PILOTO NUEVO BYTES:",
        deployed_bytes,
    )

    if (
        len(rc_files)
        != len(deployed_files)
    ):
        raise RuntimeError(
            "Cantidad de archivos distinta "
            "después del swap."
        )

    if (
        rc_bytes
        != deployed_bytes
    ):
        raise RuntimeError(
            "Total de bytes distinto "
            "después del swap."
        )

    for index, rc_path in enumerate(
        rc_files,
        start=1,
    ):
        relative = (
            rc_path.relative_to(
                RC_ROOT
            )
        )

        deployed_path = (
            PILOT_ROOT
            / relative
        )

        if not deployed_path.is_file():
            raise FileNotFoundError(
                deployed_path
            )

        rc_hash = sha256_file(
            rc_path
        )

        deployed_hash = sha256_file(
            deployed_path
        )

        if (
            rc_hash
            != deployed_hash
        ):
            raise RuntimeError(
                "SHA-256 diferente "
                "después del swap: "
                f"{relative}"
            )

        if (
            index % 100 == 0
            or index == len(rc_files)
        ):
            print(
                f"{index} / "
                f"{len(rc_files)} verificados"
            )

    print()
    print(
        "SHA-256 RC -> PILOTO: APROBADO"
    )


def write_deployment_record(
    rollback_root: Path,
    rc_files: list[Path],
    rc_bytes: int,
) -> Path:
    DEPLOYMENTS_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    record_path = (
        DEPLOYMENTS_ROOT
        / (
            "DEPLOY_RC2027_"
            + timestamp
            + ".json"
        )
    )

    backup = (
        latest_local_backup()
    )

    record = {
        "deployed_at": (
            datetime.now()
            .astimezone()
            .isoformat(
                timespec="seconds"
            )
        ),
        "rc_local": str(
            RC_ROOT
        ),
        "pilot_current": str(
            PILOT_ROOT
        ),
        "rollback_server": str(
            rollback_root
        ),
        "local_backup": (
            str(backup)
            if backup is not None
            else ""
        ),
        "file_count": len(
            rc_files
        ),
        "total_bytes": (
            rc_bytes
        ),
        "article_xsl_sha256": (
            EXPECTED_ARTICLE_XSL_SHA256
        ),
        "status": (
            "DEPLOYED_AND_VERIFIED"
        ),
    }

    record_path.write_text(
        json.dumps(
            record,
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    return record_path


def main() -> None:
    separator(
        "SWAP PILOTO -> RC 2027"
    )

    rc_files, rc_bytes = (
        validate_preflight()
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    rollback_root = (
        PILOT_PARENT
        / (
            "Convertidor SCIELO FACPyA - PILOTO"
            ".__ROLLBACK_ANTES_RC2027_"
            + timestamp
            + "__"
        )
    )

    failed_rc_root = (
        PILOT_PARENT
        / (
            "Convertidor SCIELO FACPyA - PILOTO"
            ".__FAILED_RC2027_"
            + timestamp
            + "__"
        )
    )

    if rollback_root.exists():
        raise FileExistsError(
            rollback_root
        )

    if failed_rc_root.exists():
        raise FileExistsError(
            failed_rc_root
        )

    separator(
        "3. SWAP"
    )

    print(
        "PILOTO ACTUAL:"
    )

    print(
        " ",
        PILOT_ROOT,
    )

    print(
        "ROLLBACK:"
    )

    print(
        " ",
        rollback_root,
    )

    print(
        "STAGING:"
    )

    print(
        " ",
        STAGE_ROOT,
    )

    print()
    print(
        "Moviendo piloto actual -> rollback..."
    )

    PILOT_ROOT.rename(
        rollback_root
    )

    old_moved = True
    new_moved = False

    try:
        print(
            "Moviendo staging -> piloto..."
        )

        STAGE_ROOT.rename(
            PILOT_ROOT
        )

        new_moved = True

    except Exception:
        print()
        print(
            "ERROR AL ACTIVAR EL STAGING."
        )

        if (
            old_moved
            and rollback_root.exists()
            and not PILOT_ROOT.exists()
        ):
            print(
                "Restaurando piloto anterior..."
            )

            rollback_root.rename(
                PILOT_ROOT
            )

            print(
                "PILOTO ANTERIOR RESTAURADO."
            )

        raise

    try:
        verify_deployed_rc(
            rc_files=rc_files,
            rc_bytes=rc_bytes,
        )

    except Exception as verification_error:
        separator(
            "VERIFICACION FALLIDA — ROLLBACK"
        )

        print(
            "ERROR:",
            repr(
                verification_error
            ),
        )

        rollback_success = False

        if (
            new_moved
            and PILOT_ROOT.exists()
        ):
            try:
                print(
                    "Apartando RC fallido..."
                )

                PILOT_ROOT.rename(
                    failed_rc_root
                )

                print(
                    "RC fallido apartado en:"
                )

                print(
                    " ",
                    failed_rc_root,
                )

            except Exception as error:
                print(
                    "NO FUE POSIBLE APARTAR "
                    "EL RC FALLIDO:"
                )

                print(
                    repr(
                        error
                    )
                )

        if (
            rollback_root.exists()
            and not PILOT_ROOT.exists()
        ):
            try:
                print(
                    "Restaurando piloto anterior..."
                )

                rollback_root.rename(
                    PILOT_ROOT
                )

                rollback_success = True

                print(
                    "PILOTO ANTERIOR RESTAURADO."
                )

            except Exception as error:
                print(
                    "ERROR CRITICO AL RESTAURAR:"
                )

                print(
                    repr(
                        error
                    )
                )

        print()

        if rollback_success:
            print(
                "ROLLBACK AUTOMATICO: APROBADO"
            )
        else:
            print(
                "ROLLBACK AUTOMATICO: "
                "REQUIERE REVISION MANUAL"
            )

        raise

    record_path = (
        write_deployment_record(
            rollback_root=rollback_root,
            rc_files=rc_files,
            rc_bytes=rc_bytes,
        )
    )

    separator(
        "DESPLIEGUE PILOTO RC 2027 — APROBADO"
    )

    print(
        "PILOTO ACTIVO:"
    )

    print(
        " ",
        PILOT_ROOT,
    )

    print()

    print(
        "ROLLBACK INMEDIATO:"
    )

    print(
        " ",
        rollback_root,
    )

    print()

    print(
        "REGISTRO:"
    )

    print(
        " ",
        record_path,
    )

    print()

    print(
        "ARCHIVOS:",
        len(
            rc_files
        ),
    )

    print(
        "BYTES:",
        rc_bytes,
    )

    print(
        "SHA-256 RC -> PILOTO: APROBADO"
    )

    print(
        "ARTICLE.XSL RC 2027: APROBADO"
    )

    print()

    print(
        "ROLLBACK: CONSERVADO"
    )

    print(
        "BACKUP LOCAL: CONSERVADO"
    )

    print(
        "SCIELO_RESULTADOS: SIN CAMBIOS"
    )

    print()

    print(
        "NO ELIMINAR EL ROLLBACK "
        "HASTA TERMINAR QA DEL PILOTO."
    )

    print()

    print(
        "SWAP RC 2027: COMPLETADO"
    )

    print(
        "=" * 108
    )


if __name__ == "__main__":
    main()