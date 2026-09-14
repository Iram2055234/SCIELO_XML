from __future__ import annotations

import hashlib
import json
import os
import shutil
import sys
from datetime import datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path.cwd()

SNAPSHOT_ROOT = ROOT / "qa" / "_release_candidates"

PROJECT_NAME = "SCIELO_XML"
RELEASE_NAME = "RC_2027_FINAL_POST_PILOTO_APROBADO"


INCLUDE_DIRECTORIES = [
    ROOT / "converter",
    ROOT / "office_app",
    ROOT / "models",
    ROOT / "templates",
    ROOT / "config",
    ROOT / "qa",
]


ROOT_FILE_PATTERNS = [
    "*.py",
    "*.json",
    "*.xsl",
    "*.xml",
    "*.txt",
    "*.md",
    "*.spec",
    "requirements*.txt",
    "pyproject.toml",
]


# Importante: estas carpetas se podan durante el recorrido; no solo se
# filtran los archivos al final. Esto evita entrar en backups anidados.
EXCLUDED_DIRECTORY_NAMES = {
    ".git",
    ".venv",
    "__pycache__",
    "_release_candidates",
    "_code_backups",
    "_pilot_backups",
    "_result_backups",
    "_pilot_deployments",
    "_e2e_1354",
    "salida_shapes_poc",
    "output",
    "dist",
    "build",
    "dist_rc2027",
    "build_rc2027",
    "dist_smoke_local",
    "build_smoke_local",
    "_SMOKE_RESULTADOS",
    "_SMOKE_TECNICO",
}


EXCLUDED_SUFFIXES = {
    ".pyc",
    ".pyo",
    ".exe",
    ".dll",
    ".docx",
    ".dotx",
    ".pdf",
    ".epub",
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".tif",
    ".tiff",
    ".zip",
}


QA_APPROVAL = {
    "post_pilot_regression": "REGRESIÓN POST-PILOTO RC 2027: APROBADA",
    "post_pilot_tests_defined": 4,
    "post_pilot_tests_executed": 4,
    "post_pilot_tests_approved": 4,
    "base_regression": "REGRESION FINAL FASE 2027: APROBADA",
    "figures": "APROBADAS",
    "tables": "APROBADAS",
    "lists": "APROBADAS",
    "equations_omml_mathml": "APROBADAS",
    "article_1354_end_to_end": "APROBADO",
    "article_1357_end_to_end": "APROBADO",
    "article_1360_end_to_end": "APROBADO",
    "article_1293_regression": "APROBADA",
    "history_workstations": "APROBADO",
    "pagination": "FUERA DE QA",
}


CRITICAL_FILES = [
    "converter/docx_reader.py",
    "converter/metadata_parser.py",
    "converter/body_parser.py",
    "converter/jats_generator.py",
    "converter/table_analyzer.py",
    "converter/table_extractor.py",
    "converter/table_position_integrator.py",
    "converter/visual_object_analyzer.py",
    "converter/figure_numbering_policy.py",
    "converter/com_figure_exporter.py",
    "converter/conversion_history.py",
    "converter/workstation_identity.py",
    "converter/pipeline.py",
    "office_app/conversion_runner.py",
    "office_app/main.py",
    "templates/article.xsl",
    "config/journal_defaults.json",
    "Convertidor SCIELO FACPyA.spec",
    "models/__init__.py",
    "models/article.py",
    "models/author.py",
    "models/affiliation.py",
    "models/figure.py",
    "models/reference.py",
    "models/section.py",
]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def should_exclude_file(path: Path) -> bool:
    if path.suffix.lower() in EXCLUDED_SUFFIXES:
        return True
    return False


def collect_directory_files(directory: Path) -> list[Path]:
    files: list[Path] = []
    if not directory.is_dir():
        return files

    for current_root, dirnames, filenames in os.walk(directory, topdown=True):
        # Poda real: os.walk no entrará a estas carpetas.
        dirnames[:] = [
            name
            for name in dirnames
            if name not in EXCLUDED_DIRECTORY_NAMES
        ]

        current_path = Path(current_root)
        for filename in filenames:
            path = current_path / filename
            if should_exclude_file(path):
                continue
            files.append(path.resolve())

    return files


def collect_files() -> list[Path]:
    files: set[Path] = set()

    for directory in INCLUDE_DIRECTORIES:
        files.update(collect_directory_files(directory))

    # Módulos Python de primer nivel: incluirlos explícitamente. Esto
    # protege dependencias importables de raíz como models.py.
    for path in ROOT.iterdir():
        if path.is_file() and path.suffix.lower() == ".py":
            files.add(path.resolve())

    for pattern in ROOT_FILE_PATTERNS:
        for path in ROOT.glob(pattern):
            if not path.is_file():
                continue
            if should_exclude_file(path):
                continue
            files.add(path.resolve())

    return sorted(
        files,
        key=lambda value: str(value.relative_to(ROOT)).casefold(),
    )


def validate_critical_files() -> None:
    missing = []
    for relative_name in CRITICAL_FILES:
        path = ROOT / Path(relative_name)
        if not path.is_file():
            missing.append(relative_name)

    if missing:
        raise RuntimeError(
            "Faltan archivos críticos:\n"
            + "\n".join(f" - {name}" for name in missing)
        )


def build_manifest(files: list[Path]) -> list[dict[str, object]]:
    manifest = []
    total = len(files)

    for index, path in enumerate(files, start=1):
        relative = path.relative_to(ROOT)
        manifest.append(
            {
                "path": relative.as_posix(),
                "size_bytes": path.stat().st_size,
                "sha256": sha256_file(path),
            }
        )
        if index % 100 == 0 or index == total:
            print(f"  Hashes: {index}/{total}", flush=True)

    return manifest


def copy_snapshot_files(files: list[Path], snapshot_directory: Path) -> None:
    source_directory = snapshot_directory / "source"
    total = len(files)

    for index, path in enumerate(files, start=1):
        relative = path.relative_to(ROOT)
        target = source_directory / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)
        if index % 100 == 0 or index == total:
            print(f"  Copia: {index}/{total}", flush=True)


def write_manifest_files(
    snapshot_directory: Path,
    manifest: list[dict[str, object]],
    created_at: str,
) -> None:
    manifest_json = {
        "project": PROJECT_NAME,
        "release": RELEASE_NAME,
        "created_at": created_at,
        "project_root": str(ROOT),
        "python": sys.executable,
        "qa_status": QA_APPROVAL,
        "critical_files": CRITICAL_FILES,
        "file_count": len(manifest),
        "files": manifest,
    }

    json_path = snapshot_directory / "MANIFEST_SHA256.json"
    json_path.write_text(
        json.dumps(manifest_json, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    txt_path = snapshot_directory / "MANIFEST_SHA256.txt"
    lines = [
        f"PROJECT: {PROJECT_NAME}",
        f"RELEASE: {RELEASE_NAME}",
        f"CREATED: {created_at}",
        "",
        "QA: REGRESIÓN POST-PILOTO RC 2027 APROBADA 4/4",
        "PAGINACION: FUERA DE QA",
        "",
        f"FILES: {len(manifest)}",
        "",
    ]
    for item in manifest:
        lines.append(f"{item['sha256']}  {item['path']}")
    txt_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    qa_path = snapshot_directory / "QA_APROBACION.txt"
    qa_lines = [
        "SCIELO_XML",
        "RELEASE CANDIDATE 2027 FINAL POST-PILOTO",
        "",
        "REGRESIÓN POST-PILOTO RC 2027: APROBADA",
        "Pruebas definidas: 4",
        "Pruebas ejecutadas: 4",
        "Pruebas aprobadas: 4",
        "",
        "REGRESION FINAL FASE 2027: APROBADA",
        "FIGURAS: APROBADAS",
        "TABLAS: APROBADAS",
        "LISTAS: APROBADAS",
        "ECUACIONES OMML/MATHML: APROBADAS",
        "1354 END-TO-END: APROBADO",
        "1357 END-TO-END: APROBADO",
        "1360 END-TO-END: APROBADO",
        "1293 REGRESION ESTRUCTURAL: APROBADA",
        "HISTORIAL + ESTACIONES: APROBADO",
        "",
        "PAGINACION: FUERA DE QA",
        "",
        "Este snapshot corresponde al estado final aprobado después del piloto,",
        "antes de construir y desplegar el EXE definitivo.",
    ]
    qa_path.write_text("\n".join(qa_lines) + "\n", encoding="utf-8")


def verify_snapshot(
    snapshot_directory: Path,
    manifest: list[dict[str, object]],
) -> None:
    source_directory = snapshot_directory / "source"
    errors = []
    total = len(manifest)

    for index, item in enumerate(manifest, start=1):
        relative = Path(str(item["path"]))
        snapshot_file = source_directory / relative
        if not snapshot_file.is_file():
            errors.append(f"FALTA: {relative}")
            continue
        actual_hash = sha256_file(snapshot_file)
        expected_hash = str(item["sha256"])
        if actual_hash != expected_hash:
            errors.append(f"HASH DISTINTO: {relative}")
        if index % 100 == 0 or index == total:
            print(f"  Verificación snapshot: {index}/{total}", flush=True)

    if errors:
        raise RuntimeError(
            "Falló la verificación del snapshot:\n" + "\n".join(errors)
        )


def create_zip(snapshot_directory: Path) -> Path:
    zip_path = snapshot_directory.parent / (snapshot_directory.name + ".zip")
    files = [path for path in snapshot_directory.rglob("*") if path.is_file()]
    total = len(files)

    with ZipFile(
        zip_path,
        "w",
        compression=ZIP_DEFLATED,
        compresslevel=6,
    ) as archive:
        for index, path in enumerate(files, start=1):
            archive_name = (
                Path(snapshot_directory.name)
                / path.relative_to(snapshot_directory)
            )
            archive.write(path, archive_name.as_posix())
            if index % 100 == 0 or index == total:
                print(f"  ZIP: {index}/{total}", flush=True)

    return zip_path


def verify_zip(
    zip_path: Path,
    manifest: list[dict[str, object]],
    snapshot_name: str,
) -> None:
    expected_paths = {
        (
            Path(snapshot_name)
            / "source"
            / Path(str(item["path"]))
        ).as_posix()
        for item in manifest
    }

    with ZipFile(zip_path, "r") as archive:
        bad_member = archive.testzip()
        if bad_member is not None:
            raise RuntimeError(f"CRC inválido en ZIP: {bad_member}")
        actual_paths = set(archive.namelist())

    missing = sorted(expected_paths - actual_paths)
    if missing:
        raise RuntimeError(
            "El ZIP no contiene todos los archivos esperados:\n"
            + "\n".join(missing[:20])
        )


def main() -> None:
    print("=" * 104, flush=True)
    print("CONGELADO RC 2027 FINAL POST-PILOTO", flush=True)
    print("=" * 104, flush=True)
    print("Buscando archivos (backups y releases temporales quedan excluidos)...", flush=True)

    validate_critical_files()
    files = collect_files()

    if not files:
        raise RuntimeError("No se encontraron archivos para congelar.")

    print(f"ARCHIVOS A CONGELAR: {len(files)}", flush=True)
    print("Calculando manifiesto SHA-256...", flush=True)

    now = datetime.now()
    timestamp = now.strftime("%Y%m%d_%H%M%S")
    created_at = now.isoformat(timespec="seconds")
    snapshot_name = f"{PROJECT_NAME}_{RELEASE_NAME}_{timestamp}"
    snapshot_directory = SNAPSHOT_ROOT / snapshot_name
    snapshot_directory.mkdir(parents=True, exist_ok=False)

    try:
        manifest = build_manifest(files)

        print("Copiando snapshot...", flush=True)
        copy_snapshot_files(files=files, snapshot_directory=snapshot_directory)

        source_directory = snapshot_directory / "source"
        root_python_files = [
            path.name
            for path in files
            if path.parent == ROOT and path.suffix.lower() == ".py"
        ]
        missing_root_python = [
            name
            for name in root_python_files
            if not (source_directory / name).is_file()
        ]
        if missing_root_python:
            raise RuntimeError(
                "El snapshot perdió módulos Python de raíz: "
                + ", ".join(sorted(missing_root_python))
            )
        models_package = source_directory / "models"
        required_models = [
            "__init__.py",
            "article.py",
            "author.py",
            "affiliation.py",
            "figure.py",
            "reference.py",
            "section.py",
        ]
        missing_models = [
            name
            for name in required_models
            if not (models_package / name).is_file()
        ]
        if missing_models:
            raise RuntimeError(
                "El snapshot no contiene el paquete models completo: "
                + ", ".join(missing_models)
            )
        print("PAQUETE MODELS EN SNAPSHOT: APROBADO", flush=True)

        print("Escribiendo manifiestos...", flush=True)
        write_manifest_files(
            snapshot_directory=snapshot_directory,
            manifest=manifest,
            created_at=created_at,
        )

        print("Verificando SHA-256 del snapshot...", flush=True)
        verify_snapshot(snapshot_directory=snapshot_directory, manifest=manifest)

        print("Creando ZIP...", flush=True)
        zip_path = create_zip(snapshot_directory)

        print("Verificando ZIP/CRC...", flush=True)
        verify_zip(
            zip_path=zip_path,
            manifest=manifest,
            snapshot_name=snapshot_name,
        )

    except BaseException:
        if snapshot_directory.exists():
            shutil.rmtree(snapshot_directory, ignore_errors=True)
        possible_zip = SNAPSHOT_ROOT / (snapshot_name + ".zip")
        if possible_zip.exists():
            possible_zip.unlink()
        raise

    manifest_by_path = {str(item["path"]): item for item in manifest}
    critical_report = []
    for relative_name in CRITICAL_FILES:
        item = manifest_by_path.get(relative_name)
        if item is None:
            raise RuntimeError(
                "Archivo crítico ausente del manifiesto: " + relative_name
            )
        critical_report.append((relative_name, str(item["sha256"])))

    print("=" * 104)
    print("RELEASE CANDIDATE 2027 FINAL — CONGELADO")
    print("=" * 104)
    print("SNAPSHOT:", snapshot_directory)
    print("ZIP:", zip_path)
    print()
    print("ARCHIVOS:", len(manifest))
    print("VERIFICACION SHA-256: APROBADA")
    print("VERIFICACION ZIP: APROBADA")
    print()
    print("ARCHIVOS CRITICOS")
    print("-" * 104)
    for relative_name, digest in critical_report:
        print(relative_name)
        print("  SHA256:", digest)
    print()
    print("QA REGISTRADO:")
    print("  REGRESIÓN POST-PILOTO RC 2027: APROBADA 4/4")
    print("  REGRESION FINAL FASE 2027: APROBADA")
    print("  FIGURAS: APROBADAS")
    print("  TABLAS: APROBADAS")
    print("  LISTAS: APROBADAS")
    print("  ECUACIONES OMML/MATHML: APROBADAS")
    print("  1354 END-TO-END: APROBADO")
    print("  1357 END-TO-END: APROBADO")
    print("  1360 END-TO-END: APROBADO")
    print("  1293 REGRESION: APROBADA")
    print("  HISTORIAL + ESTACIONES: APROBADO")
    print("  PAGINACION: FUERA DE QA")
    print()
    print("ESTADO APROBADO: CONGELADO")
    print("=" * 104)


if __name__ == "__main__":
    main()
