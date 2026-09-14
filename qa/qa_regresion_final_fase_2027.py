from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import NamedTuple


ROOT = Path.cwd()


class TestDefinition(NamedTuple):
    name: str
    module: str
    approval_marker: str


TESTS = [
    TestDefinition(
        name="Contrato 2027 / 1354",
        module="qa.qa_1354_contrato_2027",
        approval_marker=(
            "CONTRATO 2027 / 1354: APROBADO"
        ),
    ),
    TestDefinition(
        name="Parser + JATS / lote completo",
        module="qa.qa_regresion_parser_jats_lote",
        approval_marker=(
            "REGRESION PARSER + JATS: APROBADA"
        ),
    ),
    TestDefinition(
        name="Figuras COM / 12 pilotos",
        module="qa.qa_regresion_figuras_lote",
        approval_marker=(
            "REGRESION FIGURAS LOTE: APROBADA"
        ),
    ),
    TestDefinition(
        name="1354 / end-to-end",
        module="qa.qa_1354_end_to_end_2027",
        approval_marker=(
            "END-TO-END 1354: APROBADO"
        ),
    ),
    TestDefinition(
        name="1196 / MathML HTML5",
        module="qa.qa_1196_mathml_html5_visual",
        approval_marker=(
            "MATHML HTML5 1196: APROBADO"
        ),
    ),
]


def separator(
    title: str,
    char: str = "=",
    width: int = 104,
) -> None:
    print()
    print(
        char * width
    )

    print(
        title
    )

    print(
        char * width
    )


def check_module_exists(
    module_name: str,
) -> None:
    spec = importlib.util.find_spec(
        module_name
    )

    if spec is None:
        raise ModuleNotFoundError(
            "No encontré el módulo QA requerido: "
            f"{module_name}"
        )


def run_test(
    number: int,
    definition: TestDefinition,
) -> dict[str, object]:
    separator(
        (
            f"PRUEBA {number}/{len(TESTS)}"
            f" — {definition.name}"
        ),
        char="-",
    )

    print(
        "MODULO:",
        definition.module,
    )

    print(
        "MARCADOR ESPERADO:",
        repr(
            definition.approval_marker
        ),
    )

    print()

    check_module_exists(
        definition.module
    )

    start = time.monotonic()

    result = subprocess.run(
        [
            sys.executable,
            "-B",
            "-m",
            definition.module,
        ],
        cwd=str(
            ROOT
        ),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        env={
            **os.environ,
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8": "1",
        },
    )

    elapsed = (
        time.monotonic()
        - start
    )

    stdout = (
        result.stdout
        or ""
    )

    stderr = (
        result.stderr
        or ""
    )

    if stdout:
        print(
            stdout.rstrip()
        )

    if stderr:
        print()
        print(
            "STDERR:"
        )

        print(
            stderr.rstrip()
        )

    marker_found = (
        definition.approval_marker
        in stdout
    )

    return_code_ok = (
        result.returncode
        == 0
    )

    approved = (
        marker_found
        and return_code_ok
    )

    print()
    print(
        "RETURN CODE:",
        result.returncode,
    )

    print(
        "MARCADOR:",
        (
            "ENCONTRADO"
            if marker_found
            else "NO ENCONTRADO"
        ),
    )

    print(
        "TIEMPO:",
        f"{elapsed:.1f} s",
    )

    print(
        "RESULTADO:",
        (
            "APROBADO"
            if approved
            else "ERROR/REVISAR"
        ),
    )

    return {
        "name": definition.name,
        "module": definition.module,
        "approved": approved,
        "return_code": (
            result.returncode
        ),
        "marker_found": (
            marker_found
        ),
        "elapsed": elapsed,
    }


def validate_current_xslt() -> None:
    """
    Validación mínima adicional del estado final
    de article.xsl.

    No modifica ningún archivo.
    """

    from lxml import etree

    xsl_path = (
        ROOT
        / "templates"
        / "article.xsl"
    )

    if not xsl_path.is_file():
        raise FileNotFoundError(
            xsl_path
        )

    xsl_namespace = (
        "http://www.w3.org/1999/XSL/Transform"
    )

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    tree = etree.parse(
        str(
            xsl_path
        ),
        parser,
    )

    # Compilación real.
    etree.XSLT(
        tree
    )

    ns = {
        "xsl": xsl_namespace,
    }

    list_templates = tree.xpath(
        (
            ".//xsl:template"
            "[@match='list']"
        ),
        namespaces=ns,
    )

    list_item_templates = tree.xpath(
        (
            ".//xsl:template"
            "[@match='list-item']"
        ),
        namespaces=ns,
    )

    formula_templates = tree.xpath(
        (
            ".//xsl:template"
            "[@match='disp-formula']"
        ),
        namespaces=ns,
    )

    math_templates = tree.xpath(
        (
            ".//xsl:template"
            "[@mode='html-math']"
        ),
        namespaces=ns,
    )

    sec_templates = tree.xpath(
        (
            ".//xsl:template"
            "[@match='sec']"
        ),
        namespaces=ns,
    )

    if len(
        list_templates
    ) != 1:
        raise RuntimeError(
            "article.xsl debe contener exactamente "
            "un template match='list'."
        )

    if len(
        list_item_templates
    ) != 1:
        raise RuntimeError(
            "article.xsl debe contener exactamente "
            "un template match='list-item'."
        )

    if len(
        formula_templates
    ) != 1:
        raise RuntimeError(
            "article.xsl debe contener exactamente "
            "un template match='disp-formula'."
        )

    if len(
        math_templates
    ) != 1:
        raise RuntimeError(
            "article.xsl debe contener exactamente "
            "un template mode='html-math'."
        )

    if len(
        sec_templates
    ) != 1:
        raise RuntimeError(
            "article.xsl debe contener exactamente "
            "un template match='sec'."
        )

    apply_nodes = sec_templates[
        0
    ].xpath(
        ".//xsl:apply-templates",
        namespaces=ns,
    )

    selects = [
        str(
            node.get(
                "select",
                "",
            )
        )
        for node in apply_nodes
        if node.get(
            "select"
        )
    ]

    content_select = next(
        (
            value
            for value in selects
            if (
                "table-wrap" in value
                and "sec" in value
            )
        ),
        "",
    )

    required = [
        "p",
        "table-wrap",
        "sec",
        "list",
        "disp-formula",
    ]

    missing = [
        value
        for value in required
        if value not in content_select
    ]

    if missing:
        raise RuntimeError(
            "El flujo XSLT de <sec> perdió "
            "tipos de contenido: "
            + ", ".join(
                missing
            )
        )

    print(
        "article.xsl compile: APROBADO"
    )

    print(
        "template list: 1 / 1 | APROBADO"
    )

    print(
        "template list-item: 1 / 1 | APROBADO"
    )

    print(
        "template disp-formula: 1 / 1 | APROBADO"
    )

    print(
        "mode html-math: 1 / 1 | APROBADO"
    )

    print(
        "sec select:",
        repr(
            content_select
        ),
    )


def main() -> None:
    separator(
        "REGRESION FINAL — FASE 2027"
    )

    print(
        "PROYECTO:",
        ROOT,
    )

    print(
        "PYTHON:",
        sys.executable,
    )

    print()

    print(
        "ESTA PRUEBA NO MODIFICA "
        "CONVERTER/PRODUCCION."
    )

    print(
        "PAGINACION: FUERA DE ALCANCE."
    )

    results = []

    total_start = (
        time.monotonic()
    )

    for number, definition in enumerate(
        TESTS,
        start=1,
    ):
        try:
            result = run_test(
                number=number,
                definition=definition,
            )

        except Exception as error:
            print()
            print(
                "RESULTADO: ERROR/REVISAR"
            )

            print(
                "EXCEPCION:",
                repr(
                    error
                ),
            )

            result = {
                "name": (
                    definition.name
                ),
                "module": (
                    definition.module
                ),
                "approved": False,
                "return_code": None,
                "marker_found": False,
                "elapsed": 0.0,
            }

        results.append(
            result
        )

        if not result[
            "approved"
        ]:
            separator(
                "REGRESION DETENIDA",
                char="!",
            )

            print(
                "Falló:",
                result[
                    "name"
                ],
            )

            print()

            print(
                "No ejecutaré las pruebas "
                "restantes para evitar mezclar "
                "un fallo con resultados posteriores."
            )

            break

    all_tests_executed = (
        len(
            results
        )
        == len(
            TESTS
        )
    )

    all_tests_approved = (
        all_tests_executed
        and all(
            bool(
                result[
                    "approved"
                ]
            )
            for result in results
        )
    )

    xslt_approved = False

    if all_tests_approved:
        separator(
            "VALIDACION FINAL ARTICLE.XSL"
        )

        try:
            validate_current_xslt()

            xslt_approved = True

            print()
            print(
                "ARTICLE.XSL: APROBADO"
            )

        except Exception as error:
            print(
                "ARTICLE.XSL: ERROR/REVISAR"
            )

            print(
                "ERROR:",
                repr(
                    error
                ),
            )

    elapsed_total = (
        time.monotonic()
        - total_start
    )

    separator(
        "RESUMEN FINAL — FASE 2027"
    )

    approved_count = sum(
        1
        for result in results
        if result[
            "approved"
        ]
    )

    for index, result in enumerate(
        results,
        start=1,
    ):
        print(
            f"{index}.",
            result[
                "name"
            ],
            "|",
            (
                "APROBADO"
                if result[
                    "approved"
                ]
                else "ERROR/REVISAR"
            ),
            "|",
            f"{result['elapsed']:.1f} s",
        )

    if len(
        results
    ) < len(
        TESTS
    ):
        for index in range(
            len(
                results
            ),
            len(
                TESTS
            ),
        ):
            print(
                f"{index + 1}.",
                TESTS[
                    index
                ].name,
                "| NO EJECUTADO",
            )

    print()

    print(
        "PRUEBAS DEFINIDAS:",
        len(
            TESTS
        ),
    )

    print(
        "PRUEBAS EJECUTADAS:",
        len(
            results
        ),
    )

    print(
        "PRUEBAS APROBADAS:",
        approved_count,
    )

    print(
        "ARTICLE.XSL:",
        (
            "APROBADO"
            if xslt_approved
            else "REVISAR/NO EJECUTADO"
        ),
    )

    print(
        "TIEMPO TOTAL:",
        f"{elapsed_total:.1f} s",
    )

    print()

    print(
        "FIGURAS:",
        (
            "APROBADAS"
            if (
                all_tests_approved
                and xslt_approved
            )
            else "REVISAR"
        ),
    )

    print(
        "TABLAS:",
        (
            "APROBADAS"
            if (
                all_tests_approved
                and xslt_approved
            )
            else "REVISAR"
        ),
    )

    print(
        "LISTAS:",
        (
            "APROBADAS"
            if (
                all_tests_approved
                and xslt_approved
            )
            else "REVISAR"
        ),
    )

    print(
        "ECUACIONES OMML/MATHML:",
        (
            "APROBADAS"
            if (
                all_tests_approved
                and xslt_approved
            )
            else "REVISAR"
        ),
    )

    print(
        "1354 END-TO-END:",
        (
            "APROBADO"
            if (
                all_tests_approved
                and xslt_approved
            )
            else "REVISAR"
        ),
    )

    print(
        "1293 REGRESION ESTRUCTURAL:",
        (
            "APROBADA"
            if (
                all_tests_approved
                and xslt_approved
            )
            else "REVISAR"
        ),
    )

    print(
        "PAGINACION:",
        "FUERA DE QA",
    )

    print()

    final_approved = (
        all_tests_approved
        and xslt_approved
    )

    if final_approved:
        print(
            "REGRESION FINAL FASE 2027: "
            "APROBADA"
        )

    else:
        print(
            "REGRESION FINAL FASE 2027: "
            "REVISAR"
        )

        raise SystemExit(
            1
        )

    print(
        "=" * 104
    )


if __name__ == "__main__":
    main()