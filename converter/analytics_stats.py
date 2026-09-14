# -*- coding: utf-8 -*-
from __future__ import annotations

import math
import statistics

from collections import Counter
from typing import Any, Iterable, Mapping, Sequence


Number = int | float


# ============================================================
# LIMPIEZA NUMERICA
# ============================================================

def finite_float(
    value: Any,
) -> float | None:
    """
    Convierte un valor a float únicamente si es finito.

    None, bool, NaN, +/-inf y valores no numéricos
    se consideran ausentes.
    """

    if value is None:
        return None

    if isinstance(
        value,
        bool,
    ):
        return None

    try:
        result = float(
            value
        )

    except (
        TypeError,
        ValueError,
    ):
        return None

    if not math.isfinite(
        result
    ):
        return None

    return result


def numeric_values(
    values: Iterable[Any],
) -> list[float]:
    """
    Devuelve solo valores numéricos finitos.
    """

    output: list[float] = []

    for value in values:

        number = finite_float(
            value
        )

        if number is not None:
            output.append(
                number
            )

    return output


def paired_numeric_values(
    xs: Iterable[Any],
    ys: Iterable[Any],
) -> tuple[
    list[float],
    list[float],
]:
    """
    Conserva únicamente pares donde X e Y son válidos.
    """

    clean_x: list[float] = []
    clean_y: list[float] = []

    for x_value, y_value in zip(
        xs,
        ys,
    ):

        x_number = finite_float(
            x_value
        )

        y_number = finite_float(
            y_value
        )

        if (
            x_number is None
            or y_number is None
        ):
            continue

        clean_x.append(
            x_number
        )

        clean_y.append(
            y_number
        )

    return (
        clean_x,
        clean_y,
    )


# ============================================================
# PERCENTILES
# ============================================================

def percentile(
    values: Iterable[Any],
    probability: float,
) -> float | None:
    """
    Percentil con interpolación lineal.

    probability debe estar entre 0 y 1.
    """

    if not (
        0.0
        <= probability
        <= 1.0
    ):
        raise ValueError(
            "probability debe estar "
            "entre 0 y 1."
        )

    data = sorted(
        numeric_values(
            values
        )
    )

    count = len(
        data
    )

    if count == 0:
        return None

    if count == 1:
        return data[0]

    position = (
        (count - 1)
        * probability
    )

    lower = math.floor(
        position
    )

    upper = math.ceil(
        position
    )

    if lower == upper:
        return data[
            lower
        ]

    weight = (
        position
        - lower
    )

    return (
        data[lower]
        * (1.0 - weight)
        + data[upper]
        * weight
    )


def percentile_rank(
    value: Any,
    reference: Iterable[Any],
) -> float | None:
    """
    Percentile rank con tratamiento de empates mediante
    mid-rank.

    Ejemplo:
    50 significa aproximadamente centro de la distribución.
    """

    number = finite_float(
        value
    )

    data = numeric_values(
        reference
    )

    if (
        number is None
        or not data
    ):
        return None

    less = sum(
        1
        for item in data
        if item < number
    )

    equal = sum(
        1
        for item in data
        if item == number
    )

    return (
        100.0
        * (
            less
            + 0.5
            * equal
        )
        / len(
            data
        )
    )


# ============================================================
# ESTADISTICA DESCRIPTIVA
# ============================================================

def describe(
    values: Iterable[Any],
) -> dict[str, Any]:
    """
    Resumen estadístico completo.

    Reporta desviación/varianza poblacional y muestral
    por separado para evitar ambigüedad.
    """

    raw = list(
        values
    )

    data = numeric_values(
        raw
    )

    count = len(
        data
    )

    missing = (
        len(
            raw
        )
        - count
    )

    if count == 0:

        return {
            "n": 0,
            "missing": missing,
            "minimum": None,
            "maximum": None,
            "range": None,
            "mean": None,
            "median": None,
            "variance_population": None,
            "std_population": None,
            "variance_sample": None,
            "std_sample": None,
            "cv_pct": None,
            "p25": None,
            "p50": None,
            "p75": None,
            "p90": None,
            "p95": None,
            "iqr": None,
            "lower_fence": None,
            "upper_fence": None,
        }

    minimum = min(
        data
    )

    maximum = max(
        data
    )

    mean = statistics.fmean(
        data
    )

    median = statistics.median(
        data
    )

    variance_population = (
        statistics.pvariance(
            data
        )
        if count >= 1
        else None
    )

    std_population = (
        statistics.pstdev(
            data
        )
        if count >= 1
        else None
    )

    variance_sample = (
        statistics.variance(
            data
        )
        if count >= 2
        else None
    )

    std_sample = (
        statistics.stdev(
            data
        )
        if count >= 2
        else None
    )

    p25 = percentile(
        data,
        0.25,
    )

    p50 = percentile(
        data,
        0.50,
    )

    p75 = percentile(
        data,
        0.75,
    )

    p90 = percentile(
        data,
        0.90,
    )

    p95 = percentile(
        data,
        0.95,
    )

    if (
        p25 is None
        or p75 is None
    ):
        iqr = None
        lower_fence = None
        upper_fence = None

    else:

        iqr = (
            p75
            - p25
        )

        lower_fence = (
            p25
            - 1.5
            * iqr
        )

        upper_fence = (
            p75
            + 1.5
            * iqr
        )

    cv_pct = None

    if (
        std_population is not None
        and mean != 0
    ):
        cv_pct = (
            abs(
                std_population
                / mean
            )
            * 100.0
        )

    return {
        "n": count,
        "missing": missing,
        "minimum": minimum,
        "maximum": maximum,
        "range": (
            maximum
            - minimum
        ),
        "mean": mean,
        "median": median,
        "variance_population": (
            variance_population
        ),
        "std_population": (
            std_population
        ),
        "variance_sample": (
            variance_sample
        ),
        "std_sample": (
            std_sample
        ),
        "cv_pct": cv_pct,
        "p25": p25,
        "p50": p50,
        "p75": p75,
        "p90": p90,
        "p95": p95,
        "iqr": iqr,
        "lower_fence": (
            lower_fence
        ),
        "upper_fence": (
            upper_fence
        ),
    }


# ============================================================
# Z-SCORE
# ============================================================

def z_scores(
    values: Sequence[Any],
) -> list[float | None]:
    """
    Z-score poblacional:

        z = (x - media) / desviación estándar

    Se usa la desviación estándar poblacional del conjunto
    observado porque el objetivo principal es detectar
    observaciones relativas dentro del dataset del Dashboard.
    """

    cleaned = numeric_values(
        values
    )

    if not cleaned:

        return [
            None
            for _ in values
        ]

    mean = statistics.fmean(
        cleaned
    )

    std = statistics.pstdev(
        cleaned
    )

    if std == 0:

        return [
            (
                0.0
                if finite_float(
                    value
                )
                is not None
                else None
            )
            for value in values
        ]

    result: list[
        float | None
    ] = []

    for value in values:

        number = finite_float(
            value
        )

        if number is None:
            result.append(
                None
            )

            continue

        result.append(
            (
                number
                - mean
            )
            / std
        )

    return result


def median_absolute_deviation(
    values: Iterable[Any],
) -> float | None:
    """
    MAD = mediana(|x - mediana(x)|)
    """

    data = numeric_values(
        values
    )

    if not data:
        return None

    med = statistics.median(
        data
    )

    deviations = [
        abs(
            value
            - med
        )
        for value in data
    ]

    return statistics.median(
        deviations
    )


def robust_z_scores(
    values: Sequence[Any],
) -> list[float | None]:
    """
    Modified/robust z-score:

        0.67448975 * (x - mediana) / MAD

    Es menos sensible a outliers que el z-score clásico.
    """

    data = numeric_values(
        values
    )

    if not data:

        return [
            None
            for _ in values
        ]

    med = statistics.median(
        data
    )

    mad = median_absolute_deviation(
        data
    )

    if (
        mad is None
        or mad == 0
    ):

        return [
            (
                0.0
                if (
                    finite_float(
                        value
                    )
                    == med
                )
                else None
            )
            for value in values
        ]

    scale = 0.67448975

    result: list[
        float | None
    ] = []

    for value in values:

        number = finite_float(
            value
        )

        if number is None:

            result.append(
                None
            )

            continue

        result.append(
            scale
            * (
                number
                - med
            )
            / mad
        )

    return result


# ============================================================
# OUTLIERS
# ============================================================

def iqr_outlier_flags(
    values: Sequence[Any],
    multiplier: float = 1.5,
) -> list[bool | None]:
    """
    Detecta outliers por regla de Tukey:

        x < Q1 - k*IQR
        x > Q3 + k*IQR
    """

    data = numeric_values(
        values
    )

    if not data:

        return [
            None
            for _ in values
        ]

    q1 = percentile(
        data,
        0.25,
    )

    q3 = percentile(
        data,
        0.75,
    )

    if (
        q1 is None
        or q3 is None
    ):

        return [
            None
            for _ in values
        ]

    iqr = (
        q3
        - q1
    )

    lower = (
        q1
        - multiplier
        * iqr
    )

    upper = (
        q3
        + multiplier
        * iqr
    )

    output: list[
        bool | None
    ] = []

    for value in values:

        number = finite_float(
            value
        )

        if number is None:

            output.append(
                None
            )

        else:

            output.append(
                bool(
                    number < lower
                    or number > upper
                )
            )

    return output


def z_outlier_flags(
    values: Sequence[Any],
    threshold: float = 2.0,
) -> list[bool | None]:
    """
    Marca observaciones con |z| >= threshold.
    """

    scores = z_scores(
        values
    )

    return [
        (
            abs(
                score
            )
            >= threshold
            if score is not None
            else None
        )
        for score in scores
    ]


def robust_z_outlier_flags(
    values: Sequence[Any],
    threshold: float = 3.5,
) -> list[bool | None]:
    """
    Umbral 3.5 para modified z-score.
    """

    scores = robust_z_scores(
        values
    )

    return [
        (
            abs(
                score
            )
            >= threshold
            if score is not None
            else None
        )
        for score in scores
    ]


# ============================================================
# RANKING
# ============================================================

def rank_rows(
    rows: Sequence[Mapping[str, Any]],
    metric: str,
    *,
    label_key: str = "article_code",
    top_n: int | None = 10,
    descending: bool = True,
) -> list[dict[str, Any]]:
    """
    Ranking genérico por cualquier métrica numérica.
    """

    candidates: list[
        dict[str, Any]
    ] = []

    reference_values = numeric_values(
        row.get(
            metric
        )
        for row in rows
    )

    for row in rows:

        value = finite_float(
            row.get(
                metric
            )
        )

        if value is None:
            continue

        candidates.append(
            {
                "label": (
                    str(
                        row.get(
                            label_key,
                            ""
                        )
                        or ""
                    )
                ),
                "value": value,
                "percentile": (
                    percentile_rank(
                        value,
                        reference_values,
                    )
                ),
                "source": dict(
                    row
                ),
            }
        )

    candidates.sort(
        key=lambda item: item[
            "value"
        ],
        reverse=descending,
    )

    if top_n is not None:

        candidates = candidates[
            : max(
                0,
                int(
                    top_n
                ),
            )
        ]

    for index, item in enumerate(
        candidates,
        start=1,
    ):

        item[
            "rank"
        ] = index

    return candidates


# ============================================================
# RANKING ESTADISTICO CON Z-SCORE
# ============================================================

def enrich_rows_with_zscore(
    rows: Sequence[Mapping[str, Any]],
    metric: str,
) -> list[dict[str, Any]]:
    """
    Añade z-score, robust z-score y outliers IQR
    preservando el orden original.
    """

    values = [
        row.get(
            metric
        )
        for row in rows
    ]

    classic = z_scores(
        values
    )

    robust = robust_z_scores(
        values
    )

    iqr_flags = iqr_outlier_flags(
        values
    )

    output: list[
        dict[str, Any]
    ] = []

    for (
        row,
        z_value,
        robust_value,
        iqr_flag,
    ) in zip(
        rows,
        classic,
        robust,
        iqr_flags,
    ):

        item = dict(
            row
        )

        item[
            f"{metric}_zscore"
        ] = z_value

        item[
            f"{metric}_robust_zscore"
        ] = robust_value

        item[
            f"{metric}_iqr_outlier"
        ] = iqr_flag

        output.append(
            item
        )

    return output


# ============================================================
# RANKDATA CON EMPATES
# ============================================================

def average_ranks(
    values: Sequence[float],
) -> list[float]:
    """
    Ranks ascendentes 1..n con promedio para empates.
    """

    indexed = sorted(
        enumerate(
            values
        ),
        key=lambda item: item[1],
    )

    ranks = [
        0.0
        for _ in values
    ]

    index = 0

    while index < len(
        indexed
    ):

        end = index + 1

        while (
            end
            < len(
                indexed
            )
            and indexed[end][1]
            == indexed[index][1]
        ):
            end += 1

        first_rank = (
            index
            + 1
        )

        last_rank = end

        average = (
            first_rank
            + last_rank
        ) / 2.0

        for position in range(
            index,
            end,
        ):

            original_index = (
                indexed[
                    position
                ][0]
            )

            ranks[
                original_index
            ] = average

        index = end

    return ranks


# ============================================================
# CORRELACION
# ============================================================

def pearson_correlation(
    xs: Iterable[Any],
    ys: Iterable[Any],
) -> dict[str, Any]:
    """
    Correlación de Pearson.

    No calcula p-value deliberadamente: la capa visual debe
    presentar la relación como exploratoria mientras la muestra
    sea pequeña y antes de validar supuestos inferenciales.
    """

    x_values, y_values = (
        paired_numeric_values(
            xs,
            ys,
        )
    )

    count = len(
        x_values
    )

    if count < 2:

        return {
            "n": count,
            "r": None,
        }

    mean_x = statistics.fmean(
        x_values
    )

    mean_y = statistics.fmean(
        y_values
    )

    centered_x = [
        value
        - mean_x
        for value in x_values
    ]

    centered_y = [
        value
        - mean_y
        for value in y_values
    ]

    denominator_x = math.sqrt(
        sum(
            value ** 2
            for value
            in centered_x
        )
    )

    denominator_y = math.sqrt(
        sum(
            value ** 2
            for value
            in centered_y
        )
    )

    denominator = (
        denominator_x
        * denominator_y
    )

    if denominator == 0:

        return {
            "n": count,
            "r": None,
        }

    numerator = sum(
        x_value
        * y_value
        for x_value, y_value
        in zip(
            centered_x,
            centered_y,
        )
    )

    correlation = (
        numerator
        / denominator
    )

    correlation = max(
        -1.0,
        min(
            1.0,
            correlation,
        ),
    )

    return {
        "n": count,
        "r": correlation,
    }


def spearman_correlation(
    xs: Iterable[Any],
    ys: Iterable[Any],
) -> dict[str, Any]:
    """
    Spearman = Pearson sobre los rangos.
    """

    x_values, y_values = (
        paired_numeric_values(
            xs,
            ys,
        )
    )

    count = len(
        x_values
    )

    if count < 2:

        return {
            "n": count,
            "rho": None,
        }

    x_ranks = average_ranks(
        x_values
    )

    y_ranks = average_ranks(
        y_values
    )

    result = pearson_correlation(
        x_ranks,
        y_ranks,
    )

    return {
        "n": result[
            "n"
        ],
        "rho": result[
            "r"
        ],
    }


# ============================================================
# REGRESION LINEAL EXPLORATORIA
# ============================================================

def linear_regression(
    xs: Iterable[Any],
    ys: Iterable[Any],
) -> dict[str, Any]:
    """
    Regresión lineal simple:

        y = intercept + slope*x

    Devuelve R² pero no inferencia causal ni p-value.
    """

    x_values, y_values = (
        paired_numeric_values(
            xs,
            ys,
        )
    )

    count = len(
        x_values
    )

    if count < 2:

        return {
            "n": count,
            "slope": None,
            "intercept": None,
            "r": None,
            "r_squared": None,
        }

    mean_x = statistics.fmean(
        x_values
    )

    mean_y = statistics.fmean(
        y_values
    )

    ss_x = sum(
        (
            value
            - mean_x
        ) ** 2
        for value
        in x_values
    )

    if ss_x == 0:

        return {
            "n": count,
            "slope": None,
            "intercept": None,
            "r": None,
            "r_squared": None,
        }

    covariance_sum = sum(
        (
            x_value
            - mean_x
        )
        * (
            y_value
            - mean_y
        )
        for x_value, y_value
        in zip(
            x_values,
            y_values,
        )
    )

    slope = (
        covariance_sum
        / ss_x
    )

    intercept = (
        mean_y
        - slope
        * mean_x
    )

    correlation = pearson_correlation(
        x_values,
        y_values,
    )

    r = correlation[
        "r"
    ]

    return {
        "n": count,
        "slope": slope,
        "intercept": intercept,
        "r": r,
        "r_squared": (
            r ** 2
            if r is not None
            else None
        ),
    }


# ============================================================
# DATASET PARA SCATTER
# ============================================================

def scatter_dataset(
    rows: Sequence[Mapping[str, Any]],
    x_key: str,
    y_key: str,
    *,
    label_key: str = "article_code",
) -> list[dict[str, Any]]:
    """
    Prepara puntos de dispersión, preservando etiqueta.
    """

    output: list[
        dict[str, Any]
    ] = []

    for row in rows:

        x_value = finite_float(
            row.get(
                x_key
            )
        )

        y_value = finite_float(
            row.get(
                y_key
            )
        )

        if (
            x_value is None
            or y_value is None
        ):
            continue

        output.append(
            {
                "label": str(
                    row.get(
                        label_key,
                        "",
                    )
                    or ""
                ),
                "x": x_value,
                "y": y_value,
                "source": dict(
                    row
                ),
            }
        )

    return output


# ============================================================
# BOXPLOT
# ============================================================

def boxplot_summary(
    values: Iterable[Any],
) -> dict[str, Any]:
    """
    Valores necesarios para un boxplot.
    """

    data = numeric_values(
        values
    )

    summary = describe(
        data
    )

    if not data:

        return {
            **summary,
            "whisker_low": None,
            "whisker_high": None,
            "outliers": [],
        }

    lower_fence = summary[
        "lower_fence"
    ]

    upper_fence = summary[
        "upper_fence"
    ]

    inliers = [
        value
        for value in data
        if (
            lower_fence is None
            or value
            >= lower_fence
        )
        and (
            upper_fence is None
            or value
            <= upper_fence
        )
    ]

    outliers = [
        value
        for value in data
        if (
            lower_fence is not None
            and value
            < lower_fence
        )
        or (
            upper_fence is not None
            and value
            > upper_fence
        )
    ]

    return {
        **summary,
        "whisker_low": (
            min(
                inliers
            )
            if inliers
            else None
        ),
        "whisker_high": (
            max(
                inliers
            )
            if inliers
            else None
        ),
        "outliers": sorted(
            outliers
        ),
    }


# ============================================================
# HISTOGRAMA
# ============================================================

def histogram(
    values: Iterable[Any],
    bins: int | None = None,
) -> dict[str, Any]:
    """
    Histograma sin depender de NumPy.

    Si bins es None utiliza una aproximación de
    Sturges para elegir el número de intervalos.
    """

    data = numeric_values(
        values
    )

    count = len(
        data
    )

    if count == 0:

        return {
            "n": 0,
            "edges": [],
            "counts": [],
        }

    minimum = min(
        data
    )

    maximum = max(
        data
    )

    if bins is None:

        bins = max(
            1,
            math.ceil(
                math.log2(
                    count
                )
                + 1
            ),
        )

    bins = max(
        1,
        int(
            bins
        ),
    )

    if minimum == maximum:

        return {
            "n": count,
            "edges": [
                minimum,
                maximum,
            ],
            "counts": [
                count,
            ],
        }

    width = (
        maximum
        - minimum
    ) / bins

    edges = [
        minimum
        + index
        * width
        for index in range(
            bins + 1
        )
    ]

    counts = [
        0
        for _ in range(
            bins
        )
    ]

    for value in data:

        if value == maximum:

            bucket = (
                bins
                - 1
            )

        else:

            bucket = int(
                (
                    value
                    - minimum
                )
                / width
            )

            bucket = max(
                0,
                min(
                    bins - 1,
                    bucket,
                ),
            )

        counts[
            bucket
        ] += 1

    return {
        "n": count,
        "edges": edges,
        "counts": counts,
    }


# ============================================================
# CURVA NORMAL TEORICA
# ============================================================

def normal_curve(
    values: Iterable[Any],
    *,
    points: int = 81,
    sigma_span: float = 3.0,
) -> dict[str, Any]:
    """
    Genera puntos de una Normal(mu, sigma) para superponer
    a un histograma.

    Esto NO demuestra normalidad de los datos.
    Solo produce una curva teórica de comparación.
    """

    data = numeric_values(
        values
    )

    count = len(
        data
    )

    if count < 2:

        return {
            "n": count,
            "mean": None,
            "std": None,
            "x": [],
            "density": [],
        }

    mean = statistics.fmean(
        data
    )

    std = statistics.pstdev(
        data
    )

    if std == 0:

        return {
            "n": count,
            "mean": mean,
            "std": 0.0,
            "x": [],
            "density": [],
        }

    points = max(
        3,
        int(
            points
        ),
    )

    start = (
        mean
        - sigma_span
        * std
    )

    end = (
        mean
        + sigma_span
        * std
    )

    step = (
        end
        - start
    ) / (
        points
        - 1
    )

    xs = [
        start
        + index
        * step
        for index in range(
            points
        )
    ]

    coefficient = (
        1.0
        / (
            std
            * math.sqrt(
                2.0
                * math.pi
            )
        )
    )

    density = [
        coefficient
        * math.exp(
            -0.5
            * (
                (
                    x_value
                    - mean
                )
                / std
            ) ** 2
        )
        for x_value in xs
    ]

    return {
        "n": count,
        "mean": mean,
        "std": std,
        "x": xs,
        "density": density,
    }


# ============================================================
# LIMITES DE CONTROL / SIGMA
# ============================================================

def sigma_control_limits(
    values: Iterable[Any],
    sigma: float = 3.0,
) -> dict[str, Any]:
    """
    Límites descriptivos:

        media +/- sigma * desviación estándar

    No sustituye un control estadístico de procesos formal.
    """

    data = numeric_values(
        values
    )

    if not data:

        return {
            "n": 0,
            "mean": None,
            "std": None,
            "lower": None,
            "upper": None,
            "sigma": sigma,
        }

    mean = statistics.fmean(
        data
    )

    std = statistics.pstdev(
        data
    )

    return {
        "n": len(
            data
        ),
        "mean": mean,
        "std": std,
        "lower": (
            mean
            - sigma
            * std
        ),
        "upper": (
            mean
            + sigma
            * std
        ),
        "sigma": sigma,
    }


# ============================================================
# MEDIA MOVIL
# ============================================================

def moving_average(
    values: Sequence[Any],
    window: int = 3,
) -> list[float | None]:
    """
    Media móvil simple.

    Hasta completar la ventana devuelve None.
    """

    window = int(
        window
    )

    if window <= 0:

        raise ValueError(
            "window debe ser positivo."
        )

    output: list[
        float | None
    ] = []

    clean_buffer: list[
        float
    ] = []

    for value in values:

        number = finite_float(
            value
        )

        clean_buffer.append(
            number
            if number is not None
            else math.nan
        )

        if len(
            clean_buffer
        ) < window:

            output.append(
                None
            )

            continue

        current = clean_buffer[
            -window:
        ]

        if any(
            not math.isfinite(
                item
            )
            for item in current
        ):

            output.append(
                None
            )

            continue

        output.append(
            statistics.fmean(
                current
            )
        )

    return output


# ============================================================
# HEATMAP
# ============================================================

def heatmap_matrix(
    rows: Sequence[Mapping[str, Any]],
    value_keys: Sequence[str],
    *,
    label_key: str = "article_code",
    normalization: str = "zscore",
) -> dict[str, Any]:
    """
    Prepara una matriz normalizada para heatmap.

    normalization:
        zscore
        minmax
        none
    """

    normalization = (
        str(
            normalization
        )
        .strip()
        .casefold()
    )

    if normalization not in {
        "zscore",
        "minmax",
        "none",
    }:

        raise ValueError(
            "normalization debe ser "
            "zscore, minmax o none."
        )

    columns = list(
        value_keys
    )

    labels = [
        str(
            row.get(
                label_key,
                "",
            )
            or ""
        )
        for row in rows
    ]

    raw_columns: dict[
        str,
        list[Any],
    ] = {
        key: [
            row.get(
                key
            )
            for row in rows
        ]
        for key in columns
    }

    normalized_columns: dict[
        str,
        list[float | None],
    ] = {}

    for key in columns:

        values = raw_columns[
            key
        ]

        if normalization == "zscore":

            normalized_columns[
                key
            ] = z_scores(
                values
            )

        elif normalization == "minmax":

            valid = numeric_values(
                values
            )

            if not valid:

                normalized_columns[
                    key
                ] = [
                    None
                    for _ in values
                ]

                continue

            minimum = min(
                valid
            )

            maximum = max(
                valid
            )

            span = (
                maximum
                - minimum
            )

            column_values: list[
                float | None
            ] = []

            for value in values:

                number = finite_float(
                    value
                )

                if number is None:

                    column_values.append(
                        None
                    )

                elif span == 0:

                    column_values.append(
                        0.5
                    )

                else:

                    column_values.append(
                        (
                            number
                            - minimum
                        )
                        / span
                    )

            normalized_columns[
                key
            ] = column_values

        else:

            normalized_columns[
                key
            ] = [
                finite_float(
                    value
                )
                for value in values
            ]

    matrix: list[
        list[float | None]
    ] = []

    raw_matrix: list[
        list[float | None]
    ] = []

    for row_index in range(
        len(
            rows
        )
    ):

        matrix.append(
            [
                normalized_columns[
                    key
                ][
                    row_index
                ]
                for key in columns
            ]
        )

        raw_matrix.append(
            [
                finite_float(
                    raw_columns[
                        key
                    ][
                        row_index
                    ]
                )
                for key in columns
            ]
        )

    return {
        "labels": labels,
        "columns": columns,
        "normalization": (
            normalization
        ),
        "matrix": matrix,
        "raw_matrix": raw_matrix,
    }


# ============================================================
# MATRIZ DE CORRELACION
# ============================================================

def correlation_matrix(
    rows: Sequence[Mapping[str, Any]],
    fields: Sequence[str],
    *,
    method: str = "pearson",
) -> dict[str, Any]:
    """
    Matriz de correlaciones pairwise.

    Cada celda también conserva n porque con valores NULL
    el tamaño muestral puede cambiar entre pares.
    """

    method = (
        str(
            method
        )
        .strip()
        .casefold()
    )

    if method not in {
        "pearson",
        "spearman",
    }:

        raise ValueError(
            "method debe ser pearson "
            "o spearman."
        )

    fields = list(
        fields
    )

    matrix: list[
        list[float | None]
    ] = []

    sample_sizes: list[
        list[int]
    ] = []

    for field_x in fields:

        correlation_row: list[
            float | None
        ] = []

        n_row: list[int] = []

        for field_y in fields:

            xs = [
                row.get(
                    field_x
                )
                for row in rows
            ]

            ys = [
                row.get(
                    field_y
                )
                for row in rows
            ]

            if method == "pearson":

                result = (
                    pearson_correlation(
                        xs,
                        ys,
                    )
                )

                value = result[
                    "r"
                ]

            else:

                result = (
                    spearman_correlation(
                        xs,
                        ys,
                    )
                )

                value = result[
                    "rho"
                ]

            correlation_row.append(
                value
            )

            n_row.append(
                int(
                    result[
                        "n"
                    ]
                )
            )

        matrix.append(
            correlation_row
        )

        sample_sizes.append(
            n_row
        )

    return {
        "fields": fields,
        "method": method,
        "matrix": matrix,
        "sample_sizes": (
            sample_sizes
        ),
    }


# ============================================================
# WATERFALL
# ============================================================

def waterfall_dataset(
    components: Sequence[
        tuple[
            str,
            Any,
        ]
    ],
    *,
    start: float = 0.0,
) -> list[dict[str, Any]]:
    """
    Prepara una cascada acumulativa genérica.

    Cada componente se interpreta como incremento/decremento.
    """

    current = float(
        start
    )

    output: list[
        dict[str, Any]
    ] = []

    for label, raw_value in components:

        value = finite_float(
            raw_value
        )

        if value is None:
            continue

        previous = current

        current = (
            previous
            + value
        )

        output.append(
            {
                "label": str(
                    label
                ),
                "value": value,
                "start": previous,
                "end": current,
                "direction": (
                    "increase"
                    if value > 0
                    else (
                        "decrease"
                        if value < 0
                        else "neutral"
                    )
                ),
            }
        )

    return output


# ============================================================
# PARETO
# ============================================================

def pareto_dataset(
    labels: Iterable[Any],
) -> list[dict[str, Any]]:
    """
    Frecuencia + porcentaje acumulado para gráfico de Pareto.
    """

    cleaned = [
        str(
            value
        ).strip()
        for value in labels
        if str(
            value
            or ""
        ).strip()
    ]

    counts = Counter(
        cleaned
    )

    ordered = sorted(
        counts.items(),
        key=lambda item: (
            -item[1],
            item[0],
        ),
    )

    total = sum(
        count
        for _label, count
        in ordered
    )

    cumulative = 0

    output: list[
        dict[str, Any]
    ] = []

    for label, count in ordered:

        cumulative += count

        output.append(
            {
                "label": label,
                "count": count,
                "share_pct": (
                    100.0
                    * count
                    / total
                    if total
                    else 0.0
                ),
                "cumulative_pct": (
                    100.0
                    * cumulative
                    / total
                    if total
                    else 0.0
                ),
            }
        )

    return output


# ============================================================
# CONTEXTO DE MUESTRA
# ============================================================

def sample_context(
    n: int,
) -> dict[str, Any]:
    """
    Mensaje conservador para el Dashboard.

    Los umbrales son heurísticos de comunicación,
    no una prueba estadística.
    """

    n = max(
        0,
        int(
            n
        ),
    )

    if n == 0:

        level = "sin_datos"

        message = (
            "Sin observaciones disponibles."
        )

    elif n < 10:

        level = "muy_pequena"

        message = (
            "Muestra muy pequeña. "
            "Interpretar únicamente como "
            "análisis exploratorio."
        )

    elif n < 30:

        level = "limitada"

        message = (
            "Muestra limitada. "
            "Las relaciones observadas deben "
            "interpretarse con cautela."
        )

    else:

        level = "ampliada"

        message = (
            "Muestra más amplia. "
            "La inferencia formal todavía requiere "
            "validar supuestos estadísticos."
        )

    return {
        "n": n,
        "level": level,
        "message": message,
    }
