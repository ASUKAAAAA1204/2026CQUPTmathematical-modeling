#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import math
import random
import statistics as st
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path


RELATIVE_WIDTH_THRESHOLD = 1.0
MIN_RETENTION_RATIO = 0.05
SUPPORT_PROB_ACCEPT = 0.975
SUPPORT_PROB_REJECT = 0.025
CALIBRATION_FLOOR_RATIO = 0.40
BOOTSTRAP_REPS = 1000


def find_base_dir() -> Path:
    return Path(__file__).resolve().parent.parent


def read_csv_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def parse_date(text: str) -> datetime:
    return datetime.strptime(text.strip(), "%Y-%m-%d")


def to_float(text: str) -> float:
    return float(str(text).strip())


def mean_sd(values: list[float]) -> tuple[float, float]:
    mean = sum(values) / len(values)
    sd = (sum((x - mean) ** 2 for x in values) / len(values)) ** 0.5
    return mean, sd


def invert_matrix(matrix: list[list[float]]) -> list[list[float]]:
    n = len(matrix)
    a = [row[:] + [1.0 if i == j else 0.0 for j in range(n)] for i, row in enumerate(matrix)]
    for col in range(n):
        pivot = max(range(col, n), key=lambda r: abs(a[r][col]))
        if abs(a[pivot][col]) < 1e-12:
            raise ValueError("Singular matrix")
        if pivot != col:
            a[col], a[pivot] = a[pivot], a[col]
        pv = a[col][col]
        for j in range(2 * n):
            a[col][j] /= pv
        for r in range(n):
            if r == col:
                continue
            factor = a[r][col]
            if factor:
                for j in range(2 * n):
                    a[r][j] -= factor * a[col][j]
    return [[a[i][j + n] for j in range(n)] for i in range(n)]


def fit_ols(x_rows: list[list[float]], y: list[float]) -> dict[str, object]:
    n = len(x_rows)
    p = len(x_rows[0])
    xtx = [[0.0] * p for _ in range(p)]
    xty = [0.0] * p
    for x, yi in zip(x_rows, y):
        for i in range(p):
            xty[i] += x[i] * yi
            for j in range(p):
                xtx[i][j] += x[i] * x[j]
    inv = invert_matrix(xtx)
    beta = [sum(inv[i][j] * xty[j] for j in range(p)) for i in range(p)]
    yhat = [sum(beta[j] * row[j] for j in range(p)) for row in x_rows]
    resid = [yi - yh for yi, yh in zip(y, yhat)]
    sse = sum(e * e for e in resid)
    ybar = sum(y) / n
    sst = sum((yi - ybar) ** 2 for yi in y)
    sigma2 = sse / max(1, n - p)
    return {
        "beta": beta,
        "resid": resid,
        "rmse": (sse / n) ** 0.5,
        "r2": 1.0 - sse / sst if sst > 0 else float("nan"),
        "sigma": sigma2**0.5,
    }


def percentile(values: list[float], q: float) -> float:
    if not values:
        return float("nan")
    values = sorted(values)
    if len(values) == 1:
        return values[0]
    pos = (len(values) - 1) * q
    lo = int(math.floor(pos))
    hi = int(math.ceil(pos))
    if lo == hi:
        return values[lo]
    return values[lo] * (hi - pos) + values[hi] * (pos - lo)


def normal_cdf(z: float) -> float:
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def read_experiment(csv_path: Path):
    rows = []
    types_order = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        next(reader)
        for r in reader:
            sample_type = r[0]
            if sample_type not in types_order:
                types_order.append(sample_type)
            rows.append(
                {
                    "type": sample_type,
                    "id": int(r[1]),
                    "day": int(r[2]),
                    "mag": float(r[3]),
                    "temp": float(r[4]),
                    "hum": float(r[5]),
                }
            )
    return rows, types_order


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_xlsx_weather(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = workbook.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
        first_sheet = sheets[0]
        sheet_id = first_sheet.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]
        target = relmap[sheet_id].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target

        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            ns_t = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
            for si in shared_xml:
                shared.append("".join(t.text or "" for t in si.iter(ns_t)))

        rows = []
        worksheet = ET.fromstring(zf.read(target))
        ns_c = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
        ns_v = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v"
        ns_is = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is"
        ns_t = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
        for row in worksheet.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
            values = []
            for cell in row.iter(ns_c):
                ref = cell.attrib.get("r", "")
                idx = col_to_idx("".join(filter(str.isalpha, ref))) if ref else len(values)
                cell_type = cell.attrib.get("t")
                v = cell.find(ns_v)
                inline = cell.find(ns_is)
                value = ""
                if v is not None:
                    raw = v.text or ""
                    value = shared[int(raw)] if cell_type == "s" else raw
                elif inline is not None:
                    value = "".join(t.text or "" for t in inline.iter(ns_t))
                while len(values) <= idx:
                    values.append("")
                values[idx] = value
            if values and values[0].isdigit():
                rows.append(
                    {
                        "seq": int(values[0]),
                        "year": int(values[1]),
                        "month": int(values[2]),
                        "date": int(values[3]),
                        "weather": values[4],
                        "temp": float(values[5]),
                        "hum": float(values[6]),
                    }
                )
        return rows


def theta_static(sample_type: str) -> float:
    if "铁钉" in sample_type or "铁夹" in sample_type:
        return 1.0
    return 1.5


def build_training_weather(weather_rows: list[dict[str, object]]) -> dict[int, dict[str, float | str]]:
    return {
        int(row["seq"]): {
            "temp": float(row["temp"]),
            "hum": float(row["hum"]),
            "weather": str(row["weather"]),
        }
        for row in weather_rows
    }


def standardize_stats(training_weather: dict[int, dict[str, float]]) -> dict[str, float | list[float]]:
    temps = [training_weather[d]["temp"] for d in range(1, 91)]
    hums = [training_weather[d]["hum"] for d in range(1, 91)]
    temp_mean = sum(temps) / len(temps)
    hum_mean = sum(hums) / len(hums)

    cum_temp, cum_hum, cum_inter = [], [], []
    a = b = c = 0.0
    for d in range(1, 91):
        tc = training_weather[d]["temp"] - temp_mean
        hc = training_weather[d]["hum"] - hum_mean
        a += tc
        b += hc
        c += tc * hc
        cum_temp.append(a)
        cum_hum.append(b)
        cum_inter.append(c)

    day_mean, day_sd = mean_sd(list(range(1, 91)))
    ct_mean, ct_sd = mean_sd(cum_temp)
    ch_mean, ch_sd = mean_sd(cum_hum)
    cth_mean, cth_sd = mean_sd(cum_inter)
    return {
        "temp_mean": temp_mean,
        "hum_mean": hum_mean,
        "day_mean": day_mean,
        "day_sd": day_sd,
        "ct_mean": ct_mean,
        "ct_sd": ct_sd,
        "ch_mean": ch_mean,
        "ch_sd": ch_sd,
        "cth_mean": cth_mean,
        "cth_sd": cth_sd,
    }


def fit_main_model(observations: list[dict[str, object]], types_order: list[str], weather_rows: list[dict[str, object]]):
    weather = build_training_weather(weather_rows)
    stats = standardize_stats(weather)

    nail, clip, ordinary, rusty = types_order
    m0 = {}
    for row in observations:
        if row["day"] == 0:
            m0[(row["type"], row["id"])] = row["mag"]

    def features_for_path(sample_type: str, day: int, cum_temp_val: float, cum_hum_val: float, cum_inter_val: float):
        dz = (day - stats["day_mean"]) / stats["day_sd"]
        ct = (cum_temp_val - stats["ct_mean"]) / stats["ct_sd"]
        ch = (cum_hum_val - stats["ch_mean"]) / stats["ch_sd"]
        cth = (cum_inter_val - stats["cth_mean"]) / stats["cth_sd"]
        n = 1.0 if sample_type == nail else 0.0
        cl = 1.0 if sample_type == clip else 0.0
        r = 1.0 if sample_type == rusty else 0.0
        return [1.0, n, cl, r, dz, dz * dz, ct, ch, cth, r * dz, n * dz, cl * dz]

    x_rows, y = [], []
    for row in observations:
        if row["day"] == 0:
            continue
        day = int(row["day"])
        cum_temp_val = 0.0
        cum_hum_val = 0.0
        cum_inter_val = 0.0
        for d in range(1, day + 1):
            tc = weather[d]["temp"] - stats["temp_mean"]
            hc = weather[d]["hum"] - stats["hum_mean"]
            cum_temp_val += tc
            cum_hum_val += hc
            cum_inter_val += tc * hc
        x_rows.append(features_for_path(row["type"], day, cum_temp_val, cum_hum_val, cum_inter_val))
        y.append(math.log(m0[(row["type"], row["id"])] / row["mag"]))

    model = fit_ols(x_rows, y)
    cluster_rows = {}
    for row in observations:
        if row["day"] == 0:
            continue
        cluster_rows.setdefault((row["type"], row["id"]), []).append(row)

    return {
        "model": model,
        "m0": m0,
        "types_order": types_order,
        "weather": weather,
        "stats": stats,
        "cluster_rows": cluster_rows,
        "features_for_path": features_for_path,
    }


def normalize_weather_series(case_info: dict[str, str], weather_rows: list[dict[str, str]]) -> list[dict[str, object]]:
    lightning_date = parse_date(case_info["lightning_date"])
    measurement_date = parse_date(case_info["measurement_date"])
    delay_days = (measurement_date - lightning_date).days
    if delay_days <= 0:
        raise ValueError("measurement_date must be later than lightning_date")

    parsed = []
    for row in weather_rows:
        date_text = row.get("date", "").strip()
        if date_text:
            date_value = parse_date(date_text)
        else:
            date_value = lightning_date + timedelta(days=len(parsed) + 1)
        day_index_text = row.get("day_index", "").strip()
        day_index = int(day_index_text) if day_index_text else (date_value - lightning_date).days
        parsed.append(
            {
                "date": date_value,
                "day_index": day_index,
                "temp_c": to_float(row["temp_c"]),
                "hum_pct": to_float(row["hum_pct"]),
                "weather_text": row.get("weather_text", ""),
                "station_id": row.get("station_id", ""),
                "is_imputed": row.get("is_imputed", "false"),
            }
        )

    parsed = sorted(parsed, key=lambda x: x["day_index"])
    if not parsed:
        raise ValueError("weather_series is empty")

    if parsed[0]["day_index"] <= 0:
        for i, row in enumerate(parsed, start=1):
            row["day_index"] = i

    if len(parsed) < delay_days:
        last = parsed[-1]
        while len(parsed) < delay_days:
            new_day = len(parsed) + 1
            parsed.append(
                {
                    "date": lightning_date + timedelta(days=new_day),
                    "day_index": new_day,
                    "temp_c": last["temp_c"],
                    "hum_pct": last["hum_pct"],
                    "weather_text": last["weather_text"],
                    "station_id": last["station_id"],
                    "is_imputed": "true",
                }
            )

    dense = []
    by_day = {row["day_index"]: row for row in parsed}
    last_known = parsed[0]
    for day in range(1, delay_days + 1):
        if day in by_day:
            last_known = by_day[day]
            dense.append(last_known)
        else:
            dense.append(
                {
                    "date": lightning_date + timedelta(days=day),
                    "day_index": day,
                    "temp_c": last_known["temp_c"],
                    "hum_pct": last_known["hum_pct"],
                    "weather_text": last_known["weather_text"],
                    "station_id": last_known["station_id"],
                    "is_imputed": "true",
                }
            )
    return dense


def cumulative_path_features(weather_series: list[dict[str, object]], stats: dict[str, float]) -> list[dict[str, float | str]]:
    rows = []
    cum_temp = cum_hum = cum_inter = 0.0
    for day, row in enumerate(weather_series, start=1):
        tc = float(row["temp_c"]) - float(stats["temp_mean"])
        hc = float(row["hum_pct"]) - float(stats["hum_mean"])
        cum_temp += tc
        cum_hum += hc
        cum_inter += tc * hc
        rows.append(
            {
                "day": day,
                "weather": str(row["weather_text"]),
                "temp_c": float(row["temp_c"]),
                "hum_pct": float(row["hum_pct"]),
                "cum_temp": cum_temp,
                "cum_hum": cum_hum,
                "cum_inter": cum_inter,
            }
        )
    return rows


def predict_survival(beta: list[float], features_row: list[float]) -> float:
    return math.exp(-sum(b * x for b, x in zip(beta, features_row)))


def build_monotone_retention_series(
    beta: list[float],
    sample_type: str,
    path_rows: list[dict[str, float | str]],
    features_for_path,
) -> list[float]:
    series = []
    running_min = 1.0
    for path_row in path_rows:
        raw = predict_survival(
            beta,
            features_for_path(
                sample_type,
                int(path_row["day"]),
                float(path_row["cum_temp"]),
                float(path_row["cum_hum"]),
                float(path_row["cum_inter"]),
            ),
        )
        raw = min(max(raw, 0.0), 1.0)
        running_min = min(running_min, raw)
        series.append(running_min)
    return series


def build_case_outputs(case_info, sample_rows, weather_series, fitted):
    stats = fitted["stats"]
    model = fitted["model"]
    features_for_path = fitted["features_for_path"]
    types_order = fitted["types_order"]
    cluster_rows = fitted["cluster_rows"]
    cluster_keys = list(cluster_rows.keys())

    path_rows = cumulative_path_features(weather_series, stats)
    delay_days = len(path_rows)

    retention_curve_rows = []
    threshold_curve_rows = []
    retention_last = {}
    bootstrap_retention_map = {sample_type: [] for sample_type in types_order}

    for sample_type in types_order:
        theta0 = theta_static(sample_type)
        retention_series = build_monotone_retention_series(model["beta"], sample_type, path_rows, features_for_path)
        for path_row, retention in zip(path_rows, retention_series):
            retention_last[sample_type] = retention
            retention_curve_rows.append(
                [
                    case_info["case_id"],
                    sample_type,
                    path_row["day"],
                    path_row["weather"],
                    round(float(path_row["temp_c"]), 6),
                    round(float(path_row["hum_pct"]), 6),
                    round(retention, 6),
                ]
            )
            threshold_curve_rows.append(
                [
                    case_info["case_id"],
                    sample_type,
                    path_row["day"],
                    round(theta0, 6),
                    round(theta0 * retention, 6),
                ]
            )

    training_weather = fitted["weather"]
    for sample_type in types_order:
        bootstrap_retention_map[sample_type] = []
    rng = random.Random(20260514)
    for _ in range(BOOTSTRAP_REPS):
        sampled_keys = [rng.choice(cluster_keys) for _ in cluster_keys]
        boot_train_rows = []
        for key in sampled_keys:
            boot_train_rows.extend(cluster_rows[key])
        boot_x_rows = []
        boot_y = []
        for row in boot_train_rows:
            day = int(row["day"])
            cum_temp = cum_hum = cum_inter = 0.0
            for d in range(1, day + 1):
                tc = training_weather[d]["temp"] - float(stats["temp_mean"])
                hc = training_weather[d]["hum"] - float(stats["hum_mean"])
                cum_temp += tc
                cum_hum += hc
                cum_inter += tc * hc
            boot_x_rows.append(features_for_path(row["type"], day, cum_temp, cum_hum, cum_inter))
            boot_y.append(math.log(fitted["m0"][(row["type"], row["id"])] / row["mag"]))
        try:
            boot_model = fit_ols(boot_x_rows, boot_y)
        except ValueError:
            continue
        for sample_type in types_order:
            retention_series = build_monotone_retention_series(boot_model["beta"], sample_type, path_rows, features_for_path)
            bootstrap_retention_map[sample_type].append(retention_series[-1])

    decision_rows = []
    bootstrap_rows = []
    for sample in sample_rows:
        sample_type = sample["sample_type"]
        observed = to_float(sample["observed_magnetism_mt"])
        theta0 = theta_static(sample_type)
        retention = retention_last[sample_type]
        theta_dynamic = theta0 * retention
        m0_hat = observed / retention
        boot_ret = bootstrap_retention_map[sample_type]
        boot_m0 = [observed / max(r, 1e-12) for r in boot_ret]
        lower = percentile(boot_m0, 0.025) if boot_m0 else float("nan")
        upper = percentile(boot_m0, 0.975) if boot_m0 else float("nan")
        boot_sd = st.pstdev(boot_m0) if len(boot_m0) > 1 else 0.0
        relative_width = (upper - lower) / max(m0_hat, 1e-12) if not math.isnan(lower) and not math.isnan(upper) else float("nan")
        support_gap = m0_hat - theta0
        model_sd = abs(m0_hat * float(model["sigma"]))
        calibration_floor = CALIBRATION_FLOOR_RATIO * theta0
        combined_sd = math.sqrt(boot_sd**2 + model_sd**2 + calibration_floor**2)
        support_probability = normal_cdf(support_gap / combined_sd) if combined_sd > 0 else float("nan")
        boot_support = (sum(1 for v in boot_m0 if v >= theta0) + 0.5) / (len(boot_m0) + 1.0) if boot_m0 else float("nan")

        gray = math.isnan(relative_width) or relative_width > RELATIVE_WIDTH_THRESHOLD or retention < MIN_RETENTION_RATIO
        if gray:
            decision = "灰区/证据不足"
        elif lower >= theta0 and support_probability >= SUPPORT_PROB_ACCEPT:
            decision = "支持曾遭雷击"
        elif upper < theta0 and support_probability <= SUPPORT_PROB_REJECT:
            decision = "不支持曾遭雷击"
        else:
            decision = "灰区/证据不足"

        bootstrap_rows.append(
            [
                case_info["case_id"],
                sample["sample_id"],
                sample_type,
                delay_days,
                round(retention, 6),
                round(theta_dynamic, 6),
                round(m0_hat, 6),
                round(lower, 6) if not math.isnan(lower) else "",
                round(upper, 6) if not math.isnan(upper) else "",
                round(boot_sd, 6) if not math.isnan(boot_sd) else "",
                round(relative_width, 6) if not math.isnan(relative_width) else "",
                round(boot_support, 6) if not math.isnan(boot_support) else "",
                round(support_probability, 6) if not math.isnan(support_probability) else "",
            ]
        )
        decision_rows.append(
            [
                case_info["case_id"],
                sample["sample_id"],
                sample_type,
                delay_days,
                round(observed, 6),
                round(retention, 6),
                round(theta_dynamic, 6),
                round(m0_hat, 6),
                round(lower, 6) if not math.isnan(lower) else "",
                round(upper, 6) if not math.isnan(upper) else "",
                round(relative_width, 6) if not math.isnan(relative_width) else "",
                round(support_probability, 6) if not math.isnan(support_probability) else "",
                decision,
            ]
        )

    return retention_curve_rows, threshold_curve_rows, bootstrap_rows, decision_rows


def main() -> None:
    base = find_base_dir()
    data_dir = base / "原始数据"
    problem4_output_dir = base / "结果输出" / "问题四"
    problem4_output_dir.mkdir(parents=True, exist_ok=True)

    experiment_csv = next(p for p in data_dir.iterdir() if p.suffix.lower() == ".csv")
    weather_xlsx = next(p for p in data_dir.iterdir() if p.suffix.lower() == ".xlsx")
    observations, types_order = read_experiment(experiment_csv)
    weather_rows = read_xlsx_weather(weather_xlsx)
    fitted = fit_main_model(observations, types_order, weather_rows)

    case_rows = read_csv_rows(problem4_output_dir / "q4_case_info_template.csv")
    sample_rows = read_csv_rows(problem4_output_dir / "q4_sample_info_template.csv")
    weather_rows = read_csv_rows(problem4_output_dir / "q4_weather_series_template.csv")
    if not case_rows or not sample_rows or not weather_rows:
        raise ValueError("Problem 4 templates are incomplete")

    case_info = case_rows[0]
    dense_weather = normalize_weather_series(case_info, weather_rows)
    retention_curve_rows, threshold_curve_rows, bootstrap_rows, decision_rows = build_case_outputs(
        case_info,
        sample_rows,
        dense_weather,
        fitted,
    )

    write_csv(
        problem4_output_dir / "q4_case_retention_curve.csv",
        ["case_id", "sample_type", "day", "weather", "temp_c", "hum_pct", "retention_ratio"],
        retention_curve_rows,
    )
    write_csv(
        problem4_output_dir / "q4_case_threshold_curve.csv",
        ["case_id", "sample_type", "day", "theta_static_mt", "theta_dynamic_mt"],
        threshold_curve_rows,
    )
    write_csv(
        problem4_output_dir / "q4_case_bootstrap.csv",
        [
            "case_id",
            "sample_id",
            "sample_type",
            "delay_days",
            "retention_ratio",
            "dynamic_threshold_mt",
            "m0_hat_mt",
            "m0_hat_lower_mt",
            "m0_hat_upper_mt",
            "bootstrap_sd_mt",
            "relative_width",
            "bootstrap_support_rate",
            "support_probability",
        ],
        bootstrap_rows,
    )
    write_csv(
        problem4_output_dir / "q4_case_run_result.csv",
        [
            "case_id",
            "sample_id",
            "sample_type",
            "delay_days",
            "observed_magnetism_mt",
            "retention_ratio",
            "dynamic_threshold_mt",
            "m0_hat_mt",
            "m0_hat_lower_mt",
            "m0_hat_upper_mt",
            "relative_width",
            "support_probability",
            "decision",
        ],
        decision_rows,
    )
    print(f"Generated case outputs in: {problem4_output_dir}")


if __name__ == "__main__":
    main()
