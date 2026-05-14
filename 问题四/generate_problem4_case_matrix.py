#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import math
import random
import statistics as st
from datetime import timedelta
from pathlib import Path

import run_problem4_case as q4


DELAYS = [7, 30, 60]
WEATHER_SCENARIOS = ["normal", "humid_hot"]
TARGET_DECISIONS = ["支持曾遭雷击", "灰区/证据不足", "不支持曾遭雷击"]
BOOTSTRAP_REPS_MATRIX = 200


def write_csv(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def choose_windows(weather_rows: list[dict[str, object]], delays: list[int]):
    by_seq = {int(r["seq"]): r for r in weather_rows}
    windows = {}
    for delay in delays:
        candidates = []
        for start in range(1, 91 - delay + 1):
            chunk = [by_seq[d] for d in range(start, start + delay)]
            temps = [float(r["temp"]) for r in chunk]
            hums = [float(r["hum"]) for r in chunk]
            score = st.mean([t * h for t, h in zip(temps, hums)])
            candidates.append((score, start, chunk))
        candidates.sort(key=lambda x: x[0])
        windows[("normal", delay)] = candidates[len(candidates) // 2][2]
        windows[("humid_hot", delay)] = candidates[-1][2]
    return windows


def build_case_info(case_id: str, delay_days: int):
    lightning_date = q4.parse_date("2026-07-01")
    measurement_date = lightning_date + timedelta(days=delay_days)
    return {
        "case_id": case_id,
        "lightning_date": lightning_date.strftime("%Y-%m-%d"),
        "measurement_date": measurement_date.strftime("%Y-%m-%d"),
        "sampling_date": measurement_date.strftime("%Y-%m-%d"),
        "site_name": "representative_case",
        "site_lon": "",
        "site_lat": "",
        "weather_source": "training_weather_window",
    }, lightning_date


def convert_chunk_to_case_weather(chunk: list[dict[str, object]], lightning_date):
    rows = []
    for idx, item in enumerate(chunk, start=1):
        rows.append(
            {
                "date": (lightning_date + timedelta(days=idx)).strftime("%Y-%m-%d"),
                "day_index": idx,
                "temp_c": float(item["temp"]),
                "hum_pct": float(item["hum"]),
                "weather_text": str(item["weather"]),
                "station_id": "training_path",
                "is_imputed": "false",
            }
        )
    return rows


def candidate_m0_values(theta0: float, target: str) -> list[float]:
    if target == "支持曾遭雷击":
        return [theta0 * x for x in [2.50, 2.20, 2.00, 1.80, 1.60, 1.40]]
    if target == "不支持曾遭雷击":
        return [theta0 * x for x in [0.05, 0.10, 0.20, 0.30, 0.40, 0.50]]
    return [theta0 * x for x in [1.02, 1.00, 0.98, 0.95, 1.05, 0.90, 1.10]]


def build_training_cum_values(training_weather, stats, day: int):
    cum_temp = cum_hum = cum_inter = 0.0
    for d in range(1, day + 1):
        tc = training_weather[d]["temp"] - float(stats["temp_mean"])
        hc = training_weather[d]["hum"] - float(stats["hum_mean"])
        cum_temp += tc
        cum_hum += hc
        cum_inter += tc * hc
    return cum_temp, cum_hum, cum_inter


def build_scenario_profile(fitted, sample_type: str, weather_series, bootstrap_reps: int = BOOTSTRAP_REPS_MATRIX):
    stats = fitted["stats"]
    model = fitted["model"]
    features_for_path = fitted["features_for_path"]
    path_rows = q4.cumulative_path_features(weather_series, stats)
    point_series = q4.build_monotone_retention_series(model["beta"], sample_type, path_rows, features_for_path)
    point_retention = point_series[-1]

    cluster_rows = fitted["cluster_rows"]
    cluster_keys = list(cluster_rows.keys())
    training_weather = fitted["weather"]
    rng = random.Random(20260514)
    boot_retentions = []

    for _ in range(bootstrap_reps):
        sampled_keys = [rng.choice(cluster_keys) for _ in cluster_keys]
        boot_train_rows = []
        for key in sampled_keys:
            boot_train_rows.extend(cluster_rows[key])
        boot_x_rows = []
        boot_y = []
        for row in boot_train_rows:
            day = int(row["day"])
            cum_temp, cum_hum, cum_inter = build_training_cum_values(training_weather, stats, day)
            boot_x_rows.append(features_for_path(row["type"], day, cum_temp, cum_hum, cum_inter))
            boot_y.append(math.log(fitted["m0"][(row["type"], row["id"])] / row["mag"]))
        try:
            boot_model = q4.fit_ols(boot_x_rows, boot_y)
        except ValueError:
            continue
        boot_series = q4.build_monotone_retention_series(boot_model["beta"], sample_type, path_rows, features_for_path)
        boot_retentions.append(boot_series[-1])

    return {
        "path_rows": path_rows,
        "retention_series": point_series,
        "retention_ratio": point_retention,
        "bootstrap_retentions": boot_retentions,
    }


def evaluate_case(sample_type: str, theta0: float, observed_mt: float, profile, sigma_log: float):
    retention = profile["retention_ratio"]
    dynamic_threshold = theta0 * retention
    m0_hat = observed_mt / max(retention, 1e-12)
    boot_m0 = [observed_mt / max(r, 1e-12) for r in profile["bootstrap_retentions"]]
    lower = q4.percentile(boot_m0, 0.025) if boot_m0 else float("nan")
    upper = q4.percentile(boot_m0, 0.975) if boot_m0 else float("nan")
    boot_sd = st.pstdev(boot_m0) if len(boot_m0) > 1 else 0.0
    relative_width = (upper - lower) / max(m0_hat, 1e-12) if not math.isnan(lower) and not math.isnan(upper) else float("nan")
    support_gap = m0_hat - theta0
    model_sd = abs(m0_hat * sigma_log)
    calibration_floor = q4.CALIBRATION_FLOOR_RATIO * theta0
    combined_sd = math.sqrt(boot_sd**2 + model_sd**2 + calibration_floor**2)
    support_probability = q4.normal_cdf(support_gap / combined_sd) if combined_sd > 0 else float("nan")
    bootstrap_support_rate = (sum(1 for v in boot_m0 if v >= theta0) + 0.5) / (len(boot_m0) + 1.0) if boot_m0 else float("nan")

    gray = (
        math.isnan(relative_width)
        or relative_width > q4.RELATIVE_WIDTH_THRESHOLD
        or retention < q4.MIN_RETENTION_RATIO
    )
    if gray:
        decision = "灰区/证据不足"
    elif lower >= theta0 and support_probability >= q4.SUPPORT_PROB_ACCEPT:
        decision = "支持曾遭雷击"
    elif upper < theta0 and support_probability <= q4.SUPPORT_PROB_REJECT:
        decision = "不支持曾遭雷击"
    else:
        decision = "灰区/证据不足"

    return {
        "retention_ratio": retention,
        "dynamic_threshold_mt": dynamic_threshold,
        "m0_hat_mt": m0_hat,
        "m0_hat_lower_mt": lower,
        "m0_hat_upper_mt": upper,
        "bootstrap_sd_mt": boot_sd,
        "relative_width": relative_width,
        "bootstrap_support_rate": bootstrap_support_rate,
        "support_probability": support_probability,
        "decision": decision,
    }


def pick_case_for_target(sample_type: str, profile, sigma_log: float, target: str):
    theta0 = q4.theta_static(sample_type)
    for m0_target in candidate_m0_values(theta0, target):
        observed_mt = m0_target * profile["retention_ratio"]
        result = evaluate_case(sample_type, theta0, observed_mt, profile, sigma_log)
        if result["decision"] == target:
            return observed_mt, m0_target, result, True
    observed_mt = candidate_m0_values(theta0, target)[0] * profile["retention_ratio"]
    result = evaluate_case(sample_type, theta0, observed_mt, profile, sigma_log)
    return observed_mt, candidate_m0_values(theta0, target)[0], result, False


def main():
    base = q4.find_base_dir()
    data_dir = base / "原始数据"
    out_dir = base / "结果输出" / "问题四"
    out_dir.mkdir(parents=True, exist_ok=True)

    experiment_csv = next(p for p in data_dir.iterdir() if p.suffix.lower() == ".csv")
    weather_xlsx = next(p for p in data_dir.iterdir() if p.suffix.lower() == ".xlsx")
    observations, types_order = q4.read_experiment(experiment_csv)
    weather_rows = q4.read_xlsx_weather(weather_xlsx)
    fitted = q4.fit_main_model(observations, types_order, weather_rows)
    sigma_log = float(fitted["model"]["sigma"])
    windows = choose_windows(weather_rows, DELAYS)

    matrix_rows = []
    weather_detail_rows = []
    retention_rows = []
    threshold_rows = []
    bootstrap_rows = []
    summary_rows = []
    overall_hits = []
    case_counter = 1

    for sample_type in types_order:
        for delay in DELAYS:
            for weather_name in WEATHER_SCENARIOS:
                case_id_base = f"REP-{case_counter:03d}"
                case_counter += 1
                case_info, lightning_date = build_case_info(case_id_base, delay)
                weather_series = convert_chunk_to_case_weather(windows[(weather_name, delay)], lightning_date)
                profile = build_scenario_profile(fitted, sample_type, weather_series)
                if profile["retention_ratio"] < q4.MIN_RETENTION_RATIO:
                    target_list = ["灰区/证据不足"]
                else:
                    target_list = TARGET_DECISIONS

                for idx, target in enumerate(target_list, start=1):
                    observed_mt, target_m0, result, hit = pick_case_for_target(sample_type, profile, sigma_log, target)
                    case_id = f"{case_id_base}-{idx}"
                    overall_hits.append(1 if hit else 0)
                    matrix_rows.append(
                        [
                            case_id,
                            sample_type,
                            delay,
                            weather_name,
                            target,
                            round(target_m0, 6),
                            round(observed_mt, 6),
                            round(result["retention_ratio"], 6),
                            round(result["dynamic_threshold_mt"], 6),
                            round(result["m0_hat_mt"], 6),
                            round(result["m0_hat_lower_mt"], 6) if not math.isnan(result["m0_hat_lower_mt"]) else "",
                            round(result["m0_hat_upper_mt"], 6) if not math.isnan(result["m0_hat_upper_mt"]) else "",
                            round(result["relative_width"], 6) if not math.isnan(result["relative_width"]) else "",
                            round(result["support_probability"], 6) if not math.isnan(result["support_probability"]) else "",
                            result["decision"],
                            "yes" if hit else "no",
                        ]
                    )
                    bootstrap_rows.append(
                        [
                            case_id,
                            sample_type,
                            delay,
                            weather_name,
                            target,
                            round(result["bootstrap_sd_mt"], 6) if not math.isnan(result["bootstrap_sd_mt"]) else "",
                            round(result["bootstrap_support_rate"], 6) if not math.isnan(result["bootstrap_support_rate"]) else "",
                            round(result["support_probability"], 6) if not math.isnan(result["support_probability"]) else "",
                        ]
                    )
                    for wr in weather_series:
                        weather_detail_rows.append(
                            [case_id, sample_type, delay, weather_name, target, wr["day_index"], wr["date"], wr["temp_c"], wr["hum_pct"], wr["weather_text"]]
                        )
                    for path_row, retention in zip(profile["path_rows"], profile["retention_series"]):
                        retention_rows.append(
                            [case_id, sample_type, delay, weather_name, target, path_row["day"], path_row["weather"], round(path_row["temp_c"], 6), round(path_row["hum_pct"], 6), round(retention, 6)]
                        )
                        threshold_rows.append(
                            [case_id, sample_type, delay, weather_name, target, path_row["day"], round(q4.theta_static(sample_type), 6), round(q4.theta_static(sample_type) * retention, 6)]
                        )

    for target in TARGET_DECISIONS:
        subset = [row for row in matrix_rows if row[4] == target]
        hits = sum(1 for row in subset if row[15] == "yes")
        summary_rows.append([target, len(subset), hits, round(hits / len(subset), 6) if subset else ""])
    summary_rows.append(["overall", len(matrix_rows), sum(overall_hits), round(sum(overall_hits) / len(overall_hits), 6) if overall_hits else ""])

    write_csv(
        out_dir / "q4_representative_case_matrix.csv",
        ["case_id", "sample_type", "delay_days", "weather_scenario", "target_decision", "target_initial_mt", "observed_magnetism_mt", "retention_ratio", "dynamic_threshold_mt", "m0_hat_mt", "m0_hat_lower_mt", "m0_hat_upper_mt", "relative_width", "support_probability", "actual_decision", "target_hit"],
        matrix_rows,
    )
    write_csv(
        out_dir / "q4_representative_case_weather.csv",
        ["case_id", "sample_type", "delay_days", "weather_scenario", "target_decision", "day", "date", "temp_c", "hum_pct", "weather_text"],
        weather_detail_rows,
    )
    write_csv(
        out_dir / "q4_representative_case_retention.csv",
        ["case_id", "sample_type", "delay_days", "weather_scenario", "target_decision", "day", "weather", "temp_c", "hum_pct", "retention_ratio"],
        retention_rows,
    )
    write_csv(
        out_dir / "q4_representative_case_threshold.csv",
        ["case_id", "sample_type", "delay_days", "weather_scenario", "target_decision", "day", "theta_static_mt", "theta_dynamic_mt"],
        threshold_rows,
    )
    write_csv(
        out_dir / "q4_representative_case_bootstrap.csv",
        ["case_id", "sample_type", "delay_days", "weather_scenario", "target_decision", "bootstrap_sd_mt", "bootstrap_support_rate", "support_probability"],
        bootstrap_rows,
    )
    write_csv(
        out_dir / "q4_representative_case_summary.csv",
        ["target_decision", "case_count", "target_hit_count", "target_hit_rate"],
        summary_rows,
    )

    print(f"generated_matrix_cases={len(matrix_rows)}")
    print(f"overall_hit_rate={sum(overall_hits) / len(overall_hits):.6f}")


if __name__ == "__main__":
    main()
