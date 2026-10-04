"""Тест на утечку: портим данные после as_of и проверяем, что признаки и прогнозы не меняются.

Для каждого фолда из src.validation.FOLDS:
- в sales_daily все строки с date >= as_of: tickets * 10 + 500;
- в khl_games все матчи с date >= as_of: счёт переставлен и к каждому + 7 голов;
- сравниваем с исходными данными train, val (без таргета) и прогнозы zone_mean, smart, catboost.

Если хоть что-то изменилось, значит, будущее после as_of доходит до модели.
Порча делается в памяти (подмена src.data.load_raw), файлы в data/ не меняются.

Запуск из корня репо: python -m tests.test_leakage  (CatBoost на 3 фолдах, ~30 с)
"""
import sys

import numpy as np
import pandas as pd

import src.data as data
from src.baseline import smart_baseline_predict, zone_mean_predict
from src.model import catboost_predict
from src.validation import FOLDS, make_fold

MODELS = {"zone_mean": zone_mean_predict, "smart": smart_baseline_predict, "catboost": catboost_predict}
_load_raw = data.load_raw


def poisoned_loader(as_of):
    as_of = pd.Timestamp(as_of)

    def load_raw(data_dir=data.DATA_DIR):
        raw = _load_raw(data_dir)
        sales = raw["sales"].copy()
        fut = sales["date"] >= as_of
        sales.loc[fut, "tickets"] = sales.loc[fut, "tickets"] * 10 + 500
        khl = raw["khl"].copy()
        fut_k = khl["date"] >= as_of
        goals = khl.loc[fut_k, ["away_goals", "home_goals"]].to_numpy()
        khl.loc[fut_k, ["home_goals", "away_goals"]] = goals + 7
        load_raw.counts = int(fut.sum()), int(fut_k.sum())
        return {**raw, "sales": sales, "khl": khl}

    return load_raw


def run_fold(as_of):
    train, val, _ = make_fold(as_of)
    preds = {name: np.asarray(fn(train, val, as_of)) for name, fn in MODELS.items()}
    return train, val.drop(columns="tickets_total"), preds


def check_fold(as_of):
    clean = run_fold(as_of)
    loader = poisoned_loader(as_of)
    data.load_raw = loader
    try:
        dirty = run_fold(as_of)
    finally:
        data.load_raw = _load_raw

    n_sales, n_khl = loader.counts
    checks = {"train": clean[0].equals(dirty[0]), "val (признаки)": clean[1].equals(dirty[1])}
    checks |= {f"прогноз {name}": np.array_equal(clean[2][name], dirty[2][name]) for name in MODELS}
    print(f"\nФолд {as_of}: испорчено {n_sales} строк продаж и {n_khl} матчей КХЛ после as_of")
    for name, ok in checks.items():
        print(f"  {'OK  ' if ok else 'FAIL'} {name}")
    return all(checks.values())


def main():
    results = [check_fold(as_of) for as_of in FOLDS]
    ok = all(results)
    print(f"\n{'OK' if ok else 'FAIL'}: {sum(results)} из {len(results)} фолдов без утечки")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
