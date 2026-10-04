"""Тесты замороженной модели src/model.py: согласованность признаков, повторяемость, неизменность прогноза.

1. Срезы против build_table: признаки train-срезов (sold_as_of, sold_last_7d, sold_last_14d, days_left)
   совпадают с build_table(as_of=дата среза). Модель учится на тех же признаках, что видит при прогнозе.
2. Число потоков: прогноз с thread_count=1 совпадает с прогнозом по умолчанию (все ядра), то есть
   predictions.csv не зависит от железа.
3. Заморозка: predict_test(catboost_predict) совпадает с predictions.csv в репозитории до билета.
   Если тест упал, значит, изменились модель или данные, и цифры отчёта и презентации надо пересчитать.

Запуск из корня репо: python -m tests.test_model  (~30 с)
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import src.model as model
from src.data import build_table, load_daily_sales
from src.validation import predict_test

ROOT = Path(__file__).resolve().parent.parent
SNAP_STEP = 10  # проверяем каждую 10-ю дату среза (~100 вызовов build_table)


def check_snapshots():
    full = build_table()
    train = full[full["tickets_total"].notna()]
    snaps = model._train_snapshots(train, load_daily_sales())
    dates = sorted(snaps["snap_date"].unique())[::SNAP_STEP]
    cols = ["sold_as_of", "sold_last_7d", "sold_last_14d"]
    checked = mismatches = 0
    for d in dates:
        s = snaps[snaps["snap_date"] == d].set_index(model.KEY)
        bt = build_table(as_of=d).set_index(model.KEY).loc[s.index]
        bt_days_left = np.minimum(bt["days_to_match"], (bt["match_date"] - bt["sales_open"]).dt.days)
        mismatches += int((s[cols] != bt[cols]).to_numpy().sum() + (s["days_left"] != bt_days_left).sum())
        checked += s.shape[0] * (len(cols) + 1)
    print(f"  {len(dates)} дат среза из {snaps['snap_date'].nunique()}, {checked} значений, расхождений: {mismatches}")
    return mismatches == 0


def check_threads():
    default = predict_test(model.catboost_predict)
    params = model.CB_PARAMS
    model.CB_PARAMS = {**params, "thread_count": 1}
    try:
        single = predict_test(model.catboost_predict)
    finally:
        model.CB_PARAMS = params
    diff = int((default["tickets_total"] != single["tickets_total"]).sum())
    print(f"  строк с разным прогнозом (все ядра против 1 потока): {diff} из {len(default)}")
    return diff == 0


def check_frozen():
    saved = pd.read_csv(ROOT / "predictions.csv", encoding="utf-8")
    sub = predict_test(model.catboost_predict)
    same = sub.equals(saved)
    print(f"  сумма прогноза: сейчас {sub['tickets_total'].sum()}, в predictions.csv {saved['tickets_total'].sum()}; "
          f"{'совпадает до билета' if same else 'РАСХОДИТСЯ'}")
    return same


CHECKS = {
    "срезы train совпадают с build_table(as_of)": check_snapshots,
    "прогноз не зависит от числа потоков": check_threads,
    "прогноз совпадает с predictions.csv": check_frozen,
}


def main():
    results = []
    for name, check in CHECKS.items():
        print(f"\n{name}")
        ok = check()
        print(f"  {'OK' if ok else 'FAIL'}")
        results.append(ok)
    ok = all(results)
    print(f"\n{'OK' if ok else 'FAIL'}: {sum(results)} из {len(results)} проверок пройдены")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
