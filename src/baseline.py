"""Baseline-модели по контракту predict_fn(train, val, as_of) из src.validation.

- zone_mean_predict: среднее tickets_total по зоне за все прошлые матчи (RMSE 141.0).
- smart_baseline_predict: среднее по зоне за окно season, N=10 (RMSE 135.2).
  Параметры выбраны сеткой окно {last, season} × N {5, 10, 15} × вариант {A, B}
  в notebooks/02_baseline.ipynb; выбор сделан только по валидации.
"""
import numpy as np
import pandas as pd

WINDOW = "season"
N = 10


def zone_level(train, val, n, window):
    """Среднее tickets_total по зоне за окно матчей до as_of; для каждой строки val.

    window="last"   — последние n сыгранных матчей (через границу сезонов);
    window="season" — все сыгранные матчи сезона строки val, если их >= n,
                      иначе добираем до n последними матчами прошлых сезонов.
    """
    played = train.drop_duplicates("match_id").sort_values(["match_date", "match_id"])
    out = pd.Series(np.nan, index=val.index)
    for season, idx in val.groupby("season").groups.items():
        if window == "last":
            ids = played["match_id"].tail(n)
        else:
            cur = played[played["season"] == season]
            past = played[played["season"] != season]
            ids = pd.concat([past["match_id"].tail(max(n - len(cur), 0)), cur["match_id"]])
        level = train[train["match_id"].isin(ids)].groupby("zone")["tickets_total"].mean()
        out[idx] = val.loc[idx, "zone"].map(level)
    return out.to_numpy()


def zone_mean_predict(train, val, as_of):
    """Baseline: среднее tickets_total по зоне по всем матчам train."""
    zone_mean = train.groupby("zone")["tickets_total"].mean()
    return val["zone"].map(zone_mean).to_numpy()


def smart_baseline_predict(train, val, as_of):
    """Умный baseline: уровень зоны за окно WINDOW/N, не ниже уже проданного (sold_as_of).

    Клиппинг по capacity и округление делает postprocess в evaluate/predict_test.
    """
    level = zone_level(train, val, N, WINDOW)
    return np.maximum(level, val["sold_as_of"].to_numpy())
