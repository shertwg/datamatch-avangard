"""Схема валидации, имитирующая тест: «псевдо-сегодня» в прошлом + следующие матчи.

Ситуация на 15.12.2025: последний матч сыгран 07.12, впереди 17 матчей (остаток сезона),
на 6 из них продажи уже открыты. Каждый фолд воспроизводит её на прошлой дате as_of:
- train — все матчи, сыгранные до as_of (tickets_total.notna() в build_table(as_of));
- val   — следующие до n_val матчей после as_of (по умолчанию внутри того же сезона),
          признаки — из build_table(as_of), истинный итог — из build_table() на 15.12.2025.

Фолды по умолчанию (валидации не пересекаются):
- 2023-12-15: train 20 матчей, val M021–M034 (14), 5 уже в продаже — тот же момент календаря;
- 2024-11-25: train 51, val M052–M068 (17), 5 в продаже — ровно 17 матчей до конца сезона;
- 2025-09-20: train 73, val M074–M085 (12), 3 в продаже — текущий сезон 2025/26.
"""
from dataclasses import dataclass

import numpy as np
import pandas as pd

from src.data import build_table, load_raw

FOLDS = ("2023-12-15", "2024-11-25", "2025-09-20")
KEY = ["match_id", "zone"]


def rmse(y_true, y_pred):
    y_true, y_pred = np.asarray(y_true, dtype=float), np.asarray(y_pred, dtype=float)
    return float(np.sqrt(np.mean((y_true - y_pred) ** 2)))


def postprocess(pred, df):
    """Правила сдачи: sold_as_of <= прогноз <= capacity, целое.

    pred — массив длины len(df), df — строки из build_table с колонками sold_as_of и capacity.
    """
    pred = np.asarray(pred, dtype=float)
    if pred.shape != (len(df),):
        raise ValueError(f"predict_fn вернула {pred.shape}, ожидалось ({len(df)},)")
    if not np.isfinite(pred).all():
        raise ValueError("в прогнозе есть NaN/inf")
    low, high = df["sold_as_of"].to_numpy(), df["capacity"].to_numpy()
    return np.rint(np.clip(pred, low, high)).astype(int)


def val_matches(full, as_of, n_val=17, same_season=True):
    """match_id валидации: первые n_val матчей с известным итогом, сыгранных в день as_of или позже."""
    as_of = pd.Timestamp(as_of)
    m = (full[full["tickets_total"].notna()]
         .drop_duplicates("match_id")
         .sort_values(["match_date", "match_id"]))
    m = m[m["match_date"] >= as_of]
    if same_season:
        m = m[m["season"] == m["season"].iloc[0]]
    return list(m["match_id"].head(n_val))


def make_fold(as_of, full=None, n_val=17, same_season=True):
    """(train, val, y_true) для одной даты as_of.

    train — строки build_table(as_of) с известным tickets_total.
    val   — строки build_table(as_of) валидационных матчей, tickets_total в них NaN.
    y_true — итог продаж по строкам val (тот же порядок) из полной таблицы.
    """
    full = build_table() if full is None else full
    cut = build_table(as_of)
    ids = val_matches(full, as_of, n_val, same_season)

    train = cut[cut["tickets_total"].notna()].reset_index(drop=True)
    val = cut[cut["match_id"].isin(ids)].sort_values(KEY).reset_index(drop=True)
    assert val["tickets_total"].isna().all(), "в val виден таргет — утечка"
    assert not set(ids) & set(train["match_id"]), "матч одновременно в train и val"

    y_true = val[KEY].merge(full[KEY + ["tickets_total"]], on=KEY, how="left",
                            validate="one_to_one")["tickets_total"]
    assert y_true.notna().all()
    return train, val, y_true.astype(int).to_numpy()


@dataclass
class EvalResult:
    """Результат evaluate.

    preds     — fold, match_id, zone, in_sale, y_true, y_pred: по строке на (фолд, матч, зона);
    overall   — RMSE по всем строкам всех фолдов;
    by_fold   — RMSE, число матчей и строк по каждому фолду;
    by_zone   — RMSE: зоны × (фолды + «все»);
    by_status — RMSE для матчей, уже в продаже на as_of, и ещё не в продаже.
    """

    preds: pd.DataFrame
    overall: float
    by_fold: pd.DataFrame
    by_zone: pd.DataFrame
    by_status: pd.DataFrame

    def __repr__(self):
        return f"EvalResult(RMSE={self.overall:.2f}, folds={list(self.by_fold.index.drop('all'))})"


def evaluate(predict_fn, folds=FOLDS, n_val=17, same_season=True):
    """Прогоняет модель через валидацию по фолдам «псевдо-сегодня».

    predict_fn(train, val, as_of) -> array-like длины len(val):
        train : pd.DataFrame — строки build_table(as_of) по матчам, сыгранным до as_of,
                с таргетом tickets_total (копия, можно менять);
        val   : pd.DataFrame — строки build_table(as_of) валидационных матчей (8 зон на матч,
                отсортированы по match_id, zone), tickets_total = NaN; sold_as_of и прочие
                признаки продаж посчитаны строго по date < as_of;
        as_of : pd.Timestamp — дата фолда; по ней же обрезайте внешние признаки
                (khl_games, season_tickets_daily): только date < as_of.
        Возвращает прогноз tickets_total в порядке строк val. Его не нужно клиппировать
        и округлять: evaluate сам применяет postprocess (sold_as_of <= y <= capacity, целое),
        так все модели сравниваются по одним правилам сдачи.

    folds — даты as_of (по умолчанию FOLDS); можно передать свои.
    Порядок строк фиксирован (фолды в порядке folds, внутри — match_id, zone), строки
    не перемешиваются, поэтому при детерминированной predict_fn результат воспроизводим.

    Возвращает EvalResult: .overall (общий RMSE), .by_fold, .by_zone, .by_status, .preds.
    """
    full = build_table()
    parts = []
    for as_of in folds:
        as_of = pd.Timestamp(as_of)
        train, val, y_true = make_fold(as_of, full, n_val, same_season)
        y_pred = postprocess(predict_fn(train.copy(), val.copy(), as_of), val)
        parts.append(pd.DataFrame({
            "fold": str(as_of.date()),
            "match_id": val["match_id"],
            "zone": val["zone"],
            "in_sale": val["sales_started"].astype(bool),
            "y_true": y_true,
            "y_pred": y_pred,
        }))
    preds = pd.concat(parts, ignore_index=True)

    def agg(g):
        return pd.Series({"rmse": rmse(g["y_true"], g["y_pred"]),
                          "matches": g["match_id"].nunique(), "rows": len(g)})

    cols = ["y_true", "y_pred", "match_id"]
    by_fold = preds.groupby("fold", sort=False)[cols].apply(agg)
    by_fold.loc["all"] = agg(preds)
    by_fold[["matches", "rows"]] = by_fold[["matches", "rows"]].astype(int)

    zone_rmse = lambda g: rmse(g["y_true"], g["y_pred"])
    by_zone = preds.groupby(["zone", "fold"], sort=False)[cols].apply(zone_rmse).unstack("fold")
    by_zone = by_zone[list(dict.fromkeys(preds["fold"]))].sort_index()
    by_zone["all"] = preds.groupby("zone")[cols].apply(zone_rmse)

    by_status = preds.groupby("in_sale")[cols].apply(agg)
    by_status.index = by_status.index.map({True: "в продаже", False: "ещё не в продаже"})
    by_status[["matches", "rows"]] = by_status[["matches", "rows"]].astype(int)

    return EvalResult(preds, rmse(preds["y_true"], preds["y_pred"]), by_fold, by_zone, by_status)


def predict_test(predict_fn):
    """Обучает predict_fn на всех 85 train-матчах (as_of = 15.12.2025) и прогнозирует test.

    Возвращает DataFrame match_id, zone, tickets_total в порядке submission_template.csv
    (136 строк, целые, sold_as_of <= tickets_total <= capacity).
    """
    df = build_table()
    as_of = pd.Timestamp("2025-12-15")
    train = df[df["tickets_total"].notna()].reset_index(drop=True)
    test = df[df["is_test"] == 1].sort_values(KEY).reset_index(drop=True)
    pred = postprocess(predict_fn(train.copy(), test.copy(), as_of), test)
    test["tickets_total"] = pred

    tmpl = load_raw()["submission"][KEY]
    sub = tmpl.merge(test[KEY + ["tickets_total"]], on=KEY, how="left", validate="one_to_one")
    assert len(sub) == 136 and sub["tickets_total"].notna().all()
    return sub
