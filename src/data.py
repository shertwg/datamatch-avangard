"""Сборка данных для всех ноутбуков: загрузка CSV и таблица match_id × zone на дату as_of.

Допущения:
- seats и season_tickets из zones.csv считаем известными заранее для любого матча
  (организаторы дают их и для test), поэтому они не обрезаются по as_of.
- Отсутствующая строка в sales_daily.csv = 0 продаж в этот день.
- Признаки из khl_games и season_tickets_daily здесь не строятся (делаются отдельно);
  khl_games с обрезкой по as_of отдаёт load_khl.
"""
from pathlib import Path

import numpy as np
import pandas as pd

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
CUTOFF = pd.Timestamp("2025-12-15")

FILES = {
    "matches": ("matches.csv", ["date", "sales_open"]),
    "sales": ("sales_daily.csv", ["date"]),
    "zones": ("zones.csv", []),
    "season_tickets": ("season_tickets_daily.csv", ["date"]),
    "khl": ("khl_games.csv", ["date"]),
    "submission": ("submission_template.csv", []),
}


def load_raw(data_dir=DATA_DIR):
    """Читает все CSV из data_dir в словарь датафреймов, даты распарсены."""
    data_dir = Path(data_dir)
    return {
        name: pd.read_csv(data_dir / fname, encoding="utf-8", parse_dates=dates)
        for name, (fname, dates) in FILES.items()
    }


def _sales_sum(sales, start, end, name):
    """Сумма tickets по (match_id, zone) для продаж с start <= date < end."""
    mask = (sales["date"] < end) if start is None else sales["date"].between(start, end, inclusive="left")
    return sales[mask].groupby(["match_id", "zone"])["tickets"].sum().rename(name)


def _check_as_of(as_of):
    as_of = pd.Timestamp(as_of)
    if as_of > CUTOFF:
        raise ValueError(f"as_of={as_of.date()} позже даты отсечки {CUTOFF.date()} — утечка")
    return as_of


def load_daily_sales(as_of="2025-12-15", data_dir=DATA_DIR):
    """Строки sales_daily.csv с date < as_of (match_id, zone, date, days_before, tickets)."""
    as_of = _check_as_of(as_of)
    sales = load_raw(data_dir)["sales"]
    return sales[sales["date"] < as_of].reset_index(drop=True)


def load_khl(as_of="2025-12-15", data_dir=DATA_DIR):
    """Строки khl_games.csv с date < as_of (season, date, stage, home, away, home_goals, away_goals)."""
    as_of = _check_as_of(as_of)
    khl = load_raw(data_dir)["khl"]
    return khl[khl["date"] < as_of].reset_index(drop=True)


def build_table(as_of="2025-12-15", data_dir=DATA_DIR):
    """Таблица match_id × zone (все 102 матча × 8 зон) с признаками на момент as_of.

    Всё, что зависит от продаж, берётся строго с date < as_of. Таргет tickets_total есть только
    у матчей, сыгранных до as_of: все их продажи были не позже дня матча, то есть раньше as_of,
    поэтому итог равен sold_as_of и утечки нет. days_to_match < 0 — матч уже сыгран.
    """
    as_of = _check_as_of(as_of)

    raw = load_raw(data_dir)
    m = raw["matches"].rename(columns={"date": "match_date"})
    sales = raw["sales"]

    df = raw["zones"].merge(m, on="match_id", how="left", validate="many_to_one")
    df["hour"] = df["time"].str[:2].astype(int)
    df["month"] = df["match_date"].dt.month
    df["is_weekend"] = (df["weekday"] >= 6).astype(int)
    df["is_test"] = (df["part"] == "test").astype(int)
    df["capacity"] = df["seats"] - df["season_tickets"]

    key = ["match_id", "zone"]
    for s in [
        _sales_sum(sales, None, as_of, "sold_as_of"),
        _sales_sum(sales, as_of - pd.Timedelta(days=7), as_of, "sold_last_7d"),
        _sales_sum(sales, as_of - pd.Timedelta(days=14), as_of, "sold_last_14d"),
    ]:
        df = df.merge(s, left_on=key, right_index=True, how="left")
        df[s.name] = df[s.name].fillna(0).astype(int)

    df["sales_started"] = (df["sales_open"] < as_of).astype(int)
    df["days_to_match"] = (df["match_date"] - as_of).dt.days
    df["days_since_open"] = (as_of - df["sales_open"]).dt.days.clip(lower=0)

    played = df["match_date"] < as_of
    df["tickets_total"] = np.where(played, df["sold_as_of"], np.nan)

    return df.sort_values(key).reset_index(drop=True)


if __name__ == "__main__":
    df = build_table()
    train, test = df[df["is_test"] == 0], df[df["is_test"] == 1]
    print(f"Всего строк: {len(df)}  train: {len(train)}  test: {len(test)}")
    assert len(train) == 85 * 8 and len(test) == 17 * 8

    assert not df.duplicated(["match_id", "zone"]).any(), "дубли match_id × zone"
    t = df["tickets_total"].dropna()
    assert (t >= 0).all(), "отрицательный tickets_total"
    assert (t <= df.loc[t.index, "capacity"]).all(), "tickets_total > capacity"
    assert df.loc[df["is_test"] == 1, "tickets_total"].isna().all(), "у test есть таргет"
    assert df.loc[df["is_test"] == 0, "tickets_total"].notna().all(), "у train нет таргета"

    started = test.groupby("match_id")["sales_started"].max()
    print("Test-матчи с начавшимися продажами:", list(started[started == 1].index))
    assert started.sum() == 6

    print(f"Макс. заполняемость train: {(t / df.loc[t.index, 'capacity']).max():.3f}")
    print("sold_as_of по test-матчам:")
    print(test.groupby("match_id")[["sold_as_of", "sold_last_7d", "sold_last_14d"]].sum()
          .query("sold_as_of > 0").to_string())

    print("\nПропуски по колонкам:")
    na = df.isna().sum()
    print(na[na > 0].to_string() if na.any() else "нет")

    past = build_table(as_of="2025-11-01")
    late = past["match_date"] >= pd.Timestamp("2025-11-01")
    assert past.loc[late, "tickets_total"].isna().all()
    assert (past["sold_as_of"] <= df["sold_as_of"]).all()
    print(f"\nas_of=2025-11-01: строк с таргетом {past['tickets_total'].notna().sum()}, "
          f"матчей с начавшимися продажами без итога: "
          f"{past.loc[(past['sales_started'] == 1) & late, 'match_id'].nunique()}")

    daily = load_daily_sales("2025-11-01")
    assert (daily["date"] < pd.Timestamp("2025-11-01")).all()
    agg = daily.groupby(["match_id", "zone"])["tickets"].sum()
    assert agg.reindex(pd.MultiIndex.from_frame(past[["match_id", "zone"]]), fill_value=0).to_numpy().tolist() \
        == past["sold_as_of"].tolist(), "load_daily_sales расходится с sold_as_of"
    try:
        load_daily_sales("2025-12-16")
        raise AssertionError("load_daily_sales пропустила дату после отсечки")
    except ValueError:
        pass
    print(f"load_daily_sales(2025-11-01): {len(daily)} строк, сходится с sold_as_of")

    khl_all = load_raw()["khl"]
    khl = load_khl("2025-11-01")
    assert (khl["date"] < pd.Timestamp("2025-11-01")).all()
    assert len(khl) == (khl_all["date"] < pd.Timestamp("2025-11-01")).sum()
    assert len(load_khl()) <= len(khl_all)
    try:
        load_khl("2025-12-16")
        raise AssertionError("load_khl пропустила дату после отсечки")
    except ValueError:
        pass
    print(f"load_khl(2025-11-01): {len(khl)} из {len(khl_all)} игр, "
          f"последняя {khl['date'].max().date()}; load_khl(): {len(load_khl())} игр")

    print("\nКолонки:", list(df.columns))
    print("Все проверки пройдены.")
