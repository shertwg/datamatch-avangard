"""CatBoost-модель по контракту predict_fn(train, val, as_of) из src.validation.

Идея: на дату прогноза часть матчей уже в продаже, часть ещё нет. Чтобы модель видела обе
ситуации, каждый сыгранный матч разворачивается в «срезы»: состояние продаж за h дней до матча
(h из SNAP_DAYS) плюс срез «продажи ещё не открыты». Признаки среза считаются только по данным
с date < даты среза (матчи, продажи, khl_games), таргет — полный итог матча в зоне.

Признаки:
- уровень зоны: среднее tickets_total по зоне за текущий сезон (окно как у smart baseline)
  и за последние 5 матчей;
- продажи на дату среза: sold_as_of, sold_last_7d, sold_last_14d, доля прошедшего окна продаж,
  дней до матча, темп (продано / типичная доля к этому дню), продано / уровень, продано / capacity,
  продажи всего матча относительно ожидаемых к этому дню;
- матч: зона, день недели, час, месяц, capacity, season_tickets;
- соперник: во сколько раз матчи с ним в прошлом были выше/ниже уровня зоны, число таких матчей;
- КХЛ: доля побед и средняя разница шайб соперника в сезоне, доля побед Авангарда в сезоне,
  форма Авангарда за последние 10 игр.

Таргет — остаток продаж (итог − sold_as_of), нормированный двумя способами; учатся две модели:
- на уровень зоны — для матчей, продажи на которые ещё не открыты (переносится между сезонами);
- на свободные места (capacity − sold_as_of) — для матчей в продаже (учитывает потолок зоны).

Допущение: sales_open известна заранее для всех матчей (как seats в zones.csv).
"""
import numpy as np
import pandas as pd
from catboost import CatBoostRegressor

from src.data import load_daily_sales, load_raw

AVANGARD = "HC Avangard Omsk"
SNAP_DAYS = (1, 3, 5, 7, 9, 11, 13, 15, 18, 21, 25, 29, 33, 40)  # дней до матча на дату среза
LEVEL_N = 10  # окно уровня зоны: текущий сезон, добор до N матчей из прошлых (как smart baseline)
MAX_H = 70    # кривая продаж считается для h = 0..MAX_H дней до матча
KEY = ["match_id", "zone"]

FEATURES = [
    "zone", "weekday", "hour", "month", "capacity", "season_tickets",
    "lvl_season", "lvl_last5",
    "sold_as_of", "sold_last_7d", "sold_last_14d", "frac_elapsed", "days_left",
    "pace_est", "sold_lvl", "sold_cap", "lvl_cap", "match_sold_ratio",
    "opp_rel", "opp_n",
    "opp_winrate", "opp_gd", "avg_winrate", "avg_form10", "avg_gd10",
]

# bootstrap_type="No" и random_strength=0 делают обучение детерминированным: на 3 фолдах
# разброс RMSE между seed без них был ±3, что сравнимо с выигрышем от признаков.
CB_PARAMS = dict(iterations=800, learning_rate=0.03, depth=4, loss_function="RMSE",
                 bootstrap_type="No", random_strength=0,
                 random_state=42, verbose=0, allow_writing_files=False)


# --- История: уровни зон, соперник, кривая продаж -----------------------------------------

def _played_info(train):
    """Сыгранные матчи по зонам; rel — итог / среднее LEVEL_N предыдущих матчей в этой зоне."""
    t = train[["match_id", "match_date", "season", "opponent", "zone", "tickets_total"]]
    t = t.sort_values(["match_date", "match_id", "zone"]).reset_index(drop=True)
    pre_lvl = (t.groupby("zone")["tickets_total"]
               .transform(lambda s: s.shift(1).rolling(LEVEL_N, min_periods=3).mean()))
    return t.assign(rel=t["tickets_total"] / pre_lvl)


def _match_fracs(daily, train):
    """Доля итога матча в зоне, проданная строго раньше, чем за h дней до матча.

    DataFrame с индексом (match_id, zone) и колонками h = 0..MAX_H.
    """
    per = daily.pivot_table(index=KEY, columns="days_before", values="tickets",
                            aggfunc="sum", fill_value=0)
    per = per.reindex(columns=range(MAX_H + 2), fill_value=0)
    after = per.iloc[:, ::-1].cumsum(axis=1).iloc[:, ::-1]  # продано с days_before >= k
    tot = train.set_index(KEY)["tickets_total"].reindex(after.index)
    frac = after.loc[:, 1:MAX_H + 1].div(tot.where(tot > 0), axis=0).clip(0, 1)
    frac.columns = range(MAX_H + 1)  # колонка k+1 (days_before > h) -> h
    return frac


def _history(played, fracs, rows):
    """Признаки уровня зоны, соперника и кривой продаж на snap_date каждой строки.

    Используются только матчи с match_date < snap_date.
    """
    cols = ["lvl_season", "lvl_last5", "opp_rel", "opp_n", "curve"]
    res = pd.DataFrame(np.nan, index=rows.index, columns=cols)
    for d, idx in rows.groupby("snap_date").groups.items():
        r = rows.loc[idx]
        past = played[played["match_date"] < d]
        if past.empty:
            continue
        order = past.drop_duplicates("match_id")[["match_id", "season"]]

        last5 = past[past["match_id"].isin(order["match_id"].tail(5))]
        res.loc[idx, "lvl_last5"] = r["zone"].map(last5.groupby("zone")["tickets_total"].mean())
        for season, sidx in r.groupby("season").groups.items():
            cur = order.loc[order["season"] == season, "match_id"]
            prev = order.loc[order["season"] != season, "match_id"]
            ids = pd.concat([prev.tail(max(LEVEL_N - len(cur), 0)), cur])
            lvl = past[past["match_id"].isin(ids)].groupby("zone")["tickets_total"].mean()
            res.loc[sidx, "lvl_season"] = r.loc[sidx, "zone"].map(lvl)

        opp = past.dropna(subset=["rel"]).groupby(["opponent", "zone"])["rel"].agg(["mean", "size"])
        key = pd.MultiIndex.from_frame(r[["opponent", "zone"]])
        res.loc[idx, "opp_rel"] = opp["mean"].reindex(key).to_numpy()
        res.loc[idx, "opp_n"] = opp["size"].reindex(key).fillna(0).to_numpy()

        f = fracs[fracs.index.get_level_values("match_id").isin(order["match_id"])]
        curve = f.groupby(level="zone").mean()  # зона × h
        h = r["days_left"].clip(upper=MAX_H).astype(int).to_numpy()
        z = r["zone"].to_numpy()
        res.loc[idx, "curve"] = curve.reindex(index=z).to_numpy()[np.arange(len(r)), h]
    return res


# --- КХЛ ----------------------------------------------------------------------------------

def _khl_long(khl):
    """Регулярный чемпионат в длинном виде: команда, дата, сезон, разница шайб, победа."""
    k = khl[khl["stage"].fillna("Regular season") == "Regular season"]
    sides = [(k["home"], k["home_goals"] - k["away_goals"]), (k["away"], k["away_goals"] - k["home_goals"])]
    g = pd.concat([pd.DataFrame({"team": team, "date": k["date"], "season": k["season"], "gd": gd})
                   for team, gd in sides], ignore_index=True).sort_values("date")
    return g.assign(win=(g["gd"] > 0).astype(float))


def _khl_features(games, rows):
    """Сила соперника и форма Авангарда на snap_date (только игры с date < snap_date).

    Сезонная статистика — по сезону строки; если в нём меньше 5 игр команды, берём прошлый сезон.
    """
    seasons = sorted(games["season"].unique())
    cols = ["opp_winrate", "opp_gd", "avg_winrate", "avg_form10", "avg_gd10"]
    res = pd.DataFrame(np.nan, index=rows.index, columns=cols)
    for d, idx in rows.groupby("snap_date").groups.items():
        r = rows.loc[idx]
        past = games[games["date"] < d]

        def season_stats(team, season):
            g = past[(past["team"] == team) & (past["season"] == season)]
            if len(g) < 5 and season in seasons and seasons.index(season) > 0:
                g = past[(past["team"] == team) & (past["season"] == seasons[seasons.index(season) - 1])]
            return (g["win"].mean(), g["gd"].mean()) if len(g) else (np.nan, np.nan)

        for (team, season), gidx in r.groupby(["opponent_khl", "season"]).groups.items():
            res.loc[gidx, ["opp_winrate", "opp_gd"]] = season_stats(team, season)
        for season, sidx in r.groupby("season").groups.items():
            res.loc[sidx, "avg_winrate"] = season_stats(AVANGARD, season)[0]
        av = past[past["team"] == AVANGARD].tail(10)
        if len(av):
            res.loc[idx, ["avg_form10", "avg_gd10"]] = av["win"].mean(), av["gd"].mean()
    return res


# --- Строки для обучения и прогноза -------------------------------------------------------

def _window(df):
    return (df["match_date"] - df["sales_open"]).dt.days


def _train_snapshots(train, daily):
    """Сыгранные матчи -> срезы «за h дней до матча» + срез «продажи не открыты» (days_left = окно)."""
    base = train.drop(columns=["sold_as_of", "sold_last_7d", "sold_last_14d",
                               "sales_started", "days_to_match", "days_since_open"])
    base = base.assign(window=_window(base))
    parts = []
    for h in SNAP_DAYS + (None,):
        s = base.copy()
        s["days_left"] = s["window"] if h is None else np.minimum(h, s["window"])
        d = daily.merge(s[KEY + ["days_left"]], on=KEY)
        before = d["days_before"] > d["days_left"]  # date < дата среза
        for name, span in [("sold_as_of", None), ("sold_last_7d", 7), ("sold_last_14d", 14)]:
            m = before if span is None else before & (d["days_before"] <= d["days_left"] + span)
            agg = d[m].groupby(KEY)["tickets"].sum()
            s[name] = agg.reindex(pd.MultiIndex.from_frame(s[KEY]), fill_value=0).to_numpy()
        s["snap_date"] = s["match_date"] - pd.to_timedelta(s["days_left"], unit="D")
        parts.append(s)
    # Срезы с h >= окна совпадают со срезом «не открыты» — оставляем один
    return pd.concat(parts, ignore_index=True).drop_duplicates(KEY + ["days_left"]).reset_index(drop=True)


def _val_rows(val, as_of):
    v = val.assign(window=_window(val), snap_date=as_of)
    v["days_left"] = np.minimum(v["days_to_match"], v["window"])
    return v


def _add_features(rows, played, fracs, games):
    rows = rows.copy()
    rows["frac_elapsed"] = (1 - rows["days_left"] / rows["window"]).clip(0, 1)
    rows = rows.join(_history(played, fracs, rows)).join(_khl_features(games, rows))

    open_ = rows["frac_elapsed"] > 0
    c = rows["curve"]
    pace = (rows["sold_as_of"] / c.clip(lower=0.02)).where(open_ & (c > 0.02))
    rows["pace_est"] = (pace / rows["lvl_season"]).clip(upper=3)
    rows["sold_lvl"] = rows["sold_as_of"] / rows["lvl_season"]
    rows["sold_cap"] = rows["sold_as_of"] / rows["capacity"]
    rows["lvl_cap"] = rows["lvl_season"] / rows["capacity"]
    # Продажи всего матча относительно ожидаемых к этому дню (уровень × типичная доля)
    g = [rows["match_id"], rows["snap_date"]]
    msold = rows["sold_as_of"].groupby(g).transform("sum")
    mexp = (rows["lvl_season"] * c.fillna(0)).groupby(g).transform("sum")
    rows["match_sold_ratio"] = (msold / mexp).where(open_ & (mexp > 0))
    return rows


def _fit_rest(X_tr, w, X_val, scale_tr, scale_val, features):
    """CatBoost на остаток продаж в долях scale; возвращает остаток в билетах для X_val."""
    y = (X_tr["tickets_total"] - X_tr["sold_as_of"]) / scale_tr
    model = CatBoostRegressor(**CB_PARAMS)
    model.fit(X_tr[features], y, sample_weight=w)
    return np.maximum(model.predict(X_val[features]), 0) * scale_val


def catboost_predict(train, val, as_of, features=FEATURES):
    """Прогноз итога tickets_total для строк val по моделям на срезах train-матчей.

    features — список признаков (по умолчанию FEATURES); нужен для абляций в 03_model.ipynb.
    """
    as_of = pd.Timestamp(as_of)
    khl = load_raw()["khl"]
    games = _khl_long(khl[khl["date"] < as_of])
    daily = load_daily_sales(as_of)
    daily = daily[daily["match_id"].isin(train["match_id"])]
    played = _played_info(train)
    fracs = _match_fracs(daily, train)

    X_tr = _add_features(_train_snapshots(train, daily), played, fracs, games)
    X_tr = X_tr[X_tr["lvl_season"].notna()].reset_index(drop=True)  # первые матчи без истории
    X_val = _add_features(_val_rows(val, as_of), played, fracs, games)

    # Каждый матч весит одинаково: срез «не открыты» весит как все открытые срезы матча вместе,
    # потому что на дату прогноза большинство матчей ещё не в продаже.
    closed = X_tr["frac_elapsed"] == 0
    n_open = (~closed).groupby([X_tr["match_id"], X_tr["zone"]]).transform("sum").clip(lower=1)
    w = np.where(closed, 1.0, 1.0 / n_open)

    by_lvl = _fit_rest(X_tr, w, X_val, X_tr["lvl_season"], X_val["lvl_season"], features)
    free_tr = (X_tr["capacity"] - X_tr["sold_as_of"]).clip(lower=1)
    free_val = (X_val["capacity"] - X_val["sold_as_of"]).clip(lower=1)
    by_free = _fit_rest(X_tr, w, X_val, free_tr, free_val, features)

    in_sale = X_val["frac_elapsed"].to_numpy() > 0
    return X_val["sold_as_of"].to_numpy() + np.where(in_sale, by_free, by_lvl)
