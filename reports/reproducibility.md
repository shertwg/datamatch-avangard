# Воспроизводимость: прогон в чистом окружении

## Итог

`notebooks/solution.ipynb`, запущенный с нуля в новом venv, выдаёт `predictions.csv`, побайтово
совпадающий с файлом в репозитории: 136 строк, сумма прогноза 85 631.

## Прогон 04.10.2026

| | |
|---|---|
| Дата | 04.10.2026 |
| Коммит | `f2b69f8` (Merge pull request #10 from shertwg/eda) |
| ОС | Windows 11 |
| Python | 3.12.10, новый venv |
| Установка | `pip install -r requirements.txt` — без ошибок, ~5 мин |
| Запуск | `solution.ipynb` через nbconvert в активированном venv, ~30 с |
| Результат | `predictions.csv` совпал побайтово с файлом в репозитории |
| sha256 | `a3c1f128f30005929748e33f67e0b505869de86add7254986dbf4052cd55a3da` |

## Шаги (PowerShell, из корня репозитория)

```powershell
py -3.12 -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
jupyter nbconvert --to notebook --execute notebooks/solution.ipynb --output solution_run.ipynb
Get-FileHash predictions.csv -Algorithm SHA256
```

`.venv/` уже в `.gitignore`. `solution_run.ipynb` — выполненная копия, коммитить её не нужно.

## Что обеспечивает воспроизводимость

- Версии всех библиотек закреплены в `requirements.txt`.
- `random_state=42` во всех моделях (`src/model.py`).
- Пути относительные, данные читаются только из `data/` через `src/data.py`.
- Ключевые цифры отчёта проверяет assert в `notebooks/04_report.ipynb → Выводы для слайдов`.
- Утечку проверяет `python -m tests.test_leakage` (см. ответы 1–2 в `reports/jury_qa.md`).

## Примечания

- **sha256 зависит от концов строк.** `.gitattributes` задаёт `* text=auto`: в git `predictions.csv`
  хранится с LF, а на Windows лежит в рабочей копии и записывается pandas с CRLF. Хэш выше — для
  Windows-файла с CRLF. Файл из git с LF (`git show HEAD:predictions.csv`) имеет sha256
  `8fe210746609ac83b97ccb5130ca54343d58caeb7f190b17b19b66015df9e3b1`. Содержимое одинаковое.
  На Linux/macOS сверять нужно с этим LF-хэшем.
- **Что изменилось после `f2b69f8`.** `src/model.py`: КХЛ теперь загружается через `load_khl(as_of)` —
  тот же фильтр `date < as_of`, другой путь в коде. `requirements.txt`: добавлены `ipykernel==7.4.0` и
  `nbconvert==7.17.1`, раньше они ставились зависимостями `jupyter`. Повторный прогон
  `predict_test(catboost_predict)` на текущем коде даёт тот же `predictions.csv` до билета.
  Полный прогон в чистом venv на текущем HEAD не делали.
