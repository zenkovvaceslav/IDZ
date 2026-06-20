import pandas as pd
from pathlib import Path


def detect_column(df, candidates, role, required=True):
    """Найти столбец таблицы по списку допустимых названий."""
    normalized = {str(col).strip().lower(): col for col in df.columns}
    for name in candidates:
        col = normalized.get(name.strip().lower())
        if col is not None:
            return col
    if not required:
        return None
    raise KeyError(f"Не найден столбец для «{role}». Ожидался один из: {', '.join(candidates)}. Есть: {', '.join(map(str, df.columns))}.")


def normalize_name(value):
    """Привести ФИО к единому виду: убрать лишние пробелы и заменить ё на е."""
    return " ".join(str(value).split()).replace("ё", "е").replace("Ё", "Е")


def is_teacher(name):
    """Отсеять не-ФИО: числа, дипломы, комиссии и прочие служебные строки."""
    text = str(name).strip().lower()
    if not text:
        return False
    if not any("а" <= ch <= "я" or ch == "ё" for ch in text):
        return False
    return not any(word in text for word in ("диплом", "комиссия", "отчисл", "отчсл", "компенс"))


def status_of(underload):
    """Вернуть текстовый статус по величине недогруза."""
    if underload > 0:
        return "недогружен"
    if underload < 0:
        return "перегружен"
    return "норма"


def rates_to_map(rates):
    """Преобразовать список ставок в словарь {ФИО: ставка}."""
    if rates is None:
        return {}
    items = rates.items() if isinstance(rates, dict) else zip(rates.iloc[:, 0], rates.iloc[:, 1])
    result = {}
    for name, rate in items:
        value = pd.to_numeric(pd.Series([rate]), errors="coerce").iloc[0]
        if str(name).strip() and pd.notna(value):
            result[normalize_name(name)] = float(value)
    return result


def compute_underload(load_df, rates=None, default_rate=1.0, norm_per_rate=830, fio_col=None, hours_col=None, rate_col=None):
    """Посчитать недогруз каждого преподавателя. Возвращает таблицу: Преподаватель, Назначено, Ставки, Норма, Недогруз, Статус."""
    fio_candidates = ("ФИО", "Преподаватель", "Названия строк", "Фамилия")
    hours_candidates = ("Нагрузка", "всего", "ч.Итого", "Назначено", "Часы")
    rate_candidates = ("Ставка", "Ставки", "в приказ", "Доля ставки")
    result_columns = ["Преподаватель", "Назначено", "Ставки", "Норма", "Недогруз", "Статус"]

    if load_df is None or len(load_df) == 0:
        return pd.DataFrame(columns=result_columns)

    fio_col = fio_col or detect_column(load_df, fio_candidates, "ФИО преподавателя")
    hours_col = hours_col or detect_column(load_df, hours_candidates, "часы нагрузки")
    rate_col = rate_col or detect_column(load_df, rate_candidates, "ставка", required=False)

    columns = [fio_col, hours_col] + ([rate_col] if rate_col else [])
    work = load_df[columns].copy()
    work["__fio"] = work[fio_col].map(normalize_name)
    work[hours_col] = pd.to_numeric(work[hours_col], errors="coerce").fillna(0)
    work = work[work["__fio"].map(is_teacher)]

    aggregation = {"Назначено": (hours_col, "sum")}
    if rate_col:
        work[rate_col] = pd.to_numeric(work[rate_col], errors="coerce")
        aggregation["Ставки"] = (rate_col, "max")
    result = work.groupby("__fio", as_index=False).agg(**aggregation)
    result = result.rename(columns={"__fio": "Преподаватель"})

    if "Ставки" not in result.columns:
        result["Ставки"] = float("nan")
    rate_map = rates_to_map(rates)
    if rate_map:
        result["Ставки"] = result["Ставки"].fillna(result["Преподаватель"].map(rate_map))
    result["Ставки"] = result["Ставки"].fillna(default_rate)

    result["Норма"] = result["Ставки"] * norm_per_rate
    result["Недогруз"] = result["Норма"] - result["Назначено"]
    result["Статус"] = result["Недогруз"].apply(status_of)
    return result[result_columns]


def total_underload(result):
    """Суммарный недогруз по тем, у кого он положительный."""
    return result.loc[result["Недогруз"] > 0, "Недогруз"].sum()


def main():
    BASE_DIR = Path(__file__).resolve().parent
    DATA_DIR = BASE_DIR / "data"
    LOAD_FILE = DATA_DIR / "нагрузка- вар2-ТиСАПРМП-2023-2024.xlsx"
    LOAD_SHEET = "общее"
    LOAD_HEADER = 0
    RATES_FILE = LOAD_FILE
    RATES_SHEET = "Лист2"
    RATES_HEADER = 3
    RATES_FIO_COL = "ФИО"
    RATES_RATE_COL = "новая"
    OUTPUT_FILE = DATA_DIR / "недогруз.xlsx"
    DEFAULT_RATE = 1.0

    load_df = pd.read_excel(LOAD_FILE, sheet_name=LOAD_SHEET, header=LOAD_HEADER)
    rates_df = pd.read_excel(RATES_FILE, sheet_name=RATES_SHEET, header=RATES_HEADER)
    rates = rates_df[[RATES_FIO_COL, RATES_RATE_COL]]

    report = compute_underload(load_df, rates=rates, default_rate=DEFAULT_RATE)

    rate_map = rates_to_map(rates)
    covered = report["Преподаватель"].isin(set(rate_map)).sum()
    print(f"Преподавателей: {len(report)}; ставка из списка найдена для {covered}, остальным проставлена {DEFAULT_RATE}.")
    print(report.to_string(index=False))
    print("\nСуммарный недогруз:", total_underload(report))
    report.to_excel(OUTPUT_FILE, index=False)
    print(f"Результат сохранён в {OUTPUT_FILE}")


if __name__ == "__main__":
    main()