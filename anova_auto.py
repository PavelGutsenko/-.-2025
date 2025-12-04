import math
import numpy as np
import pandas as pd
from scipy import stats
import os
import matplotlib.pyplot as plt
import seaborn as sns

# Tukey HSD (опционально)
try:
    from statsmodels.stats.multicomp import pairwise_tukeyhsd
    HAS_SM = True
except Exception:
    HAS_SM = False

ORDERED_SIGNAL = ["Мягкий", "Нейтральный", "Жесткий"]
ALPHA = 0.05


# === Вспомогательные функции ===
def norm_signal_value(s: str):
    """Нормализует текстовые обозначения сигналов"""
    if pd.isna(s):
        return None
    t = str(s).strip().lower()
    if t.startswith("мяг"):
        return "Мягкий"
    if t.startswith("жест") or t.startswith("жёст"):
        return "Жесткий"
    if t.startswith("нейтр"):
        return "Нейтральный"
    return None


def anova_oneway_manual(groups):
    """Ручной расчёт однофакторной ANOVA"""
    groups = [g[~np.isnan(g)] for g in groups]
    groups = [g for g in groups if g.size > 0]
    if len(groups) < 2:
        return (np.nan, np.nan, {})

    k = len(groups)
    ns = np.array([g.size for g in groups], dtype=float)
    means = np.array([g.mean() for g in groups], dtype=float)
    overall_mean = np.average(means, weights=ns)

    ss_between = float(np.sum(ns * (means - overall_mean) ** 2))
    ss_within = float(np.sum([np.sum((g - m) ** 2) for g, m in zip(groups, means)]))
    ss_total = ss_between + ss_within

    df_between = k - 1
    df_within = int(np.sum(ns) - k)

    ms_between = ss_between / df_between if df_between > 0 else np.nan
    ms_within = ss_within / df_within if df_within > 0 else np.nan

    F = ms_between / ms_within if ms_within > 0 else np.nan
    p = 1 - stats.f.cdf(F, df_between, df_within) if not math.isnan(F) else np.nan

    return F, p, {
        "ssb": ss_between, "ssw": ss_within, "sst": ss_total,
        "dfb": df_between, "dfw": df_within,
        "msb": ms_between, "msw": ms_within
    }


def effect_sizes(ssb, sst, dfb, msw):
    """Расчёт эффектов η² и ω²"""
    if any(math.isnan(x) for x in [ssb, sst, msw]) or sst <= 0:
        return (np.nan, np.nan)
    eta2 = ssb / sst
    omega2 = (ssb - dfb * msw) / (sst + msw) if (sst + msw) > 0 else np.nan
    return (eta2, omega2)


def interpret_effect(eta2):
    """Интерпретация силы эффекта"""
    if math.isnan(eta2):
        return "н/д"
    if eta2 < 0.01:
        return "пренебрежимо слабая"
    elif eta2 < 0.06:
        return "слабая"
    elif eta2 < 0.14:
        return "средняя"
    else:
        return "сильная"


def run_tukey(series, groups):
    """Пост-хок анализ Tukey HSD"""
    if not HAS_SM:
        return None
    try:
        res = pairwise_tukeyhsd(endog=series.values, groups=groups.values, alpha=ALPHA)
        return pd.DataFrame(data=res.summary().data[1:], columns=res.summary().data[0])
    except Exception:
        return None


# === Основная функция ===
def main():
    print("=== ANOVA-анализ с визуализацией ===\n")
    file_path = input("Введите путь к Excel-файлу: ").strip('" ')
    sheet_name = input("Введите имя листа (Enter — первый лист): ").strip()

    df = pd.read_excel(file_path, sheet_name=sheet_name if sheet_name else 0)

    print("\nКолонки в файле:")
    for i, c in enumerate(df.columns):
        print(f"{i+1}. {c}")

    signal_col = input("\nВведите название колонки с сигналом: ").strip()
    indicators_input = input("Введите названия числовых колонок через запятую: ").strip()
    indicator_cols = [c.strip() for c in indicators_input.split(",")]

    # Нормализуем сигнал
    df["_signal_norm"] = df[signal_col].apply(norm_signal_value)
    df = df[~df["_signal_norm"].isna()].copy()
    df["_signal_norm"] = pd.Categorical(df["_signal_norm"], categories=ORDERED_SIGNAL, ordered=True)

    for col in indicator_cols:
        groups = []
        labels = []
        for lab in ORDERED_SIGNAL:
            g = df.loc[df["_signal_norm"] == lab, col].dropna().astype(float).values
            if g.size > 0:
                groups.append(g)
                labels.append(lab)
        if len(groups) < 2:
            print(f"[Пропуск] Недостаточно данных для '{col}'")
            continue

        F, p, extra = anova_oneway_manual(groups)
        eta2, omega2 = effect_sizes(extra["ssb"], extra["sst"], extra["dfb"], extra["msw"])
        strength = interpret_effect(eta2)

        print("\n=== Показатель:", col, "===")
        print(f"F = {F:.4f}, p = {p:.4f}")
        if p < ALPHA:
            print("➡ Есть статистически значимые различия (сигнал влияет).")
        else:
            print("➡ Нет статистически значимых различий.")
        print(f"eta² = {eta2:.4f}, omega² = {omega2:.4f}, сила связи: {strength}")

        tukey_df = run_tukey(df[col].dropna().astype(float), df["_signal_norm"])
        if tukey_df is not None:
            print("\n--- Tukey HSD ---")
            print(tukey_df.to_string(index=False))

        # === Визуализация (теперь выводим на экран, не сохраняем) ===
        fig, axes = plt.subplots(1, 2, figsize=(10, 4))
        sns.boxplot(x="_signal_norm", y=col, data=df, ax=axes[0], palette="pastel")
        axes[0].set_title(f"Boxplot: {col}")
        axes[0].set_xlabel("Тип сигнала")
        axes[0].set_ylabel(col)

        sns.pointplot(x="_signal_norm", y=col, data=df, ax=axes[1], ci=95, join=False, color="steelblue")
        axes[1].set_title(f"Средние значения (95% ДИ): {col}")
        axes[1].set_xlabel("Тип сигнала")
        axes[1].set_ylabel(col)

        plt.suptitle(f"ANOVA для {col}\nF={F:.2f}, p={p:.3f}, η²={eta2:.3f}", fontsize=10)
        plt.tight_layout()

        # Вместо сохранения — показать график пользователю
        plt.show()

    print("\n✅ Анализ завершён.")


if __name__ == "__main__":
    main()
