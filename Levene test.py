import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import levene

# === 1. Ввод пути и чтение файла ===
file_path = input("Введите путь к Excel-файлу: ").strip()
try:
    df = pd.read_excel(file_path)
except Exception as e:
    print(f"Ошибка при чтении файла: {e}")
    sys.exit(1)

# === 2. Выбор столбцов ===
print("\nСписок доступных столбцов:")
for col in df.columns:
    print(f"- {col}")

dep_var = input("\nВведите НАЗВАНИЕ зависимой переменной (числовой): ").strip()
group_var = input("Введите НАЗВАНИЕ группирующей переменной (текстовой): ").strip()

if dep_var not in df.columns or group_var not in df.columns:
    print("❌ Ошибка: одно из названий столбцов указано неверно.")
    sys.exit(1)

# === 3. Очистка и подготовка данных ===
df = df[[dep_var, group_var]].dropna()
df[dep_var] = pd.to_numeric(df[dep_var], errors='coerce')
df = df.dropna(subset=[dep_var])

groups = df[group_var].unique()
samples = [df.loc[df[group_var] == g, dep_var] for g in groups if len(df.loc[df[group_var] == g, dep_var]) > 1]

if len(samples) < 2:
    print("❌ Недостаточно групп с числовыми значениями для теста.")
    sys.exit(1)

print(f"\nНайдено {len(samples)} групп: {', '.join(map(str, groups))}")

# === 4. Проведение теста Левене ===
stat, p_value = levene(*samples)

print("\n=== РЕЗУЛЬТАТ ТЕСТА ЛЕВЕНЕ ===")
print(f"Статистика: {stat:.4f}")
print(f"P-значение: {p_value:.6f}")

alpha = 0.05
if p_value > alpha:
    print(f"✅ P-value = {p_value:.6f} > {alpha}. Различия дисперсий между группами статистически незначимы — дисперсии можно считать равными.")
else:
    print(f"⚠️ P-value = {p_value:.6f} ≤ {alpha}. Различия дисперсий между группами статистически значимы — дисперсии неоднородны.")

# === 5. Визуализация ===
plt.figure(figsize=(10, 6))
sns.boxplot(x=group_var, y=dep_var, data=df)
plt.title(f"Распределение '{dep_var}' по группам '{group_var}' (тест Левене)")
plt.xlabel(group_var)
plt.ylabel(dep_var)
plt.xticks(rotation=30)
plt.grid(alpha=0.3)
plt.tight_layout()
plt.show()