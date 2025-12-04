import pandas as pd
from scipy.stats import shapiro, probplot
import matplotlib.pyplot as plt
import seaborn as sns
import sys

def check_normality_shapiro(file_path, sheet_name, column_name):
   
    try:
        #Чтение данных из Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"Ошибка: Файл '{file_path}' не найден.")
        sys.exit(1)
    except ValueError as e:
        print(f"Ошибка: Не удалось прочитать данные. Проверьте имя листа или формат файла. {e}")
        sys.exit(1)
    
    #Проверяем, существует ли столбец
    if column_name not in df.columns:
        print(f"Ошибка: Столбец '{column_name}' не найден в файле.")
        sys.exit(1)
    
    #Извлечение ряда данных
    data = df[column_name].dropna()
    
    #Проведение теста Шапиро-Уилка
    if len(data) < 3 or len(data) > 5000:
        print("Предупреждение: Тест Шапиро-Уилка лучше всего подходит для выборок размером от 3 до 5000.")
        if len(data) < 3:
            print(f"Недостаточно данных (n={len(data)}) для проведения теста.")
            return
        
    stat, p_value = shapiro(data)
    
    #Вывод результата
    alpha = 0.05
    print(f"Результаты теста Шапиро-Уилка для столбца '{column_name}':")
    print(f"Статистика теста: {stat:.4f}")
    print(f"p-value: {p_value:.4f}")
    if p_value > alpha:
        print("--> Вывод: p-value > 0.05. Распределение, вероятно, нормальное.")
    else:
        print("--> Вывод: p-value <= 0.05. Распределение, вероятно, ненормальное.")
    
    #Визуализация распределения
    plt.figure(figsize=(14, 5))
    
    #Гистограмма с KDE
    plt.subplot(1, 2, 1)
    sns.histplot(data, kde=True, bins=15, color='skyblue')
    plt.title(f"Распределение данных: {column_name}")
    plt.xlabel(column_name)
    plt.ylabel("Частота")
    plt.grid(True, alpha=0.3)
    
    #Q-Q график
    plt.subplot(1, 2, 2)
    probplot(data, dist="norm", plot=plt)
    plt.title("Q-Q график")
    
    plt.tight_layout()
    plt.show()

#Пример использования
if __name__ == "__main__":
    excel_file = r'C:\Users\HYPERPC\Documents\Научная статья\Расчеты.xlsx'  # Замени на имя своего файла
    sheet_name = 'Лист1' # Замени на имя своего листа
    data_column = 'Темпы изменения выдач кредитов физических лицам г/г' # Замени на имя своего столбца
    
    check_normality_shapiro(excel_file, sheet_name, data_column)
