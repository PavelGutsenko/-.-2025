import pandas as pd
from scipy.stats import shapiro
import sys

def check_normality_shapiro(file_path, sheet_name, column_name):

    Аргументы:
    file_path (str): Путь к Excel-файлу.
    sheet_name (str): Имя листа в файле.
    column_name (str): Имя столбца для проверки.
    """
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
    data = df[column_name].dropna() # Убираем пропущенные значения, если они есть
    
    #Проведение теста Шапиро-Уилка
    # Тест не работает с выборками меньше 3 или больше 5000.
    if len(data) < 3 or len(data) > 5000:
        print("Предупреждение: Тест Шапиро-Уилка лучше всего подходит для выборок размером от 3 до 5000.")
        if len(data) < 3:
            print(f"Недостаточно данных (n={len(data)}) для проведения теста.")
            return
        
    #Запускаем тест
    stat, p_value = shapiro(data)
    
    #Вывод результата
    alpha = 0.05 # Уровень значимости
    
    print(f"Результаты теста Шапиро-Уилка для столбца '{column_name}':")
    print(f"Статистика теста: {stat:.4f}")
    print(f"p-value: {p_value:.4f}")
    
    if p_value > alpha:
        print("--> Вывод: p-value > 0.05. Распределение, вероятно, нормальное.")
    else:
        print("--> Вывод: p-value <= 0.05. Распределение, вероятно, ненормальное.")

#Пример использования
if __name__ == "__main__":
    excel_file = r'c:\Users\HYPERPC\Documents\Научная статья\Инфляция и ключевая ставка Банка России_F01_01_2024_T30_07_2025.xlsx'  # Замени на имя своего файла
    sheet_name = 'Расчеты' # Замени на имя своего листа
    data_column = 'Индекс изменения M2' # Замени на имя своего столбца
    

    check_normality_shapiro(excel_file, sheet_name, data_column)
