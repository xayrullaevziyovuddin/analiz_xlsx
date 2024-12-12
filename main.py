import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from tkinter import Tk
from tkinter.filedialog import askopenfilename


def analyze_shipments(file_path):
    try:
        # Считываем данные
        data = pd.read_excel(file_path, sheet_name='shipments')

        # Проверка наличия необходимых столбцов
        if 'Место доставки' not in data.columns or 'Номер заказа' not in data.columns or 'Стоимость' not in data.columns:
            print("Ошибка: необходимые столбцы не найдены.")
            return

        # Группируем данные по регионам
        summary = data.groupby('Место доставки').agg({
            'Номер заказа': 'count',
            'Стоимость': 'mean'
        }).rename(columns={
            'Номер заказа': 'Количество заказов',
            'Стоимость': 'Средняя стоимость доставки'
        }).sort_values(by='Количество заказов', ascending=False)

        # Печать результатов анализа
        print("\nАнализ данных:\n")
        print(summary)

        # Визуализация
        sns.set(style="whitegrid")

        # Столбчатая диаграмма: Количество заказов
        plt.figure(figsize=(10, 6))
        sns.barplot(x=summary.index, y=summary['Количество заказов'], palette="Blues_d")
        plt.title("Количество заказов по регионам")
        plt.xlabel("Регионы")
        plt.ylabel("Количество заказов")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

        # Столбчатая диаграмма: Средняя стоимость доставки
        plt.figure(figsize=(10, 6))
        sns.barplot(x=summary.index, y=summary['Средняя стоимость доставки'], palette="Reds_d")
        plt.title("Средняя стоимость доставки по регионам")
        plt.xlabel("Регионы")
        plt.ylabel("Средняя стоимость доставки")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.show()

    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")


def get_file_path():
    Tk().withdraw()  # Скрыть основное окно Tkinter
    file_path = askopenfilename(title="Выберите файл Excel", filetypes=[("Excel Files", "*.xlsx")])
    Tk().destroy()  # Закрыть окно Tkinter после выбора файла
    return file_path


# Вызов функции анализа
file_path = get_file_path()
if file_path:
    analyze_shipments(file_path)
else:
    print("Файл не выбран.")
