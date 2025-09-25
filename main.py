import pandas as pd

# Загружаем файл, пропуская первые две строки
file_path = "_fmov2gsm.xlsx"
df = pd.read_excel(file_path, sheet_name="Снимок экрана", header=1, skiprows=[0])

# Группируем по коду и наименованию ТМЦ, суммируем количество
result = (
    df.groupby(["Код ТМЦ", "Наименование"], as_index=False)["Кол-во"]
      .sum()
)

# Сохраняем результат в новый Excel
result.to_excel("ТМЦ_сумма.xlsx", index=False)

print("Файл сохранен как ТМЦ_сумма.xlsx")