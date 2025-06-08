import pandas as pd

# Создаем простой DataFrame
data = {"Английский": ["Hello", "Goodbye"], "Русский": ["Привет", "Пока"]}
df = pd.DataFrame(data)

# Сохраняем в Excel
df.to_excel("test_pandas.xlsx", index=False)

print("Pandas работает! Файл test_pandas.xlsx создан.")