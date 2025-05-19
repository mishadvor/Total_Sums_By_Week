import pandas as pd
import streamlit as st
from io import BytesIO
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils.dataframe import dataframe_to_rows

# Заголовок приложения
st.title("📊 Обработка ОБЩИХ финансовых отчётов Wildberries (Разбивка по Датам. С графиком.)")

# Выбор режима
mode = st.radio(
    "Выберите режим обработки", ["Один файл", "Три файла (три точки продаж)"]
)

# Поле для выбора даты начала обработки
start_date = st.date_input(
    "Выберите дату начала обработки данных",
    value=pd.to_datetime("today").date(),  # значение по умолчанию — сегодня
)

# Преобразуем start_date в Timestamp
start_date = pd.to_datetime(start_date)

##########------ ЗАГРУЗКА ФАЙЛОВ ------------
dfs = []

if mode == "Один файл":
    uploaded_file = st.file_uploader(
        "Загрузите Excel-файл отчёта Wildberries", type=["xlsx", "xls"]
    )
    if uploaded_file:
        dfs.append(pd.read_excel(uploaded_file))

elif mode == "Три файла (три точки продаж)":
    st.markdown("### Загрузите 3 файла по очереди:")
    file1 = st.file_uploader("Файл для точки 1", type=["xlsx", "xls"], key="f1")
    file2 = st.file_uploader("Файл для точки 2", type=["xlsx", "xls"], key="f2")
    file3 = st.file_uploader("Файл для точки 3", type=["xlsx", "xls"], key="f3")

    if file1 and file2 and file3:
        dfs = [pd.read_excel(file1), pd.read_excel(file2), pd.read_excel(file3)]
        for i, df in enumerate(dfs):
            df["Источник"] = f"Точка {i+1}"
    else:
        st.warning("Пожалуйста, загрузите все три файла.")

if len(dfs) > 0:
    try:
        # Объединяем все датафреймы в один
        combined_df = pd.concat(dfs, ignore_index=True)

        ##########------ ЧИСТКА И ПРЕОБРАЗОВАНИЕ ДАТЫ ------------
        # Удаляем всё, что после "T"
        combined_df["Дата конца"] = (
            combined_df["Дата конца"].astype(str).str.split("T").str[0]
        )

        # Преобразуем в datetime
        combined_df["Дата конца"] = pd.to_datetime(
            combined_df["Дата конца"], errors="coerce"
        )

        # Форматируем как строка YYYY-MM-DD
        combined_df["Дата конца"] = combined_df["Дата конца"].dt.strftime("%Y-%m-%d")

        # Теперь можно сравнивать с выбранной пользователем датой
        combined_df["Дата конца"] = pd.to_datetime(combined_df["Дата конца"])
        filtered_df = combined_df[
            combined_df["Дата конца"] >= pd.to_datetime(start_date)
        ]

        ##########------ НАЧАЛО ОБРАБОТКИ ДАННЫХ ------------
        sums_per_date = (
            filtered_df.groupby("Дата конца")
            .agg(
                {
                    "Продажа": "sum",
                    "К перечислению за товар": "sum",
                    "Стоимость логистики": "sum",
                    "Общая сумма штрафов": "sum",
                    "Стоимость хранения": "sum",
                    "Стоимость платной приемки": "sum",
                    "Прочие удержания": "sum",
                    "Итого к оплате": "sum",
                }
            )
            .astype(int)
            .reset_index()
        )

        # Возвращаем дату в строковый формат для вывода
        sums_per_date["Дата конца"] = sums_per_date["Дата конца"].dt.strftime(
            "%d-%m-%Y"
        )

        ##########------ ПОСТРОЕНИЕ ГРАФИКА ------------
        buf = BytesIO()
        plt.figure(figsize=(10, 5))

        columns_to_plot = [
            "Продажа",
            "К перечислению за товар",
            "Стоимость логистики",
            "Итого к оплате",
        ]

        for column in columns_to_plot:
            plt.plot(
                sums_per_date["Дата конца"],
                sums_per_date[column],
                label=column,
                marker="o",
            )

        plt.title(f'Финансовые показатели (с {start_date.strftime("%d-%m-%Y")})')
        plt.xlabel("Дата")
        plt.ylabel("Сумма")
        plt.xticks(rotation=90)
        plt.legend()
        plt.grid(True)
        plt.tight_layout()
        plt.savefig(buf, format="png")
        plt.close()

        ##########------ СОЗДАНИЕ EXCEL-ФАЙЛА ------------
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            workbook = writer.book
            worksheet = workbook.create_sheet(title="Report")

            # Записываем DataFrame в лист
            for row in dataframe_to_rows(sums_per_date, index=False, header=True):
                worksheet.append(row)

            # Вставляем график
            img = OpenpyxlImage(buf)
            worksheet.add_image(img, "K10")  # График начинается с K10

        output.seek(0)

        # Показываем результат
        st.success("Обработка завершена!")
        st.download_button(
            label="⬇️ Скачать отчёт",
            data=output,
            file_name=f"wildberries_report_{start_date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Произошла ошибка: {str(e)}")
        st.stop()
else:
    st.info("Ожидание загрузки файлов...")
