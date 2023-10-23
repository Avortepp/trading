from docx import Document
import pandas as pd
import talib
import matplotlib.pyplot as plt
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def docx_to_csv(docx_file, csv_file):
    doc = Document(docx_file)
    pinescript_code = []

    for paragraph in doc.paragraphs:
        pinescript_code.append(paragraph.text)

    # Відокремлюємо числові та текстові дані
    numeric_data = []
    text_data = []
    for line in pinescript_code:
        try:
            numeric_data.append(float(line))
        except ValueError:
            text_data.append(line)

    # Створюємо DataFrame з числових даних
    numeric_df = pd.DataFrame(numeric_data, columns=['Numeric Data'])

    # Зберігаємо числові дані у CSV файлі
    numeric_df.to_csv(csv_file, index=False)

    # Обробляємо текстові дані (ви можете використовувати їх за потреби)
    # ...

    # Робимо графік для числових даних
    plt.figure(figsize=(10, 6))
    plt.plot(numeric_df['Numeric Data'], label='Дані для графіка', color='blue', linestyle='-', marker='o')
    plt.xlabel('Індекс')
    plt.ylabel('Значення')
    plt.title('Графік даних')
    plt.legend()

    # Зберігаємо графік в файл plot.png
    plt.savefig('plot.png')

    # Закриваємо графік
    plt.close()
    # Создаем веб-драйвер и открываем веб-сайт TradingView
    browser = webdriver.Chrome()
    browser.get('https://www.tradingview.com/chart/')
    # (остальной код остается без изменений)

    # Находим поле поиска на TradingView и вводим "SMA"
    search_box = WebDriverWait(browser, 200).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '.tv-screener-toolbar__search input')))
    search_box.send_keys('SMA')

    # Ожидаем загрузки результатов поиска (ожидание 30 секунд)
    WebDriverWait(browser, 30).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, '.tv-screener-toolbar-popup__popup-content')))

    # Выбираем SMA из результатов поиска
    sma_indicator = browser.find_element_by_css_selector('.tv-screener-toolbar__indicators .tv-screener-toolbar__indicator')
    sma_indicator.click()

    # Ожидаем загрузки данных с индикатором SMA (ожидание 30 секунд)
    WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.pane-legend')))
    # Код для дополнительных действий с индикатором SMA можно добавить здесь

    # Закрываем веб-браузер
    browser.quit()

# Пути к файлам
docx_file_path = r'C:\\TEST\\MSBOB.docx'
csv_file_path = 'output_data.csv'
try:
    df = docx_to_csv(docx_file_path, csv_file_path)
    print("Дані успішно оброблено та графік збережено у файлі plot.png.")
except Exception as e:
    print(f"Сталася помилка: {str(e)}")
# Вызываем функцию и обрабатываем данные
df = docx_to_csv(docx_file_path, csv_file_path)