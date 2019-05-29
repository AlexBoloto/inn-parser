import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time


geckodriver_path = 'C:\\Users\\MikhaylovAV1\\PycharmProjects\\GKS_forms\\geckodriver.exe'
firefox_binary_path = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"


def init_driver():
    options = Options()
    options.headless = True
    driver = webdriver.Firefox(executable_path=geckodriver_path,
                               firefox_options=options,
                               firefox_binary=firefox_binary_path)
    driver.wait = WebDriverWait(driver, 60)
    driver.maximize_window()
    return driver


def find_column_with_organizations(dataframe):
    for i in range(len(dataframe.columns.values.tolist())):
        if 'Наименование' in dataframe.columns.values.tolist()[i]:
            return dataframe.columns.values.tolist()[i]


df = pd.read_excel('source.xlsx')
organiz_col = find_column_with_organizations(df)
for i in range(2):
    print(df[organiz_col][i])
print()
the_INN = []
driver = init_driver()

try:
    for i in range(len(df[organiz_col])):
        text = []
        title = []
        link = []
        print(df[organiz_col][i])
        try:
            driver.get('https://duckduckgo.com/?q=' + 'ИНН ' + df[organiz_col][i])
            texts = driver.wait.until(EC.presence_of_all_elements_located((By.XPATH,
                                                                           "//div[@class='result__snippet js-result-snippet']")))
            titles = driver.wait.until(EC.presence_of_all_elements_located((By.XPATH,
                                                                            "//a[@class='result__a']")))
            links = driver.wait.until(EC.presence_of_all_elements_located((By.XPATH,
                                                                           "//div[@class='result__extras__url']")))
            for element in titles:
                title.append(element.text)
            for element in links:
                link.append(element.text)
            for element in texts:
                text.append(element.text)
            res_str = "".join(map(str, text)) + "".join(map(str, title)) + "".join(map(str, link))
            try:
                inn = re.search(r'\b[597]\d{9}\b', res_str).group(0)
                print(inn)
            except AttributeError:
                inn = 0
            the_INN.append((inn, df[organiz_col][i]))
            time.sleep(1)
        except TypeError:
            pass
finally:
    driver.quit()
pd.DataFrame(the_INN).to_excel('output.xlsx', header=['Организация.ИНН', 'Организация.Наименование полное (ОИВ)'],
                               index=False)
