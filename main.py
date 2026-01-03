import pandas as pd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from seleniumwire import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import requests

user_input = input("Названия компании: ")

options = webdriver.ChromeOptions()
options.add_argument("--headless")            # включаем headless режим
options.add_argument("--window-size=1920,1080")  # задаём размер окна
driver = webdriver.Chrome(options=options)
driver.get("https://new.openinfo.uz/ru?tab=facts&page=1")
wait = WebDriverWait(driver,10)
input_element = wait.until(
    EC.presence_of_element_located((By.CSS_SELECTOR,'input[placeholder="Поиск"]'))
)
input_element.send_keys(user_input)
div_element = wait.until(
    EC.presence_of_element_located((By.CSS_SELECTOR,'.absolute.z-10.w-full.mt-1.bg-white.border.border-default.rounded-xl.shadow-lg.max-h-60.overflow-y-auto.transition-transform.transition-opacity.duration-200.ease-in-out.opacity-100.translate-y-0'))
)
div_element.click()
balanse_otchet = wait.until(
    EC.presence_of_element_located((By.XPATH,'//button[text()="Балансовый отчет"]'))
)
balanse_otchet.click()
time.sleep(0.5)


balance_url = []
effect_url = []

for request in driver.requests:
    if request.response:
        content_type = request.response.headers.get("Content-Type", "")
        if "application/json" in content_type or "api" in request.url.lower():
            balance_url.append(request.url)

balance_report = balance_url
effect_otchet = wait.until(
    EC.presence_of_element_located((By.XPATH,'//button[text()="Показатели эффективности"]'))
)
effect_otchet.click()
time.sleep(0.5)

for request in driver.requests:
    if request.response:
        content_type = request.response.headers.get("Content-Type", "")
        if "application/json" in content_type or "api" in request.url.lower():
            effect_url.append(request.url)

driver.quit()
balance_api = balance_url[-1]
efficient_api = effect_url[-1]

url = balance_api
params = {
    "accounting_type": "form1",
    "report_type": "annual"
}
headers = {"User-Agent": "Mozilla/5.0"}

response = requests.get(url, params=params, headers=headers)
data = response.json()

periods = []
flat_data = []
for item in data:
    for report in item['accounting_report']:
        flat_data.append({
            "title": report.get("title"),
            "value1": report.get("value1"),
            "value2": report.get("value2")
        })

df = pd.DataFrame(flat_data)


period_frame = pd.DataFrame(data)
years = period_frame['period']
all_data = {
    'year':years,
    "Нераспределенная прибыль (непокрытый убыток) (8700)":[],
    "Долгосрочные обязательства, всего (стр.500+520+530+540+550+560+570+580+590)":[],
    "ВСЕГО по активу баланса 130+390":[],
}

for i in range(0,10):
    nerasp_income = df.loc[df["title"] == "Нераспределенная прибыль (непокрытый убыток) (8700)", "value2"].values[i]
    dolg = df.loc[df["title"] == "Долгосрочные обязательства, всего (стр.500+520+530+540+550+560+570+580+590)", "value2"].values[i]
    vsego_activ = df.loc[df["title"] == "ВСЕГО по активу баланса 130+390", "value2"].values[i]
    all_data["Нераспределенная прибыль (непокрытый убыток) (8700)"].append(nerasp_income)
    all_data["Долгосрочные обязательства, всего (стр.500+520+530+540+550+560+570+580+590)"].append(dolg)
    all_data["ВСЕГО по активу баланса 130+390"].append(vsego_activ)



efficient_url = efficient_api
headers = {"User-Agent": "Mozilla/5.0"}

efficient_response = requests.get(efficient_url,headers=headers)
efficient_data = efficient_response.json()


efficient_periods = []
efficient_flat_data = {
    "Чистая Прибыль":[],
    "Отношение долга к собственному капиталу":[],
    "Маржа EBIT":[],
    "Оборачиваемость активов":[]
}
for efficient_item in efficient_data['results']:  
    efficient_flat_data["Чистая Прибыль"].append(efficient_item.get("net_profit_margin"))
    efficient_flat_data["Отношение долга к собственному капиталу"].append(efficient_item.get("debt_to_equity_ratio"))
    efficient_flat_data["Маржа EBIT"].append(efficient_item.get("ebit_margin"))
    efficient_flat_data["Оборачиваемость активов"].append(efficient_item.get("total_asset_turnover"))

all_data.update(efficient_flat_data)
all_frames = pd.DataFrame(all_data)
all_frames.to_excel("data.xlsx")