from typing import Dict, List
import argparse
import json

import pandas as pd
import xlwings as xw
from xlwings import Sheet

from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.remote.webelement import WebElement


def set_config() -> Dict:
    arguments_parser = argparse.ArgumentParser(description="Crypto P&L")
    arguments_parser.add_argument("-r", "--regenerate", help="Regenerate the Excel file", action=argparse.BooleanOptionalAction)
    arguments = arguments_parser.parse_args()

    regenerate = arguments.regenerate

    return {
        "regenerate": regenerate
    }


def create_driver() -> WebDriver:
    options = webdriver.ChromeOptions()
    options.add_argument("headless")
    options.add_argument('log-level=2')
    return webdriver.Chrome(options=options)


def load_data() -> List[Dict]:
    with open("data.json", 'r') as file:
        data = json.load(file)
    return data


def create_excel_file(crypto_data: List[Dict]):
    dataframe = pd.DataFrame({
        "Nom de la crypto": [one_crypto["name"] for one_crypto in crypto_data],
        "Transaction": [list(zip(one_crypto["buy_price"], one_crypto["quantity"])) for one_crypto in crypto_data],
        "Montant investi (en $)": 0,
        "Prix actuel (en $)": 0,
        "P&L (en $)": 0,
        "P&L (en %)": 0,
        "P&L Total (en $)": 0,
        "P&L Total (en %)": 0
    })

    dataframe = dataframe.explode('Transaction', ignore_index=True)
    dataframe[["Prix d'achat (en $)", 'Quantité']] = pd.DataFrame(dataframe['Transaction'].tolist(), index=dataframe.index)
    dataframe.drop('Transaction', axis=1, inplace=True)

    dataframe["Montant investi (en $)"] = dataframe["Prix d'achat (en $)"] * dataframe["Quantité"]
    
    dataframe = dataframe[["Nom de la crypto", "Prix d'achat (en $)", "Quantité", "Montant investi (en $)", "Prix actuel (en $)", "P&L (en $)", "P&L (en %)", "P&L Total (en $)", "P&L Total (en %)"]]
    dataframe.to_excel("./Final.xlsx", sheet_name="Crypto", index=False)


def load_workbook() -> Sheet:
    # Load Excel File
    wb = xw.Book('Final.xlsx')
    sheet: Sheet  = wb.sheets["Crypto"]

    # Adjust the size of the column
    sheet.autofit()
    # Maximize the size of the window
    xw.apps.active.api.WindowState = -4137

    return sheet


def get_nb_rows_one_crypto(data: List[Dict], one_crypto_hash) -> int:
    return len([one_crypto_in_data["buy_price"] for one_crypto_in_data in data if one_crypto_in_data["hash"] == one_crypto_hash][0])


def select(driver: WebDriver, type_locator: str, name: str, delay: int = 3) -> WebElement :
    try:
        return WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((type_locator, name))
        )
    except TimeoutException:
        pass


def get_data_from_coinmarketcap(driver: WebDriver, crypto_hash: str) -> str:
    # Here we pass the hash of the crypto seen in the url
    # Ex : paal-ai
    url = f"https://coinmarketcap.com/currencies/{crypto_hash}"
    driver.get(url)

    coin_detail = select(driver, By.CLASS_NAME, "cmc-body-wrapper")
    coin_price = coin_detail.find_element(By.XPATH, "//span[contains(., '$')]").text

    return coin_price


def get_data_from_coinbrain(driver: WebDriver, crypto_hash: str) -> str:
    # Here we pass the hash of the crypto seen in the url
    # Ex : eth-0xa849cd6239906f23b63ba34441b73a5c6eba8a00 => For Etherium / HASHMIND
    url = f"https://coinbrain.com/coins/{crypto_hash}"
    driver.get(url)

    coin_price = driver.find_element(By.XPATH, "//span[contains(., '$')]").text
    return coin_price


async def get_coin_price(driver: WebDriver, crypto_hash: str) -> str:
    # Ex : 'Paal AI' => Name and 'paal-ai' => Hash
    try:
        return get_data_from_coinmarketcap(driver, crypto_hash)
    except AttributeError:
        # That mean that the coin is not in CoinMarketCap, so we gonna make a request to CoinBrain
        return get_data_from_coinbrain(driver, crypto_hash)


def calculate_profit_and_loss(buy_price: float, quantity: float, actual_price: float, method: str) -> float:
    # P&L $
    total_profit = (actual_price - buy_price) * quantity
    if method == "$":
        return round(total_profit, 2)
    # P&L %
    elif method == "%":
        initiale_invest = buy_price * quantity
        return round((total_profit / initiale_invest) * 100, 2)
    

def check_values_profit(sheet: Sheet, column: str, index: int):
    if sheet.range(f'{column}{index}').value < 0:
        sheet.range(f'{column}{index}').font.color = "#FF0000"
    elif sheet.range(f'{column}{index}').value > 0:
        sheet.range(f'{column}{index}').font.color = "#35ED00"
    else:
        sheet.range(f'{column}{index}').font.color = "#000000"


def write_in_excel(sheet: Sheet, crypto_name: str, result: str, index: int):
    end_of_index = index + get_nb_rows_one_crypto(load_data(), crypto_name) - 1
    sheet.range(f'E{index}:E{end_of_index}').value = result

    total_profit = 0
    total_profit_pourcentage = 0

    for current_index in range(index, end_of_index + 1):
        sheet.range(f'F{current_index}').value = calculate_profit_and_loss(
            float(sheet.range(f'B{current_index}').value),
            float(sheet.range(f'C{current_index}').value),
            float(sheet.range(f'E{current_index}').value),
            "$"
        )

        sheet.range(f'G{current_index}').value = calculate_profit_and_loss(
            float(sheet.range(f'B{current_index}').value),
            float(sheet.range(f'C{current_index}').value),
            float(sheet.range(f'E{current_index}').value),
            "%"
        )

        check_values_profit(sheet, "F", current_index)
        check_values_profit(sheet, "G", current_index)

        total_profit = total_profit + sheet.range(f'F{current_index}').value
        total_profit_pourcentage = total_profit_pourcentage + sheet.range(f'G{current_index}').value

    sheet.range(f'H{index}:H{end_of_index}').value = total_profit
    sheet.range(f'I{index}:I{end_of_index}').value = total_profit_pourcentage

    for current_index in range(index, end_of_index + 1):
        check_values_profit(sheet, "H", current_index)
        check_values_profit(sheet, "I", current_index)
