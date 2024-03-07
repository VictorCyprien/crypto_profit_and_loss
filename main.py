import time
import os
from typing import List, Dict

import asyncio

from helpers import (
    set_config, 
    create_driver, 
    load_data, 
    create_excel_file, 
    load_workbook,
    get_nb_rows_one_crypto,
    get_coin_price,
    write_in_excel
)


async def main(crypto_data: List[Dict]):
    current_sheet = load_workbook()
    current_index = 0
    print("Maintenez CTRL + C pour sortir du script")
    time.sleep(3)
    crypto_hash_list = [one_crypto["hash"] for one_crypto in crypto_data]
    while True:
        for current_loop_index, one_crypto_hash in enumerate(crypto_hash_list, start=2):
            loop = asyncio.get_event_loop()
            result = loop.create_task(get_coin_price(driver, one_crypto_hash))
            write_in_excel(current_sheet, one_crypto_hash, await result, current_loop_index + current_index)
            nb_rows = get_nb_rows_one_crypto(crypto_data, one_crypto_hash) - 1
            current_index = current_index + nb_rows
        current_index = 0
        time.sleep(2)


if __name__ == "__main__":
    config = set_config()
    regenerate = config["regenerate"]
    crypto_data = load_data()

    if regenerate and os.path.isfile("./Final.xlsx"):
        print("Regenerating Excel file...")
        os.remove("./Final.xlsx")
        create_excel_file(crypto_data)
    elif not os.path.isfile("./Final.xlsx"):
        print("Generating Excel file...")
        create_excel_file(crypto_data)

    driver = create_driver()
    asyncio.run(main(crypto_data))
