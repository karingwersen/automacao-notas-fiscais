import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from time import sleep

if __name__ == '__main__':

    #path = "C:\\Users\\karin\\Downloads\\PLANILHA NF'S 2018.xlsx"
    #planilha = openpyxl.load_workbook(path)
    #sheet = planilha.active

    #max_row = sheet.max_row
    #for i in range(1, max_row + 1):
        #cell = sheet.cell(row = i, column = 1)
        #print(cell.value)

    chrome_options = Options()
    chrome_options.add_extension("fsist.crx")


    driver = webdriver.Chrome(options = chrome_options)
    driver.get("https://www.fsist.com.br/")

    caixa_texto = driver.find_element(By.ID, "chave")
    caixa_texto.send_keys("typing")
    consulta_nota = driver.find_element(By.ID, "butconsulta")
    consulta_nota.click()


    sleep(15)
    driver.quit()