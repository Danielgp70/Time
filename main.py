import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import openpyxl

driver = webdriver.Edge()

driver.get('https://myte.accenture.com/')
time.sleep(50)
libro = openpyxl.Workbook()
hoja = libro.active
hoja['A1'] = 'Horas total'
hoja['B1'] = 'Horas externas'
hoja['C1'] = 'Horas cliente'
#Datos de una persona

elemento_div = driver.find_element(By.XPATH, "//div[@id='total-1']")
elemento_div2 = driver.find_element(By.XPATH, "//div[@id='total-2']")
elemento_div3 = driver.find_element(By.XPATH, "//div[@id='total-3']")
elemento_div4 = driver.find_element(By.XPATH, "//div[@id='total-4']")
elemento_div_total = driver.find_element(By.XPATH, "//div[@id='footer-total-hours']")

ptotal1 = elemento_div.text
ptotal2 = elemento_div2.text
ptotal3 = elemento_div3.text
ptotal4 = elemento_div4.text
ptotal5 = elemento_div_total.text

print('Texto extraído:', ptotal1, ptotal2, ptotal3, ptotal4, ptotal5)

boton = driver.find_element(By.XPATH, "//button[@aria-label='Previous Period']")
boton.click()
time.sleep(3)
elemento_div = driver.find_element(By.XPATH, "//div[@id='total-1']")
elemento_div2 = driver.find_element(By.XPATH, "//div[@id='total-2']")
elemento_div3 = driver.find_element(By.XPATH, "//div[@id='total-3']")
elemento_div4 = driver.find_element(By.XPATH, "//div[@id='total-4']")
elemento_div_total = driver.find_element(By.XPATH, "//div[@id='footer-total-hours']")
# Obtener el texto del elemento
p2total1 = elemento_div.text
p2total2 = elemento_div2.text
p2total3 = elemento_div3.text
p2total4 = elemento_div4.text
p2total5 = elemento_div_total.text

if p2total2 == "":
    p2total2 = "0.0"
p2total2 = float(p2total2)
if ptotal2 == "":
    ptotal2 = "0.0"
ptotal2 = float(ptotal2)
if p2total3 == "":
    p2total3 = "0.0"
p2total3 = float(p2total3)
if ptotal3 == "":
    ptotal3 = "0.0"
ptotal3 = float(ptotal3)
if p2total4 == "":
    p2total4 = "0.0"
p2total4 = float(p2total4)
if ptotal4 == "":
    ptotal4 = "0.0"
ptotal4 = float(ptotal4)

mes_total1 = float(ptotal1) + float(p2total1)
mes_total2 = ptotal2 + p2total2
mes_total3 = ptotal3 + p2total3
mes_total4 = ptotal4 + p2total4
mes_total5 = float(ptotal5) + float(p2total5)

print('Texto extraído:', mes_total1, mes_total2, mes_total3, mes_total4, mes_total5)

reduccion_horas = mes_total2 + mes_total3 + mes_total4

if reduccion_horas == 0.0:
    print('Texto extraído:', mes_total1)
    total2 = mes_total1
else:
    reduccion_horas = float(reduccion_horas)
    print('Texto extraído:', reduccion_horas, mes_total1)
    print(mes_total5 - reduccion_horas)
i = 2
for i in range(i, 11):
    f = str(i)
    hoja['A' + f] = mes_total5
    hoja['B' + f] = reduccion_horas
    hoja['C' + f] = mes_total5 - reduccion_horas

libro.save('datos.xlsx')
# Cerrar el navegador web
driver.quit()
libro.close()
