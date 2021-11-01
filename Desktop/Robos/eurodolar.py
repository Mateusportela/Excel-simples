# procurando o Dolar

from selenium import webdriver as opcoes_selenium_aula
from selenium.webdriver.common.keys import Keys
import pyautogui as tempoPausaPC
#biblioteca para abrir excel
import xlsxwriter
import os

meuNavegador = opcoes_selenium_aula.Chrome()
meuNavegador.get('https://www.google.com.br/')
#esperando 3 segundos
tempoPausaPC.sleep(3)

meuNavegador.find_element_by_name("q").send_keys('Dolar Hoje')
#esperando 3 segundos
tempoPausaPC.sleep(3)

meuNavegador.find_element_by_name("q").send_keys(Keys.ENTER)
#esperando 2 segundos
tempoPausaPC.sleep(2)

valorDolarPesq = meuNavegador.find_elements_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text


#-------------------------------------------------------------
#procurando o Euro

#esperando 2 segundos p/ procurar o Euro
tempoPausaPC.sleep(2)

meuNavegador.find_element_by_name("q").send_keys('')
tempoPausaPC.sleep(0)

#usando pyautogui p/ aperta tab
tempoPausaPC.press('tab')
tempoPausaPC.sleep(1)

#usando pyautogui p/ aperta enter e apagar as palavras
tempoPausaPC.press('enter')
tempoPausaPC.sleep(1)

meuNavegador.find_element_by_name("q").send_keys('Euro')
#esperando 3 segundos
tempoPausaPC.sleep(2)

meuNavegador.find_element_by_name("q").send_keys(Keys.ENTER)

valorEuroPesq = meuNavegador.find_elements_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text


#-----------------------------------------------------------------
#import xlsxwritr
#import os

nomeCaminhoArquivo = 'C:\\Users\\Windows10\\Desktop\\Salvado Excel\\Arquivo EuroDolar.xlsx'
planilhaCriada = xlsxwriter.Workbook(nomeCaminhoArquivo)
sheet1 = planilhaCriada.add_worksheet() #isso cria nova planilha em branco

sheet1.write("A1", "Dolar")
sheet1.write("B1", "Euro")
sheet1.write("A2", valorDolarPesq)
sheet1.write("B2", valorEuroPesq)


planilhaCriada.close()
os.startfile(nomeCaminhoArquivo)

print('Dolar e Euro extraido com sucesso')
