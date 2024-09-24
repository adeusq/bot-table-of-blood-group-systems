import pyautogui
import time
import openpyxl

def ler_dados_excel(arquivo_excel):
    
    wb = openpyxl.load_workbook(arquivo_excel)
    sheet = wb.active
    dados = []
    
    for row in sheet.iter_rows(min_row=1, values_only=True):
        dados.append(row)
    
    return dados

def preencher_sistema(dados, coordenadas, coordenada_inicial, coordenada_final):
    
    time.sleep(3)
    
    for linha in dados:
        
        pyautogui.click(coordenada_inicial[0], coordenada_inicial[1])

        time.sleep(0.5)
        
        for valor, coordenada in zip(linha, coordenadas):

            pyautogui.click(coordenada[0], coordenada[1])

            time.sleep(0.5)

            pyautogui.write(str(valor))
        
        pyautogui.click(coordenada_final[0], coordenada_final[1])
        
        time.sleep(0.5)

arquivo_excel = r'C:\project\bot-table-of-blood-group-systems\Tableofbloodgroupsystems.xlsx'

coordenadas = [
    (610,202),  
    (617,303), 
    (609,391),  
    (600,490),  
    (613,582), 
    (597,676), 
    (598,765),  
    (607,856)   
]

coordenada_inicial = (87, 221)

coordenada_final = (885, 922)

dados = ler_dados_excel(arquivo_excel)

preencher_sistema(dados, coordenadas, coordenada_inicial, coordenada_final)
