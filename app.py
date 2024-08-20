# importing libraries
import openpyxl
import pyautogui

# calling the excel file and the file page
workbook = openpyxl.load_workbook('produtos.xlsx')
vendas_sheet = workbook['vendas']

#From where I want you to read the spreadsheet 
for linha in vendas_sheet.iter_rows(min_row=2):
    
    #name
    pyautogui.click(2972,236,duration=2)
    pyautogui.write(linha[0].value)
    #product
    pyautogui.click(2975,285,duration=2)
    pyautogui.write(linha[1].value)
    #amount
    pyautogui.click(3046,312,duration=2)
    pyautogui.write(str(linha[2].value))
    #category
    pyautogui.click(3046,312,duration=2)
    pyautogui.write(linha[3].value)
    #save
    pyautogui.click(2913,339,duration=2)
    #ok
    pyautogui.click(827,571,duration=2)
    print(linha[0].value)
    print(linha[1].value)
    print(linha[2].value)
    print(linha[3].value)

