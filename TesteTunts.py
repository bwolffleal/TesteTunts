import os
import sys
import platform
import pandas as pd
import openpyxl

    
def Main():
    try:
        os_name = platform.system()
        if os_name == "Linux":
            print("Using Linux")
            dirname = os.getcwd()
            filename = dirname + "/Engenharia de Software - Desafio Bernardo Wolff Leal.xlsx"
            sheetname = 'engenharia_de_software'
        if os_name == "Windows":
            print("Using Windows")
            dirname = os.getcwd()
            filename = dirname + "\Engenharia de Software - Desafio Bernardo Wolff Leal.xlsx"
            sheetname = 'engenharia_de_software'
            
        wb = openpyxl.load_workbook(filename)
        ws = wb[sheetname]
        iteration = 0
        for row in ws.iter_rows():
            if iteration < 3:
                iteration = iteration + 1
                continue
            studentname = row[1].value.encode('utf-8')
            print("Aluno: " + str(studentname))
            attendance = 100-(((float(row[2].value))/60)*100)
            format_float = "{:.2f}".format(attendance)
            print("Presença: " + str(format_float) + "%")
            if row[2].value>15:
                print("Situação: Reprovado por falta")
                row[6].value = str("Reprovado por falta")
                row[7].value = str("0")
            else:
                grades = (row[3].value+row[4].value+row[5].value)/3
                rounding = grades - int(grades)
                if rounding >= 0.5:
                    grades = round(grades)
                    print("Nota arredondada")
                print("Média Final: " + str(grades))
                if grades < 50:
                    row[6].value = str("Reprovado por nota")
                    row[7].value = str("0")
                    print("Situação: Reprovado por nota")
                elif grades >= 70:
                    row[6].value = str("Aprovado")
                    row[7].value = str("0")
                    print("Situação: Aprovado")
                else:
                    row[6].value = str("Exame final")
                    print("Situação: Exame final")
                    finalgradeneeded = 100 - grades
                    print("Nota para aprovação final: " + str(finalgradeneeded))
                    row[7].value = str(finalgradeneeded)
        wb.save(filename)    
        df = pd.read_excel(io=filename, sheet_name=sheetname, header=3)
        print(df)
        
    except IOError:
        print('File not found')
        
if __name__ == '__main__':
    Main()
