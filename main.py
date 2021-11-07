# import module
import pandas as pd

######### [ НАСТРОЙКИ ] #########
FILE_NAME = "test.xlsx"        # ИМЯ ФАЙЛА, КОТОРЫЙ НУЖНО ОТКРЫТЬ (ЗАКИНУТЬ В ПАПКУ С ПРОГРАММОЙ)
FILE_OUTPUT = "output.xlsx"     # ИМЯ ФАЙЛА ДЛЯ СОХРАНЕНИЯ (СОХРАНИТСЯ В ПАПКЕ С ПРОГРАММОЙ)
#################################

# Excel Helper Engine

try:
    df = pd.read_excel(str(FILE_NAME))          # FILE NAME
    res = df.sort_values('Время договора').groupby('Наименование инструмента').tail(1)  # ФИЛЬТРАЦИЯ PANDAS
    res.to_excel(FILE_OUTPUT, index=False)    # FILE OUTPUT                                       
    print('Complete!')
except NameError:
    print("Error")



