import openpyxl as xls
import datetime as time

"""OBSERVAÇÕES DE USO
Precisa atualizar o ano para conferir as segundas-feiras
"""
year = 2024

# givens:
# planned_col, actual_col, stages_col, top_line (contains the dates), file_path (from C:/), first_colored_col, sheet_name

file_path = 'C:/Users/USUARIO/Downloads/Cronograma.xlsx'
sheet_name = 'Planilha1'


def get_stage_info(file_path, sheet_name, planned_col, top_line):
    wb = xls.load_workbook(filename=file_path)
    ws = wb[sheet_name]

    line = top_line+1
    empty_lines_counter = 0
    stages_info = []
    while True:
        cell1 = ws.cell(row=line, column=planned_col)
        text1 = cell1.value
        if text1 == None:
            empty_lines_counter += 1
            line += 1
        else:
            substrings = text1.split(' ')
            begin_list = substrings[0].split('/')
            finish_list = substrings[2].split('/')
            dates = [time.date(year, int(begin_list[1]), int(begin_list[0])), 
                     time.date(year, int(finish_list[1]), int(finish_list[0]))]
            info = [line]+dates
            stages_info.append(info)

            empty_lines_counter = 0
            line +=1

        if empty_lines_counter >= 5: 
            break

    return stages_info


def get_weeks(begin, finish):
    begin_date = begin
    finish_date = finish
    day = time.timedelta(days=1)
    week = time.timedelta(days=7)

    for i in range(0, 7):
        if begin_date.isoweekday() == 1: break          # gets last monday before everything begins
        begin_date -= day

    for i in range(0,7):
        if finish_date.isoweekday() == 5: break         #gets first friday after everything ends
        finish_date += day


    mondays = []
    monday = begin_date
    while finish_date - monday > time.timedelta():
        mondays.append(monday)
        monday += week
    
    fridays = []
    friday = finish_date
    while friday - begin_date > time.timedelta():
        fridays.append(friday)
        friday -= week
    fridays = fridays[::-1]
    

    weeks = [[mondays[i], fridays[i]] for i in range(0, len(mondays))]
    print(len(mondays), len(fridays))
    return weeks


print(get_weeks(time.date(2024, 7, 1), time.date(2024, 12, 31)))