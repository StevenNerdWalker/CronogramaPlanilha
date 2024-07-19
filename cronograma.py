import openpyxl as xls
import datetime as time

"""OBSERVAÇÕES DE USO
Precisa atualizar o ano para conferir as segundas-feiras
"""
year = time.datetime.now().year

# givens:
# planned_col, actual_col, top_line (contains the dates), file_path (from C:/), first_colored_col, sheet_name

file_path = 'C:/Users/USUARIO/Downloads/teste.xlsx'
sheet_name = 'Planilha1'


def get_stage_dates(file_path, sheet_name, planned_col, actual_col, top_line):
    '''considera duas cells adjacentes com texto no formato "dd/mm até dd/mm" e assume que está no ano atual
    OBS: the actual begin and finish dates won't necessarily have a numerical value, they can be None. Just the planned ones are guaranteed to have a date.'''
    wb = xls.load_workbook(filename=file_path)
    ws = wb[sheet_name]

    line = top_line+1
    empty_lines_counter = 0
    stages_info = []

    while True:
        cell1 = ws.cell(row=line, column=planned_col)
        text1 = cell1.value

        cell2 = ws.cell(row=line, column=actual_col)
        text2 = cell2.value

        if text1 == None:
            empty_lines_counter += 1
            line += 1

        elif text1 != None:
            substrings_planned = text1.split(' ')       #['dd/mm', 'até', 'dd/mm']  planned dates
            begin_list_planned = substrings_planned[0].split('/')       #['dd', 'mm']
            finish_list_planned = substrings_planned[2].split('/')      #['dd', 'mm']
            dates_planned = [time.date(year, int(begin_list_planned[1]), int(begin_list_planned[0])), 
                        time.date(year, int(finish_list_planned[1]), int(finish_list_planned[0]))]

            if text2 != None:
                substrings_actual = text2.split(' ')        #same but for the executed dates
                begin_list_actual = substrings_actual[0].split('/')
                finish_list_actual = substrings_actual[2].split('/')
                dates_actual = [time.date(year, int(begin_list_actual[1]), int(begin_list_actual[0])), 
                                time.date(year, int(finish_list_actual[1]), int(finish_list_actual[0]))]
            elif text2 == None:
                dates_actual = [None, None]
            
            info = [line]+dates_planned+dates_actual        #[line, planned begin date, planned finish date, actual begin date, actual finish date]
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

    for i in range(0, 8):
        if begin_date.isoweekday() == 1: break          # gets last monday before everything begins
        begin_date -= day

    for i in range(0,8):
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
    fridays = fridays[::-1]     #places the fridays in cronological order
    

    weeks = [[mondays[i], fridays[i]] for i in range(0, len(mondays))]      #[[monday1, friday1], [mon2, fri2], ...]
    return weeks

def write_n_paint(file_path, sheet_name, top_col, first_colored_col, info):
    wb = xls.load_workbook(filename=file_path, read_only=False)
    ws = wb[sheet_name]
    cell = ws.cell(row=5, column=3)
    cell.value = 'cavalo'
    wb.save(filename=file_path)

write_n_paint(file_path=file_path, sheet_name=sheet_name, top_col=1, first_colored_col=1, info=[])