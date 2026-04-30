import os
import fitz # PyMuPDF
# Текущая директория
folder  =  r'\\MGTFS1\Work\07.АМ\ОВ\1420 Алабушево-Москва\ТПУ Крюково. Этап 16\!! Замечания ГГЭ\!! Третья загрузка\PDF'  # или полный путь, например r'C:\MyFolder'
files  =  [f for f in os.listdir(folder)]
#files  =  [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]



############################## Ведомость ГЧ ##############################
vedom  =  [file for file in files if ('Ведомость ГЧ') in file]

 



############################## Количество схем ##############################
# Количество схем по теплоснабжению
ts_sh = [file for file in files if ('ПС 'in file) and ('(ТС)' in file)]

# Количество схем по отоплению
otop_sh = [file for file in files if ('ПС 'in file) and ('(О)' in file)]

# Количество схем по отоплению и теплоснабжению
ts_otop_sh = [file for file in files if ('ПС' in file) and ('(О, ТС)' in file)]

# Количество схем по вентиляции и противодымной вентиляции
ventv_pdv_sh = [file for file in files if ('ПС' in file) and ('(В, ПДВ)' in file)]
# Количество схем по вентиляции 
vent_sh = [file for file in files if ('ПС' in file) and ('(В)' in file)]
# Количество схем попротиводымной вентиляции
ventpdv_sh = [file for file in files if ('ПС'in file) and ('(ПДВ)' in file)]

# Количество функциональных схем
func_sh = [file for file in files if ('ПС'in file) and ('(Ф)' in file)]


# Количество схем кондиционированию
kond_sh = [file for file in files if ('ПС' in file)  and ('(К)' in file)]

# Количество схем трансформаторных подстанций
transph_sh = [file for file in files if ('ПС' in file) and ('трансформат' in file)]



############################## Количество планов ##############################
# Количество планов по теплоснабжению
ts_pln = [file for file in files if ('Планы ' in file) and ('(ТС)' in file)]


# Планы по отоплению
otop_pln = [file for file in files if ('Планы ' in file) and ('(О)' in file)]

# Планы по отоплению и теплоснабжению
ts_otop_pln = [file for file in files if ('Планы ' in file) and ('(О, ТС)' in file)]


# Планы по вентиляции и противодымной вентиляции
vent_pdv_pln = [file for file in files if ('Планы ' in file) and ('(В, ПДВ)' in file)]
# Планы по вентиляции 
vent_pln = [file for file in files if ('Планы ' in file) and ('(В)' in file)]
# Планы попротиводымной вентиляции

pdv_pln= [file for file in files if ('Планы ' in file) and ('(ПДВ)' in file)]
# Планы кондиционированию
kond__pln = [file for file in files if ('Планы ' in file) and ('(К)' in file)]




var_val = [
    vedom,                #Ведомость ГЧ
    ts_sh,                # Количество схем по теплоснабжению
    otop_sh,              # Количество схем по отоплению
    ts_otop_sh,           # Количество схем по отоплению и теплоснабжению
    ventv_pdv_sh,         # Количество схем по вентиляции и противодымной вентиляции
    vent_sh,              # Количество схем по вентиляции
    ventpdv_sh,           # Количество схем по противодымной вентиляции
    func_sh,              #'Количество функциональных схем': 
    kond_sh,              # Количество схем по кондиционированию
    transph_sh,           # Количество схем трансформаторных подстанций
    ts_pln,               # Количество планов по теплоснабжению
    otop_pln,             # Количество планов по отоплению
    ts_otop_pln,          # Количество планов по отоплению и теплоснабжению
    vent_pdv_pln,         # Количество планов по вентиляции и противодымной вентиляции
    vent_pln,             # Количество планов по вентиляции
    pdv_pln,              # Количество планов по противодымной вентиляции
    kond__pln             # Количество планов по кондиционированию
]



dict_val = {
    'Ведомость ГЧ':vedom,
    'Количество схем по теплоснабжению': ts_sh,
    'Количество схем по отоплению': otop_sh,
    'Количество схем по отоплению и теплоснабжению': ts_otop_sh,
    'Количество схем по вентиляции и противодымной вентиляции': ventv_pdv_sh,
    'Количество схем по вентиляции': vent_sh,
    'Количество схем по противодымной вентиляции': ventpdv_sh,
    'Количество функциональных схем': func_sh,
    'Количество схем по кондиционированию': kond_sh,
    'Количество схем трансформаторных подстанций': transph_sh,
    'Количество планов по теплоснабжению': ts_pln,
    'Количество планов по отоплению': otop_pln,
    'Количество планов по отоплению и теплоснабжению': ts_otop_pln,
    'Количество планов по вентиляции и противодымной вентиляции': vent_pdv_pln,
    'Количество планов по вентиляции': vent_pln,
    'Количество планов по противодымной вентиляции': pdv_pln,
    'Количество планов по кондиционированию': kond__pln
}

summa = sum(list(map(len,var_val)))

#print(f' Количество листов графической часть - {summa} шт.')
list_for_print = [ f'{k} {len(v)} шт. ' for k,v in dict_val.items()]



list_num=['(01)','(02)','(03)','(04)','(05)','(06)','СО']


# функция которая определяет содержуться ли значения списка в ключевом слове
def is_pos(list_,source_text):
    lst = [i for i in list_ if i in source_text]
    return True if len(lst) != 0 else False
        
# Полные пути ко всем файлам  в папке
full_paths =[os.path.join(folder, f) for f in os.listdir(folder)]

    

# Полные пути к файлам  с графической частью

full_paths_graph =[os.path.join(folder, f) for f in os.listdir(folder)if is_pos(list_num,f)== False]



# Полные пути к файлам  с текстовой частью
full_paths_txt = [os.path.join(folder, f) for f in os.listdir(folder)
                  if is_pos(list_num,f)== True and 'СО' not in f]

# Полные пути к файлам  с текстовой частью
full_paths_spec = [os.path.join(folder, f) for f in os.listdir(folder)
                  if 'СО' in f]

# Файлы  с графической частью
files_graph =  [f for f in os.listdir(folder)
                if is_pos(list_num,f)== False]
# Файлы  с текстовой частью
files_txt  =  [f for f in os.listdir(folder)
               if is_pos(list_num,f)== True and 'СО' not in f]


# Файлы  с текстовой частью
files_spec  =  [f for f in os.listdir(folder)
               if 'СО' in f]

# Функция которая считает количество листов
def get_pdf_page_count(path):
    doc = fitz.open(path)
    page_count = doc.page_count  # или len(doc)
    doc.close()
    return page_count





# Количество листов с графической частью
count1 = 0
for fl,pth in zip(files_graph,full_paths_graph):
    pages = get_pdf_page_count(pth)
    count1 += pages
print(f' Количество листов с графической частью {count1} шт.')

# Количество листов с текстовой частью
count2 = 0
for fl,pth in zip(files_txt,full_paths_txt):
    pages = get_pdf_page_count(pth)
    count2 += pages
print(f' Количество листов с текстовой частью {count2} шт.')

# Количество листов спецификации
count3 = 0
for f,pth in zip(files_spec,full_paths_spec):
    pages = get_pdf_page_count(pth)
    count3 += pages
print(f' Количество листов спецификации {count3} шт.')   




count = 0
for f,pth in zip(files,full_paths):
    pages = get_pdf_page_count(pth)
    count += pages
print(f' Общее количество листов {count} шт.') 
