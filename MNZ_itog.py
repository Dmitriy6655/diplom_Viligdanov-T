# import python modules
import os

from docxtpl import DocxTemplate, InlineImage

# create a document object
# Документ -шаблон
doc = DocxTemplate("shablon_mnz.docx")

DATE = "25.01.2023"

###################################### ОБРАБОТКА катлога с файлами ############


##УКАзываем путь к директории с файлами
path = "c:/PYTHON/diplom_Viligdanov T/Files"

# Получаем список всех файлов в каталоге
fun = lambda x: os.path.isfile(os.path.join(path, x))

files_list = filter(fun, os.listdir(path))

# Создайте список файлов в каталоге вместе с указанием размера
size_of_file = [
    (f, os.stat(os.path.join(path, f)).st_size)
    for f in files_list
]

# Iterate over list of files along with size
# and print them one by one.
# Выполнить итерацию по списку файлов с указанием размера
# и распечатайте их один за другим

# count = int(0)
# Функция для обработки списка


# format(GH, df, inDx1, f[0:], NumFile, inDx2, s)


# print("df=", GH + df)
# print("inDx1=", inDx1)
# print("inDx2=", inDx2)
# print("f=", f)
# print(type(f))
# print("size_file=", size)
#
# print("NumFile=", NumFile)
# print(size_of_file)
#
# print(len(size_of_file))

############################---Заполнение МНЗ---##########################


# create data for reports
salesTblRows = []
# for k in range(len(size_of_file)):

for f, size in size_of_file:
    k = 0
    g = str(f)
    start = g.find('_')
    end = g.find('.d')

    asg = g.find('сб')  # если нашли буквы вп, то к номеру файла прибавляем ВП
    asg1 = g.find('СБ')

    asg2 = g.find('вп')  # если нашли буквы вп, то к номеру файла прибавляем ВП
    asg22 = g.find('ВП')

    asg3 = g.find('тэ4')  # если нашли буквы вп, то к номеру файла прибавляем ВП
    asg33 = g.find('ТЭ4')

    asg4 = g.find('мэ')  # если нашли буквы мэ, то к номеру файла прибавляем МЭ
    asg44 = g.find('МЭ')

    if asg > 0 or asg1 > 0:
        inDx1 = 'Сборочный чертеж'
        inDx2 = 'СБ'
    elif asg2 > 0 or asg22 > 0:
        inDx1 = 'Ведомость покупных изделий'
        inDx2 = 'ВП'
    elif asg3 > 0 or asg33 > 0:
        inDx1 = 'Таблица соединений'
        inDx2 = 'ТЭ4'
    elif asg4 > 0 or asg44 > 0:
        inDx1 = 'Электромонтажный чертеж'
        inDx2 = 'МЭ'

    else:
        inDx1 = ''
        inDx2 = ''

    # print(end)
    df = g[start + 2:end]
    # print(df)
    GH = g[start + 1].upper()
    NameFile = f[0:15]  # обозначение конструкторского документа состоит из первых 15 символов

    # разделяем размер файла пробелами
    num = int(size)
    num = '{0:,}'.format(num).replace(',', ' ')

    # проверка на тип файла документа (dwg или xls)
    data = str(size_of_file[k])
    flag = data.find("dwg") != -1

    if (flag == True):
        name_program = "AutoCad"
    else:
        name_program = "Exel"

    salesTblRows.append(
        {"sNo": k + 1, "name_chertega": GH + df + "\n" + inDx1, "name_file": f, "program": name_program,
         "oboznach_file": NameFile + " " + inDx2, "size": num})
    k += 1

# create context to pass data to template
context = {

    "salesTblRows": salesTblRows,
    "topItemsRows": " ",
    "date": DATE,
    "kolich_file": len(size_of_file)
}

# render context into the document object
doc.render(context)

# save the document object as a word file
reportWordPath = 'ITOG-MNZ-FILE.docx'
doc.save(reportWordPath)

print("Формирование документа завершено.")