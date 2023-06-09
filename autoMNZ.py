# import python modules
import os

from docxtpl import DocxTemplate


# Документ -шаблон
doc = DocxTemplate("shablon_mnz.docx")

DATE = "25.01.2023"

###################################### ОБРАБОТКА каталога с файлами ############


# УКАзываем путь к директории с файлами
path = "c:/PYTHON/diplom_Viligdanov T/Files"

# Получаем список всех файлов в каталоге


def fun(x): return os.path.isfile(os.path.join(path, x))


files_list = filter(fun, os.listdir(path))

# Создайте список файлов в каталоге вместе с указанием размера
size_of_file = [
    (f, os.stat(os.path.join(path, f)).st_size)
    for f in files_list
]


# Выполнить итерацию по списку файлов с указанием размера
# и распечатайте их один за другим


############################ ---Заполнение МНЗ---##########################

filesRows = []
count = 0

for f, size in size_of_file:

    temp = str(f)
    start = temp.find('_')
    end = temp.find('.d')

    # print(end)
    titleDoc = temp[start + 2:end]
    # print(df)
    firstLetter = temp[start + 1].upper()
    # обозначение конструкторского документа состоит из первых 15 символов
    NameFile = f[0:15]

    # если нашли буквы вп, то к номеру файла прибавляем ВП
    lowerСase1 = temp.find('сб')
    capitalLetters1 = temp.find('СБ')

    # если нашли буквы вп, то к номеру файла прибавляем ВП
    lowerСase2 = temp.find('вп')
    capitalLetters2 = temp.find('ВП')

    # если нашли буквы вп, то к номеру файла прибавляем ВП
    lowerСase3 = temp.find('тэ4')
    capitalLetters3 = temp.find('ТЭ4')

    # если нашли буквы мэ, то к номеру файла прибавляем МЭ
    lowerСase4 = temp.find('мэ')
    capitalLetters4 = temp.find('МЭ')

    if lowerСase1 > 0 or capitalLetters1 > 0:
        inDx1 = 'Сборочный чертеж'
        inDx2 = 'СБ'
    elif lowerСase2 > 0 or capitalLetters2 > 0:
        inDx1 = 'Ведомость покупных изделий'
        inDx2 = 'ВП'
    elif lowerСase3 > 0 or capitalLetters3 > 0:
        inDx1 = 'Таблица соединений'
        inDx2 = 'ТЭ4'
    elif lowerСase4 > 0 or capitalLetters4 > 0:
        inDx1 = 'Электромонтажный чертеж'
        inDx2 = 'МЭ'

    else:
        inDx1 = ''
        inDx2 = ''

    # разделяем размер файла пробелами
    delimitedFileSize = int(size)
    delimitedFileSize = '{0:,}'.format(delimitedFileSize).replace(',', ' ')

    # проверка на тип файла документа (dwg или xls)
    data = str(size_of_file[count])
    flag = data.find("dwg") != -1

    if (flag == True):
        name_program = "AutoCad"
    else:
        name_program = "Exel"

    filesRows.append({"sNo": count + 1, "name_chertega": firstLetter + titleDoc + "\n" + inDx1, "name_file": f,
                      "program": name_program, "oboznach_file": NameFile + " " + inDx2, "size": delimitedFileSize})
    count += 1


context = {

    "filesRows": filesRows,
    "topItemsRows": " ",
    "date": DATE,
    "numberFiles": len(size_of_file)
}


doc.render(context)


reportWordPath = 'ITOG-MNZ-FILE.docx'
doc.save(reportWordPath)

print("Формирование документа завершено.")
