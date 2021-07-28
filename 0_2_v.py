from xlwt import Workbook
import xlwt

# пишешь сюда путь к файлу с общей сеткой пикет - профиль
GRID_file = open("Сетка.csv", 'r')

# пишешь сюда пути к файлам наряд заказов, с разными типами проб
MMI_file = open("MMI-ЗЕЛЕНЫЙ.csv", 'r')
LGH_file = open("ЛГХ-КРАСНЫЙ.csv", 'r')
NEOTBOR_file = open("неотбор-ЖЕЛТЫЙ.csv", 'r')
POR_file = open("ПОР-СИНИЙ.csv", 'r')

# создаю списки с номерами разных проб
MMI_list = []
LGH_list = []
NEOTBOR_list = []
POR_list = []

# общая сетка пикет - профиль
GRID_list = []

# число столбцов и строк в общей сетке
quantity_of_table_lines = 0
quantity_of_table_tops = 0

############################################################################################
# заполняю списки номерами соответствующих проб
############################################################################################
# заполняю список проб номерами из файла
for line in GRID_file:
    # разбиваю строку по разделителю на список подстрок
    mm = line.strip().split(';')
    GRID_list.append(mm)
    # счетчик строк в общей сетке
    quantity_of_table_lines += 1

# число столбцов в общей сетке
quantity_of_table_tops = len(mm)

for line in MMI_file:
    mm = line.strip().split(';')
    MMI_list.append(mm[0])

for line in LGH_file:
    mm = line.strip().split(';')
    LGH_list.append(mm[0])

for line in NEOTBOR_file:
    mm = line.strip().split(';')
    NEOTBOR_list.append(mm[0])

for line in POR_file:
    mm = line.strip().split(';')
    POR_list.append(mm[0])

############################################################################################
# создаю листы и стили
############################################################################################
book = Workbook()
sheet1 = book.add_sheet('Sheet 1', cell_overwrite_ok=True)
book.add_sheet('Sheet 2', cell_overwrite_ok=True)

# создаю стили, задаю цвет числом(см. скриншот с таблицей цветов в папке с программой)
white_style = xlwt.easyxf('pattern: pattern solid;')
green_style = xlwt.easyxf('pattern: pattern solid;')
red_style = xlwt.easyxf('pattern: pattern solid;')
yellow_style = xlwt.easyxf('pattern: pattern solid;')
blue_style = xlwt.easyxf('pattern: pattern solid;')

# неотмеченные пробы
white_style.pattern.pattern_fore_colour = 1
# зеленый ммi
green_style.pattern.pattern_fore_colour = 3
# лгх красный
red_style.pattern.pattern_fore_colour = 2
# неотбор желтый
yellow_style.pattern.pattern_fore_colour = 34
# пор синий
blue_style.pattern.pattern_fore_colour = 48

############################################################################################
# закрашиваю отобранные пробы на общей сетке соответствующим цветом
############################################################################################
# список всех отмеченных проб
noted_list = MMI_list + LGH_list + NEOTBOR_list + POR_list

# проход по столбцам сетки
for i in range(quantity_of_table_tops):
    # проход по строкам сетки
    for j in range(quantity_of_table_lines):

        # проход по списку проб mmi
        for k in range(len(MMI_list)):
            # сверяем наличие пробы в сетке
            if GRID_list[j][i] == noted_list[k]:
                # закрашиваем ячейку нужным цветом
                sheet1.write(j, i, GRID_list[j][i], white_style)

        for k in range(len(MMI_list)):
            if GRID_list[j][i] == MMI_list[k]:
                sheet1.write(j, i, GRID_list[j][i], green_style)

        for k in range(len(LGH_list)):
            if GRID_list[j][i] == LGH_list[k]:
                sheet1.write(j, i, GRID_list[j][i], red_style)

        for k in range(len(NEOTBOR_list)):
            if GRID_list[j][i] == NEOTBOR_list[k]:
                sheet1.write(j, i, GRID_list[j][i], yellow_style)

        for k in range(len(POR_list)):
            if GRID_list[j][i] == POR_list[k]:
                sheet1.write(j, i, GRID_list[j][i], blue_style)

# пишешь каждый раз новое название файла с расширением .xls - в него запишется результат
book.save('2simple.xls')
