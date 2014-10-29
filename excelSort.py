# -*- coding: cp1251 -*-
import xlrd
import xlwt

class excelSort():
    def __init__(self):
        self.rd_data = \
            {
               u"ПОЛЬЗОВАТЕЛЬ" : [-1, []],
               u"АРТИКУЛ" : [-1, []],
               u"ТОВАР" : [-1, []],
               u"РАЗМЕР" : [-1, []],
               u"КОЛ-ВО" : [-1, []],
               u"ЦЕНА БЕЗ ОРГ%" : [-1, []],
               u"ЦЕНА С ОРГ%" : [-1, []],
               u"ПРИМЕЧАНИЕ" : [-1, []] 
            }
        self.wb = ''
        
    def fileOpen(self, name):
        ''' (sortExcel, str) -> bool
            Считывание файла.
        '''
        self.wb = xlrd.open_workbook(unicode(name))
        sheet = self.wb.sheet_by_index(0) 

        # заполнение индексов полей, ориентируясь на первую строку
        # если хоть одно из полей отсутсвует будет ошибка
        ##row = sheet.row_values(0)
        ##for el in rd_index:
        ##    rd_index[el] [0] = row.index(el) 

        for i in range(sheet.ncols):
            col = sheet.col_values(i)
            name_col = col[0]
            if self.rd_data.has_key(name_col):
                self.rd_data[name_col] [1] = col

    def sortOne(self):
        ''' (sortExcel) -> bool
            Сортировка
        '''
        # формируем список
        # значение key_value - первое слово в столбце
        row_data = \
        [
            u"АРТИКУЛ",
            u"КОЛ-ВО",
            u"ТОВАР",
            u"РАЗМЕР",
            u"ЦЕНА БЕЗ ОРГ%"
        ]
        key_column = u"АРТИКУЛ"
        count_column = u"КОЛ-ВО"
        self.wr_data = [[], []]

        # сформируем новый список wr_data
        # wr_data[0] - ключевое слово key_column
        # wr_data[1] - список со значениями элементов из row_data
        for i in range(0, len(self.rd_data[key_column] [1])):
            # считаем стору из колонки с ключевыми словами
            key = self.rd_data[key_column] [1] [i]
            # и извлечем из строки первое слово
            # которое и будет являться ключом!!!
            # key = key.rsplit() [0]
            # перебор столбца с ключевыми значениями
            if key in self.wr_data[0]:
                # ключевое слово уже есть в списке

                # определим номер его строки
                index = self.wr_data[0].index(key)
                # прибавим значение к имеющейся записи
                val = self.rd_data[count_column] [1] [i]
                self.wr_data[1] [index] [row_data.index(count_column)] += val
            else:
                # заполним строку из имеющихся колонок row_data
                row = []
                for col in row_data:
                   if col == key_column:
                       val = key
                   else:
                       val = self.rd_data[col] [1] [i]
                   row.append(val)
                # и заполним нужную нам таблицу
                self.wr_data[0].append(key)
                self.wr_data[1].append(row)
            
    def fileSave(self, name):
        ''' (sortExcel, str) -> bool
            Сохранение файла.
        '''
        # созданим новый файл
        font0 = xlwt.Font()
        font0.name = 'Times New Roman'
        font0.colour_index = 2
        font0.bold = True

        style0 = xlwt.XFStyle()
        style0.font = font0

        wr_wb = xlwt.Workbook()
        ws = wr_wb.add_sheet(u'Закупка')
        for i in range(len(self.wr_data[1])):
            for j in range(len(self.wr_data[1] [i])):
                ws.write(i, j, unicode(self.wr_data[1] [i] [j]))
        wr_wb.save(unicode(name))

