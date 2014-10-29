# -*- coding: cp1251 -*-
import xlrd
import xlwt

class excelSort():
    def __init__(self):
        self.rd_data = \
            {
               u"������������" : [-1, []],
               u"�������" : [-1, []],
               u"�����" : [-1, []],
               u"������" : [-1, []],
               u"���-��" : [-1, []],
               u"���� ��� ���%" : [-1, []],
               u"���� � ���%" : [-1, []],
               u"����������" : [-1, []] 
            }
        self.wb = ''
        
    def fileOpen(self, name):
        ''' (sortExcel, str) -> bool
            ���������� �����.
        '''
        self.wb = xlrd.open_workbook(unicode(name))
        sheet = self.wb.sheet_by_index(0) 

        # ���������� �������� �����, ������������ �� ������ ������
        # ���� ���� ���� �� ����� ���������� ����� ������
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
            ����������
        '''
        # ��������� ������
        # �������� key_value - ������ ����� � �������
        row_data = \
        [
            u"�������",
            u"���-��",
            u"�����",
            u"������",
            u"���� ��� ���%"
        ]
        key_column = u"�������"
        count_column = u"���-��"
        self.wr_data = [[], []]

        # ���������� ����� ������ wr_data
        # wr_data[0] - �������� ����� key_column
        # wr_data[1] - ������ �� ���������� ��������� �� row_data
        for i in range(0, len(self.rd_data[key_column] [1])):
            # ������� ����� �� ������� � ��������� �������
            key = self.rd_data[key_column] [1] [i]
            # � �������� �� ������ ������ �����
            # ������� � ����� �������� ������!!!
            # key = key.rsplit() [0]
            # ������� ������� � ��������� ����������
            if key in self.wr_data[0]:
                # �������� ����� ��� ���� � ������

                # ��������� ����� ��� ������
                index = self.wr_data[0].index(key)
                # �������� �������� � ��������� ������
                val = self.rd_data[count_column] [1] [i]
                self.wr_data[1] [index] [row_data.index(count_column)] += val
            else:
                # �������� ������ �� ��������� ������� row_data
                row = []
                for col in row_data:
                   if col == key_column:
                       val = key
                   else:
                       val = self.rd_data[col] [1] [i]
                   row.append(val)
                # � �������� ������ ��� �������
                self.wr_data[0].append(key)
                self.wr_data[1].append(row)
            
    def fileSave(self, name):
        ''' (sortExcel, str) -> bool
            ���������� �����.
        '''
        # �������� ����� ����
        font0 = xlwt.Font()
        font0.name = 'Times New Roman'
        font0.colour_index = 2
        font0.bold = True

        style0 = xlwt.XFStyle()
        style0.font = font0

        wr_wb = xlwt.Workbook()
        ws = wr_wb.add_sheet(u'�������')
        for i in range(len(self.wr_data[1])):
            for j in range(len(self.wr_data[1] [i])):
                ws.write(i, j, unicode(self.wr_data[1] [i] [j]))
        wr_wb.save(unicode(name))

