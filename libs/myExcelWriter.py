import xlsxwriter
import json
import time
import os


class myExcelWriter():
    workbook = None
    sheet = None
    cur_row_num = 0

    def __init__(self, file_path):
        self.workbook = xlsxwriter.Workbook(file_path, {'constant_memory': False})
    #end def
    def new_sheet(self, title):
        self.sheet = self.workbook.add_worksheet(title)
        self.cur_row_num = 0
    #end def

    def append_row(self, row_data):

        if 'height' in row_data:
            self.sheet.set_row(self.cur_row_num, int(row_data['height']))
        #end if
        row_data = row_data['cols'] if 'cols' in row_data else row_data

        #print(row_data)
        for i in range(len(row_data)):
            col_data = row_data[i]

            format = None
            if 'format' in col_data:
                format = self.__parse_format(col_data['format'])
            # end def

            if not col_data:
                col_data = {'value':''}
            #end if
            if 'value' not in col_data:
                col_data['value'] = ''
            #end if
            if 'data_type' not in col_data:
                col_data['data_type'] = 'default'
            #end def
            if 'width' in col_data:
                self.sheet.set_column(i, i, int(col_data['width']))
            #end def

            #print(str(self.cur_row_num)+ ':'+str(i)+' val='+col_data['value'])
            # 动态调用类方法的关键
            methods = {
                'url': 'write_url',
                'default': 'write_cell'
            }
            method = methods.get(col_data['data_type'], 'write_cell')
            obj_method = getattr(self, method)

            #单元格合并
            if 'merge_row' in col_data:
                self.sheet.merge_range(self.cur_row_num, i, int(col_data['merge_row']), i, col_data['value'], format)
                if method != 'write_cell':
                    obj_method(self.cur_row_num, i, col_data, format)
                # end def
            elif 'merge_col' in col_data:
                self.sheet.merge_range(self.cur_row_num, i, self.cur_row_num, int(col_data['merge_col']), col_data['value'], format)
                if method != 'write_cell':
                    obj_method(self.cur_row_num, i, col_data, format)
                # end def
            else:
                obj_method(self.cur_row_num, i, col_data, format)
            #end if


        #end for
        self.cur_row_num += 1
    #end def

    def write_url(self, row, col, data, format):
        #{'constant_memory': True}
        param = data['data_type_param'] if 'data_type_param' in data else "";
        param = data['value'] if param=="" else param;
        self.sheet.write_url(row, col, param, format, data['value'])
    #end def

    def write_cell(self, row, col, data, format):
        #print('cel value '+ data['value'])
        self.sheet.write(row, col, data['value'], format)
    #end def

    def append_rows(self, data, start_row=-1):
        if start_row != -1:
            self.cur_row_num = start_row
        #end fi

        for row in data:
            self.append_row(row)
        #end for
        #self.workbook.close()
    #end def

    def save(self):
        self.workbook.close()
    # end def


    def set_range_format(self, row, col, format):
        format = self.__parse_format(format)

        if format is not None:
            self.worksheet.set_column('A:D', 20, cell_format)
        #END if
    #end def

    def __parse_format(self, format):
        if isinstance(format, str):
            try:
                format = json.dumps(format)
            except:
                format = None
            # end try
        elif isinstance(format, dict):
            format = self.workbook.add_format(format)
        else:
            format = None
        # end if

        if format is None:
            format = self.workbook.add_format(format)
        #END if

        return format
    #end def

#end class