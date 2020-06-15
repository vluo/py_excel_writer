import json
import time
import os

from libs import myExcelWriter

# Create a workbook and add a worksheet.

bg_color = '#cccccc'
border_color = '#FF6600'
common_format = {
    'border':1,
    'border_color':border_color,
    'align':'center',
    'valign':'vcenter'
}
header_format = {
    'font_size':12,
    'bold':True,
    'bg_color':bg_color,
    'align':'center'
}
title_format = {
    'font_size':14,
    'bold':True,
    'font_color':'#FF0000'
}
cat_format = {
    'text_wrap':True
}

cat_format.update(common_format)
title_format.update(common_format)
header_format.update(common_format)

introduce = '中文内容'

header = {'cols':[{'value':'title', 'merge_col':3, 'format':title_format}], 'height':30}
merge_row_num = 3
data = [
    [{'value':'cat name', 'width':12,'format':header_format},{'width':12,'value':'prod name', 'merge_col':2, 'format':header_format},{}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}],
    [{'value':'分类1', 'format':cat_format,'merge_row':merge_row_num},{'format':common_format,'value':'产品1'+introduce, 'merge_col':2, 'data_type':'url0', 'data_type_param':'http://ctoy.com'},{'format':common_format},{'format':common_format,'value':'备注 1'}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}],
    [{'format':common_format},{'value':'产品 -2'+introduce, 'format':common_format, 'data_type':'url0', 'data_type_param':'http://ctoy.com'},{'format':common_format,'value':'产品 -2'+introduce, 'data_type':'url0', 'data_type_param':'http://ctoy.com'},{'format':common_format,'value':'note 2'}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}],
    #[{'format':common_format},{'value':'p2 -1', 'format':common_format, 'data_type':'url', 'data_type_param':'http://ctoy.com'},{'format':common_format,'value':'p2 -2', 'data_type':'url', 'data_type_param':'http://ctoy.com'},{'format':common_format,'value':'note 2'}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}, {'width':10,'value':'note'+introduce, 'format':header_format}]
]

if not os.path.exists('/home/python/excelwriter/excels'):
    os.makedirs('/home/python/excelwriter/excels');
#end def

start_time = time.time()
writer = myExcelWriter.myExcelWriter('/home/python/excelwriter/excels/test.xlsx');
for sheet_num in range(1):
    writer.new_sheet('sheet_'+str(sheet_num+1))
    writer.append_row(header)
    writer.append_row(data[0])
    merge_row_num = 3
    row_data = data[1:3]
    for row_num in range(20000):
        row_data[0][0]['merge_row'] = merge_row_num
        writer.append_rows(row_data)
        merge_row_num += 2
    #end if
#end for
writer.save()


print('cost :'+str(time.time()-start_time));

print('done')


