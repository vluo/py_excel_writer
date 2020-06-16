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
    'align':'center',
    'font_name':'楷体'
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

introduce = '不管是dump还是load，带s的都是和字符串相关的，不带s的都是和文件相关的。\
不管是dump还是load，带s的都是和字符串相关的，不带s的都是和文件相关的。\
不管是dump还是load，带s的都是和字符串相关的，不带s的都是和文件相关的。\
不管是dump还是load，带s的都是和字符串相关的，不带s的都是和文件相关的。\
不管是dump还是load，带s的都是和字符串相关的，不带s的都是和文件相关的。\
不管是dump还是load，带s的都是和字符串相关的，不带s的都是和文件相关的。'

header = {'cols':[{'value':'标题', 'merge_col':8, 'format':title_format}], 'height':30}
merge_row_num = 3
data = [
    {'height':18, 'cols':[{'value':'分类', 'width':12,'format':header_format},{'width':12,'value':'产品名称', 'merge_col':2, 'format':header_format},{}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注 2', 'format':header_format}, {'width':10,'value':'备注 3', 'format':header_format}, {'width':10,'value':'备注 4', 'format':header_format}, {'width':10,'value':'备注 5', 'format':header_format}, {'width':10,'value':'备注 6', 'format':header_format}]},
    [{'value':'分类1', 'format':cat_format,'merge_row':3},{'value':'产品1', 'merge_col':2, 'data_type':'url', 'data_type_param':'http://ctoy.com'},{},{'value':'备注 1'}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}],
    [{},{'value':'产品 -2',  'data_type':'url', 'data_type_param':'http://ctoy.com'},{'value':'产品 -2', 'data_type':'url', 'data_type_param':'http://ctoy.com'},{'value':'note 2'}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}],
    #[{},{'value':'p2 -1',  'data_type':'url', 'data_type_param':'http://ctoy.com'},{'value':'p2 -2', 'data_type':'url', 'data_type_param':'http://ctoy.com'},{'value':'note 2'}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}, {'width':10,'value':'备注', 'format':header_format}]
]

save_path = '/home/python/httpserver/excels'
if not os.path.exists(save_path):
    os.makedirs(save_path);
#end def

start_time = time.time()
writer = myExcelWriter.myExcelWriter(save_path+'/test.xlsx');
json_data = []
for sheet_num in range(3):
    new_sheet = {}
    #writer.new_sheet('sheet_'+str(sheet_num+1))
    #writer.append_row(header)
    #writer.append_row(data[0])
    merge_row_num = 3
    new_sheet['name'] = 'sheet_'+str(sheet_num+1)
    new_sheet['rows'] = [header, data[0]]
    row_data = data[1:3]
    for row_num in range(10):
        row_data[0][0]['merge_row'] = merge_row_num
        new_sheet['rows'] = new_sheet['rows'] + row_data
        merge_row_num += 2
    #end if
    json_data.append(new_sheet)
#end for
#writer.save()
print(json_data[0]['rows'][2][0]['merge_row'])
json_str = json.dumps(json_data)
with open(save_path+'/data.json', 'w') as fh:
    fh.write(json_str)
#end

print('cost :'+str(time.time()-start_time));

print('done')


