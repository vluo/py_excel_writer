#! /usr/bin/env python3
# -*- coding:UTF-8 -*-
from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import cgi
import os
import datetime
import time
from libs import myExcelWriter

host = ('', 88)

'''engine = create_engine(
    'mysql+pymysql://pointgrab_user:pointgrabAaaa1111@rm-bp1d2s03ka9b803602o.mysql.rds.aliyuncs.com:3306/pointgrab_info',
    echo=False)'''


class TodoHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        self.send_error(415, 'Only post is supported')

    def do_POST(self):
        ctype, pdict = cgi.parse_header(self.headers['content-type'])

        if ctype == 'application/json':
            path = ""+self.path.replace('/', '').lower()  # 获取请求的url
            print('path='+path)
            #self.write_excel()
            try:
                method = getattr(self, path)
                method()
            except Exception as ex:
                print(ex)
                self.send_error(500, str(ex))
            #end try
        else:
            self.send_error(415, "Only json data is supported.")


    def __cur_dir(self):
        file_path = os.path.join(os.getcwd(), 'excels')
        #print('excel patg='+file_path);
        if not os.path.exists(file_path):
            os.makedirs(file_path)
        #end dif
        return file_path
    #end def

    def write_excel(self):
        post = self.load_post()
        if post:
            taget_file = os.path.join(self.__cur_dir(), 'excel_'+str(time.time())+".xlsx")
            print('taget_file=' + taget_file);
            writer = myExcelWriter.myExcelWriter(taget_file)
            if writer.import_from_data(post):
                self.render_success(taget_file)
            else:
                self.render_error('error'+writer.error)
            #end if
        else:
            self.render_error('empty json param')
        #end if
    #end def

    def load_post(self):
        # print(path)
        length = int(self.headers['content-length'])  # 获取除头部后的请求参数的长度
        datas = self.rfile.read(length)  # 获取请求参数数据，请求数据为json字符串
        #print('data > ', datas)
        try :
            return json.loads(datas.decode())
        except:
            print('error json param')
            return None
        #end try
        # print(rjson,type(rjson))
    #end def

    def render_error(self, error='failed'):
        self.render({'data':None, 'message':error, 'status':0})
    #end def

    def render_success(self, data=None, msg="done"):
        print('>>>> r s')
        self.render({'data':data, 'message':msg, 'status':1})
    #end def

    def render(self, data):
        self.send_response(200)
        self.send_header('Content-type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())
    #end def


if __name__ == '__main__':
    server = HTTPServer(host, TodoHandler)
    print("Starting server, listen at: %s:%s" % host)
    server.serve_forever()