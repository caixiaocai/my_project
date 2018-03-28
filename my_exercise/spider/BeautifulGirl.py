# coding:utf-8
from urllib import request,parse
import requests
import os
import json
import sys
import xlwt  # 安装库，操作Excel，你需要两个库：xlwt(写Excel) 和 xlrd(读Excel)
# reload(sys)
# sys.setdefaultencoding('utf-8')

'''
*****************************************************此爬虫的流程：*****************************************************
1.初始化数据（初始化字段，excel工作簿中的表格，并写好了第一栏数据，然后查询你想要的数据）
2.查询数据（选择你想要抓取的内容：年龄的范围，薪酬范围等，然后抓取数据）
3.抓取数据（抓取你想要的内容，headers，url，要查询的数据query_data，然后发起请求，解析返回的数据，当前页抓完之后就自动抓取下一页）
4.解析数据（解析返回的数据，循环我想要的数据，分别写入硬盘{对应的文件夹}和excel表格中，循环解析当前页的抓取数据）
5.保存数据（存储我想要的数据到指定地方，例如：mysql，excel）
************************************************************************************************************************

写入Excel：

xlwt.Workbook()：创建一个工作薄

工作薄对象.add_sheet(cell_overwrite_ok=True)：添加工作表，括号里是可选
参数，用于确认同一个cell单元是否可以重设值

工作表对象.write(行号，列号，插入数据，风格)，第四个参数可选

工作薄对象.save(Excel文件名)：保存到Excel文件中

读取Excel：

xlrd.open_workbook()：读取一个Excel文件获得一个工作薄对象

工作薄对象.sheets()[0]：根据索引获得工作薄里的一个工作表

工作表对象.nrows：获得行数

工作表对象.ncols：获得列数

工作表对象.row_values(pos)：读取某一行的数据，返回结果是列表类型的


'''

class wzly(object):

    def __init__(self):
        self.gender = 0 #性别
        self.start_age = 0
        self.end_age = 0
        self.start_height = 0.0
        self.end_height = 0.0
        self.marry = 1
        self.salary = 0
        self.eduction = 0
        self.count = 1  #表示数据条数,这里先去掉第一行的标题，所以是从1开始
        self.f = None
        self.sheetInfo = None
        self.create_excel()
        pass

    def create_excel(self):
        '''创建Excel'''
        # 创建工作薄
        self.f = xlwt.Workbook()
        # 工作薄对象.add_sheet(cell_overwrite_ok=True)：添加工作表，括号里是可选参数，用于确认同一个cell单元是否可以重设值
        self.sheetInfo = self.f.add_sheet('我主良缘', cell_overwrite_ok=True)
        rowTitle = ['编号', '昵称', '性别', '年龄', '身高', '籍贯', '学历', '内心独白', '照片']
        # 填充标题
        for i in range(0, len(rowTitle)):
            self.sheetInfo.write(0, i, rowTitle[i])  # 工作表对象.write(行号，列号，插入数据，风格)，第四个参数可选

    def query_data(self):
        '''
        筛选条件
        年龄，
        身高，
        教育，
        期望薪资,
        :return:
        '''
        input('请输入你的筛选条件，直接回车可以忽略本筛选条件：')
        self.query_age()
        self.query_sex()
        self.query_height()
        self.query_money()
        print('你的筛选条件是年龄:{}-{}岁\n性别是:{}\n对方身高是:{}-{}\n对方月薪是:{}'.format(self.start_age,self.end_age,self.gender,self.start_height,self.end_height,self.salary))
        self.craw_data()  # 数据抓取

    def query_age(self):
        '''
        年龄筛选
        :return:
        '''
        try:
            a = input('请输入期望对方年龄,如:25:')         # int类型的输入,input默认是str的
            age = int(a)
        except Exception as e:
            age = 25

        try:
            if 21 <= age <=30:
                self.start_age = 21
                self.end_age = 30
            elif 31<= age <=40:
                self.start_age = 31
                self.end_age = 40
            elif 41<=age<=50:
                self.start_age = 41
                self.end_age = 50
            else:
                self.start_age = 0
                self.end_age = 0
        except Exception as e:
            self.start_age = 24
            self.end_age = 26

    def query_sex(self):
        '''性别筛选'''
        try:
            sex = input('请输入期望对方性别,如:女:')  # 字符串的输入
        except Exception as e:
            sex = '女'

        try:
           if sex == '男':
               self.gender = 1
           else:
               self.gender = 2

        except Exception as e:
           self.gender = 2

    def query_height(self):
        '''身高筛选'''
        try:
            h = input('请输入期望对方身高,如:162:')
            height = int(h)
        except Exception as e:
            height = 0

            try:
                if 151 <= height <= 160:
                    self.start_height = 151
                    self.end_height = 160
                elif 161 <= height <= 170:
                    self.start_height = 161
                    self.end_height = 170
                elif 171 <= height <= 180:
                    self.start_height = 171
                    self.end_height = 180
                elif 181 <= height <= 190:
                    self.start_height = 181
                    self.end_height = 190
                else:
                    self.start_height = 0
                    self.end_height = 0
            except Exception as e:
                self.start_height = 0
                self.end_height = 0

    def query_money(self):
        '''待遇筛选'''
        try:
            m = input('请输入期望的对方月薪,如:8000:')
            money = int(m)
        except Exception as e:
            money = 0


        try:
            if 2000 <= money <5000:
               self.salary = 2
            elif 5000 <= money < 10000:
                self.salary = 3
            elif 10000 <= money <= 20000:
                self.salary = 4
            elif 20000 <= money :
                self.salary = 5
            else:
                self.salary = 0
        except Exception as e:
            self.salary = 0

    def store_info(self, nick,age,height,address,heart,education,img_url):
        '''
        存照片,与他们的内心独白
        '''
        if age < 22:
            tag = '22岁以下'
        elif 22 <= age < 28:
            tag = '22-28岁'
        elif 28 <= age < 32:
            tag = '28-32岁'
        elif 32 <= age:
            tag = '32岁以上'

        try:
            filename = '{}岁_身高{}_学历{}_{}_{}.jpg'.format(age,height,education, address, nick)

        except Exception as e:
            print(e)

        try:
            # 补全文件目录
            image_path = 'E:/store/pic/{}'.format(tag)
            # 判断文件夹是否存在。
            if not os.path.exists(image_path):
                os.makedirs(image_path)
                # os.mkdir(image_path)
                print(image_path + ' 创建成功')

            # 注意这里是写入图片，要用二进制格式写入。
            with open(image_path + '/' + filename, 'wb') as f:
                f.write(request.urlopen(img_url).read())

            txt_path = u'E:/store/txt'
            txt_name = u'内心独白.txt'
            # 判断文件夹是否存在。
            if not os.path.exists(txt_path):
                os.makedirs(txt_path)
                print(txt_path + ' 创建成功')

            # 写入txt文本
            with open(txt_path + '/' + txt_name, 'a') as f:
                f.write(heart)
        except Exception as e:
            print(e)

    def store_info_execl(self,nick,age,height,address,heart,education,img_url):
        person = []
        person.append(self.count)   #正好是数据条
        person.append(nick)
        person.append(u'女' if self.gender == 2 else u'男')
        person.append(age)
        person.append(height)
        person.append(address)
        person.append(education)
        person.append(heart)
        person.append(img_url)

        for j in range(len(person)):
            self.sheetInfo.write(self.count, j, person[j])

        # 工作薄对象.save(Excel文件名)：保存到Excel文件中
        self.f.save('我主良缘.xlsx')
        self.count += 1
        print('插入了{}条数据'.format(self.count))

    def parse_data(self,response):
        '''数据解析'''
        persons = json.loads(response).get('data').get('list')
        if persons is None:
            print('数据已经请求完毕')
            return print('888888888888888888888888888888888888888888888')

        for person in persons:
            nick = person.get('username')
            gender = person.get('gender')
            age = 2018 - int(person.get('birthdayyear'))
            address = person.get('city')
            heart = person.get('monolog')
            height = person.get('height')
            img_url = person.get('avatar')
            education = person.get('education')
            print(nick,age,height,address,heart,education)
            self.store_info(nick,age,height,address,heart,education,img_url)
            self.store_info_execl(nick,age,height,address,heart,education,img_url)

    def craw_data(self):
        '''数据抓取'''
        # 预防ip被封，模拟浏览器请求
        headers = {
            'Referer': 'http://www.lovewzly.com/jiaoyou.html',
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'
        }
        page = 1
        while True:

            query_data = {
                'page':page,
                'gender':self.gender,
                'start_age':self.start_age,
                'end_age':self.end_age,
                'strat_height':self.start_height,
                'end_height':self.end_height,
                'marry':self.marry,
                'salary':self.salary,
            }
            url = 'http://www.lovewzly.com/api/user/pc/list/search?' + parse.urlencode(query_data)
            # url = 'http://www.lovewzly.com/api/user/pc/list/search?'

            print(url)
            req = request.Request(url, headers=headers)
            response = request.urlopen(req).read()
            print(response)
            self.parse_data(response)
            page += 1


if __name__ == '__main__':
   wz = wzly()
   wz.query_data()