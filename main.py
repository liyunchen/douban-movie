# -*- coding: utf-8 -*-

import requests
from lxml import etree
import xlrd
import xlwt
import time
import random
from xlutils.copy import copy
import matplotlib as mpl
from matplotlib import pyplot as plt
import pandas as pd

# matplotlib中文显示
plt.rcParams['font.family'] = ['sans-serif']
plt.rcParams['font.sans-serif'] = ['SimHei']

#获取ip代理
def getip():
    iplist = []
    with open("ip代理.txt") as f:
        iplist = f.readlines()

    proxy= iplist[random.randint(0,len(iplist)-1)]
    proxy = proxy.replace("\n","")
    proxies={
        'http':'http://'+str(proxy),
        #'https':'https://'+str(proxy),
    }
    return proxies

headers = {
            'Host':'movie.douban.com',
            'user-agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3947.100 Safari/537.36',
            'cookie':'bid=uVCOdCZRTrM; douban-fav-remind=1; __utmz=30149280.1603808051.2.2.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not%20provided); __gads=ID=7ca757265e2366c5-22ded2176ac40059:T=1603808052:RT=1603808052:S=ALNI_MYZsGZJ8XXb1oU4zxzpMzGdK61LFA; __utmc=30149280; __utmz=223695111.1612839506.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); __utmc=223695111; dbcl2="165593539:LvLaPIrgug0"; ck=ZbYm; push_doumail_num=0; push_noty_num=0; __utmv=30149280.16559; ll="118288"; __yadk_uid=DnUc7ftXIqYlQ8RY6pYmLuNPqYp5SFzc; _vwo_uuid_v2=D7ED984782737D7813CC0049180E68C43|1b36a9232bbbe34ac958167d5bdb9a27; ct=y; __utma=30149280.1867171825.1603588354.1612893804.1612958182.6; __utma=223695111.788421403.1612839506.1612893804.1612958182.4; _pk_id.100001.4cf6=e2e8bde436a03ad7.1612839506.4.1612961476.1612894313.',
            #'accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
            #'accept-encoding': 'gzip, deflate, br',
            #'accept-language': 'zh-CN,zh;q=0.9',
            #'upgrade-insecure-requests': '1',
            #'referer':'',
        }

# 写入execl
def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿

# 初始化execl表
def initexcel(filename):
    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('sheet1')
    workbook.save(str(filename)+'.xls')
    ##写入表头
    value1 = [["用户", "影评"]]
    book_name_xls = str(filename)+'.xls'
    write_excel_xls_append(book_name_xls, value1)

###获取用户观看过的电影
def getmovieByUser(filename,u_page,userid,username):

        url="https://movie.douban.com/people/"+str(userid)+"/collect?start=0&sort=time&rating=all&filter=all&mode=grid"
        r = requests.get(url,headers=headers)
        r.encoding = 'utf8'
        s = (r.content)
        selector = etree.HTML(s)

        try:
            page = (selector.xpath('/html/body/div[3]/div[1]/div[2]/div[1]/div[3]/a[10]/text()'))[0]
            page = int(page)
            if page>2:
                page=2
            #moivelist = []
            ###遍历所有电影
            for i in range(0,int(page)):
                print("用户列表页数=" + str(u_page) + "----，用户名=" + str(username) + "----，用户电影列表页数=" + str(i))

                url = "https://movie.douban.com/people/" + str(userid) + "/collect?start="+str(i*15)+"&sort=time&rating=all&filter=all&mode=grid"

                r = requests.get(url, headers=headers)
                r.encoding = 'utf8'
                s = (r.content)
                selector = etree.HTML(s)

                for item in selector.xpath('//*[@class="grid-view"]/div[@class="item"]'):
                    text1 = item.xpath('.//*[@class="title"]/a/em/text()')
                    text2 = item.xpath('.//*[@class="title"]/a/text()')
                    text1 = (text1[0]).replace(" ","")
                    text2 = (text2[1]).replace(" ","").replace("\n","")

                    book_name_xls = str(filename)+'.xls'
                    data = []
                    #print(text1+text2)
                    data.append([str(username),str(text1+text2)])
                    write_excel_xls_append(book_name_xls, data)

                time.sleep(random.randint(1,3))
        except:
            pass


##获取用户列表
def get_alluser():
    page=218
    ###217 可以，218不可以

    for i in range(0,page):
        fname = '豆瓣'+str(i)
        initexcel(fname)
        url="https://movie.douban.com/subject/24733428/reviews?start="+str(i*20)
        r = requests.get(url,  headers=headers)
        r.encoding = 'utf8'
        s = (r.content)
        selector = etree.HTML(s)
        print(str(fname))

        for item in selector.xpath('//*[@class="review-list  "]/div'):

            userid = (item.xpath('.//*[@class="main-hd"]/a[2]/@href'))[0].replace("https://www.douban.com/people/","").replace("/","")
            username = (item.xpath('.//*[@class="main-hd"]/a[2]/text()'))[0]
            #print(userid)
            #print(username)
            getmovieByUser(fname,i,userid,username)
            time.sleep(random.randint(1,2))

        time.sleep(random.randint(2,3))

def read_excel():
    # 打开workbook
    data = xlrd.open_workbook('豆瓣.xls')
    # 获取sheet页
    table = data.sheet_by_name('sheet1')
    # 已有内容的行数和列数
    nrows = table.nrows
    datalist=[]
    for row in range(nrows):
        temp_list = table.row_values(row)
        if temp_list[0] != "用户" and temp_list[1] != "影评":
            data = []
            data.append([str(temp_list[0]), str(temp_list[1])])
            datalist.append(data)

    return datalist

###分析1：电影观看次数排行
def analysis1():
    dict ={}
    ###从excel读取数据
    movie_data = read_excel()
    for i in range(0, len(movie_data)):
        key = str(movie_data[i][0][1])
        try:
            dict[key] = dict[key] +1
        except:
            dict[key]=1
    ###从小到大排序
    dict = sorted(dict.items(), key=lambda kv: (kv[1], kv[0]))
    name=[]
    num=[]
    for i in range(len(dict)-1,len(dict)-16,-1):
        print(dict[i])
        name.append(((dict[i][0]).split("/"))[0])
        num.append(dict[i][1])

    plt.figure(figsize=(16, 9))
    plt.title('电影观看次数排行(高->低)')
    plt.bar(name, num, facecolor='lightskyblue', edgecolor='white')
    plt.savefig('电影观看次数排行.png')

###分析2：用户画像（用户观影相同率最高）
def analysis2():
    dict = {}
    ###从excel读取数据
    movie_data = read_excel()

    userlist=[]
    for i in range(0, len(movie_data)):
        user = str(movie_data[i][0][0])
        moive = (str(movie_data[i][0][1]).split("/"))[0]
        #print(user)
        #print(moive)

        try:
            dict[user] = dict[user]+","+str(moive)
        except:
            dict[user] =str(moive)
            userlist.append(user)

    num_dict={}
    # 待画像用户(取第一个）
    flag_user=userlist[0]
    print(flag_user)
    movies = (dict[flag_user]).split(",")
    for i in range(0,len(userlist)):
        #判断是否是待画像用户
        if flag_user != userlist[i]:
            num_dict[userlist[i]]=0
            #待画像用户的所有电影
            for j in range(0,len(movies)):
                #判断当前用户与待画像用户共同电影个数
                if movies[j] in dict[userlist[i]]:
                    # 相同加1
                    num_dict[userlist[i]] = num_dict[userlist[i]]+1
    ###从小到大排序
    num_dict = sorted(num_dict.items(), key=lambda kv: (kv[1], kv[0]))
    #用户名称
    username = []
    #观看相同电影次数
    num = []
    for i in range(len(num_dict) - 1, len(num_dict) - 9, -1):
        username.append(num_dict[i][0])
        num.append(num_dict[i][1])

    plt.figure(figsize=(25, 9))
    plt.title('用户画像（用户观影相同率最高）')
    plt.scatter(username, num, color='r')
    plt.plot(username, num)
    plt.savefig('用户画像（用户观影相同率最高）.png')

###分析3：用户之间进行电影推荐（与其他用户同时被观看过）
def analysis3():
    dict = {}
    ###从excel读取数据
    movie_data = read_excel()

    userlist=[]
    for i in range(0, len(movie_data)):
        user = str(movie_data[i][0][0])
        moive = (str(movie_data[i][0][1]).split("/"))[0]
        #print(user)
        #print(moive)

        try:
            dict[user] = dict[user]+","+str(moive)
        except:
            dict[user] =str(moive)
            userlist.append(user)

    num_dict={}
    # 待画像用户(取第2个）
    flag_user=userlist[0]
    print(flag_user)
    movies = (dict[flag_user]).split(",")
    for i in range(0,len(userlist)):
        #判断是否是待画像用户
        if flag_user != userlist[i]:
            num_dict[userlist[i]]=0
            #待画像用户的所有电影
            for j in range(0,len(movies)):
                #判断当前用户与待画像用户共同电影个数
                if movies[j] in dict[userlist[i]]:
                    # 相同加1
                    num_dict[userlist[i]] = num_dict[userlist[i]]+1
    ###从小到大排序
    num_dict = sorted(num_dict.items(), key=lambda kv: (kv[1], kv[0]))

    # 去重（用户与观影率最高的用户两者之间重复的电影去掉）
    user_movies = dict[flag_user]
    new_movies = dict[num_dict[len(num_dict)-1][0]].split(",")
    for i in range(0,len(new_movies)):
        if new_movies[i] not in user_movies:
            print("给用户（"+str(flag_user)+"）推荐电影："+str(new_movies[i]))

###分析4：电影之间进行电影推荐（与其他电影同时被观看过）
def analysis4():
    dict = {}
    ###从excel读取数据
    movie_data = read_excel()

    userlist=[]
    for i in range(0, len(movie_data)):
        user = str(movie_data[i][0][0])
        moive = (str(movie_data[i][0][1]).split("/"))[0]
        try:
            dict[user] = dict[user]+","+str(moive)
        except:
            dict[user] =str(moive)
            userlist.append(user)

    movie_list=[]
    # 待获取推荐的电影
    flag_movie = "送你一朵小红花"
    for i in range(0,len(userlist)):
        if flag_movie in dict[userlist[i]]:
             moives = dict[userlist[i]].split(",")
             for j in range(0,len(moives)):
                 if moives[j] != flag_movie:
                     movie_list.append(moives[j])

    data_dict = {}
    for key in movie_list:
        data_dict[key] = data_dict.get(key, 0) + 1

    ###从小到大排序
    data_dict = sorted(data_dict.items(), key=lambda kv: (kv[1], kv[0]))
    for i in range(len(data_dict) - 1, len(data_dict) -16, -1):
            print("根据电影"+str(flag_movie)+"]推荐："+str(data_dict[i][0]))



###分析1：电影观看次数排行
#analysis1()
###分析2：用户画像（用户观影相同率最高）
#analysis2()
###分析3：用户之间进行电影推荐（与其他用户同时被观看过）
#analysis3()
###分析4：电影之间进行电影推荐（与其他电影同时被观看过）
#analysis4()


#initexcel('豆瓣')
#get_alluser()


