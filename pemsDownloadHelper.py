#!/usr/bin/env python
# coding: utf-8
# Author: Lumos
# Description: to batch download dataset from PeMS

# In[1]:


import pandas as pd
import datetime
from datetime import timedelta
import os
from selenium import webdriver
from selenium.webdriver.support.select import Select
import time
import glob
import shutil


# In[2]:


## 定义了一个类，用来处理表格数据，然后生成自动化过程所需要的4类数据：含有每个vds数据的网址、开始时间、结束时间、事故id
class pemsData:
    ##类的一些可访问属性
    interval = 3  #事故的前后时间间隔，默认为3小时
    dataset = {}  #用来存放生成数据的字典。
    file_path = ''  #读取表格文件的路径。

    ## 不可访问属性
    __fileDir = ''  #文件所在目录。用于后面兼容相对路径和绝对路径。不用管。

    ## 类的初始化，需要输入文件路径，可以输入事故的前后时间间隔，默认为3小时。
    def __init__(self, file_path, interval=3):
        self.file_path = file_path
        self.interval = interval
        self.__fileDir = os.getcwd()
        self.__creat_data(file_path, interval)

    ## 重新设置时间间隔，当初始化的时候输错或者想试下其他时间间隔的数据，可以调用这个函数。
    def reset_interval(self, interval):
        self.interval = interval
        self.__creat_data(self.file_path, self.interval)

    ## 重新设置文件读取路径，想要更换读取的文件可以调用这个函数。
    def reset_file(self, file_path):
        self.file_path = file_path
        self.__creat_data(self.file_path, self.interval)

    ## 数据生成汇总
    def __creat_data(self, file_path, interval):
        ## 重置数据表
        self.dataset = {}

        ## 获取文件里所有sheet里的数据，以字典的形式储存：{sheet名：该sheet里的dataFrame}
        ## 这个try,except就是上面说的用来兼容相对路径和绝对路径的
        try:
            io = pd.io.excel.ExcelFile(file_path)
            sheet_names = pd.ExcelFile(file_path).sheet_names
        except:
            io = pd.io.excel.ExcelFile(self.__fileDir + '\\' + file_path)
            sheet_names = pd.ExcelFile(self.__fileDir + '\\' +
                                       file_path).sheet_names
        all_data = pd.read_excel(io, sheet_name=sheet_names)

        ## 对每个sheet里的数据循环遍历
        for sheet in sheet_names:
            data = all_data[sheet]

            vds_column = data['match_VDSID']
            date_column = data['COLLISION_DATE']
            time_column = data['COLLISION_TIME']
            case_column = data['match_CASEID']

            case_list = case_column.apply(str).values.tolist()  ## 将事故id列处理成列表
            url_list = self.__getUrls(vds_column)  ## 将vds_id列处理成带有vds的网址的列表
            from_list, to_list = self.__getDate(
                date_column, time_column, interval)  ##将日期和时间列处理成一定间隔时间的起止时间列表

            ## 将上面处理的数据添加到dataset里
            self.dataset[sheet] = {
                'vds_urls': url_list,
                'from_list': from_list,
                'to_list': to_list,
                'case_list': case_list
            }

    ## 将vds_id列处理成带有vds的网址的列表
    def __getUrls(self, vds_column):
        vds_list = vds_column.values.tolist()
        vds_urls = []
        for i in vds_list:
            vds_url = 'https://pems.dot.ca.gov/?dnode=VDS&content=loops&tab=det_timeseries&station_id=' + str(
                i)
            vds_urls.append(vds_url)
        return vds_urls

    ## 将日期和时间列处理成一定间隔时间的起止时间列表
    def __getDate(self, date_column, time_column, interval):
        date_list = date_column.dt.strftime('%m/%d/%Y').apply(
            str).values.tolist()
        time_list = time_column.values.tolist()
        f_list = []
        t_list = []
        for i in range(len(date_list)):
            temp = "{:0>4d}".format(time_list[i])
            time_format = temp[:2] + ':' + temp[2:]
            combined = date_list[i] + " " + time_format
            d = datetime.datetime.strptime(combined, '%m/%d/%Y %H:%M')
            from_d = (d - timedelta(hours=interval)).strftime('%m/%d/%Y %H:%M')
            to_d = (d + timedelta(hours=interval)).strftime('%m/%d/%Y %H:%M')
            f_list.append(from_d)
            t_list.append(to_d)

        return f_list, t_list


# In[3]:


##定义了一个类，用来处理表格数据，然后生成自动化过程所需要的2类数据：直接下载链接、事故id（用于重命名）。下面的注释基本同上
class pemsUrlData():
    ##类的一些可访问属性
    interval = 3  #事故的前后时间间隔，默认为3小时
    dataset = {}  #用来存放生成数据的字典。
    file_path = ''  #读取表格文件的路径。

    ## 不可访问属性
    __fileDir = ''  #文件所在目录。用于后面兼容相对路径和绝对路径。不用管。

    def __init__(self, file_path, interval=3):
        self.file_path = file_path
        self.interval = interval
        self.__fileDir = os.getcwd()
        self.__creat_urldata(file_path, interval)

    def reset_interval(self, interval):
        self.interval = interval
        self.__creat_urldata(self.file_path, self.interval)

    def reset_file(self, file_path):
        self.file_path = file_path
        self.__creat_urldata(self.file_path, self.interval)

    def __creat_urldata(self, file_path, interval):
        self.dataset = {}
        
        ## 这个try,except就是上面说的用来兼容相对路径和绝对路径的
        try:
            io = pd.io.excel.ExcelFile(file_path)
            sheet_names = pd.ExcelFile(file_path).sheet_names
        except:
            io = pd.io.excel.ExcelFile(self.__fileDir + '\\' + file_path)
            sheet_names = pd.ExcelFile(self.__fileDir + '\\' +
                                       file_path).sheet_names
        all_data = pd.read_excel(io, sheet_name=sheet_names)

        for sheet in sheet_names:
            data = all_data[sheet]

            vds_column = data['match_VDSID']
            date_column = data['COLLISION_DATE']
            time_column = data['COLLISION_TIME']
            case_column = data['match_CASEID']

            case_list = case_column.apply(str).values.tolist()
            vds_list = vds_column.apply(str).values.tolist()
            from_list, to_list = self.__getDate(date_column, time_column,
                                                interval)
            urls = []
            for i in range(len(vds_list)):
                url = 'https://pems.dot.ca.gov/?report_form=1&dnode=VDS&content=loops&tab=det_timeseries&export=xls&station_id=' + vds_list[
                    i] + '&s_time_id=' + from_list[i] + '&e_time_id=' + to_list[
                        i] + '&tod=all&tod_from=0&tod_to=0&dow_0=on&dow_1=on&dow_2=on&dow_3=on&dow_4=on&dow_5=on&dow_6=on&holidays=on&q=flow&q2=speed&gn=5min&agg=on&lane1=on&lane2=on&lane3=on&lane4=on'
                urls.append(url)

            self.dataset[sheet] = {'urls': urls, 'case_list': case_list}

    def __getDate(self, date_column, time_column, interval):
        date_list = date_column.dt.strftime('%m/%d/%Y').apply(
            str).values.tolist()
        time_list = time_column.values.tolist()
        combined_list = []
        f_list = []
        t_list = []
        for i in range(len(date_list)):
            temp = "{:0>4d}".format(time_list[i])
            time_format = temp[:2] + ':' + temp[2:]
            combined = date_list[i] + " " + time_format
            d = datetime.datetime.strptime(combined, '%m/%d/%Y %H:%M')
            from_d = str((d - timedelta(hours=interval)).timestamp())
            to_d = str((d + timedelta(hours=interval)).timestamp())
            f_list.append(from_d)
            t_list.append(to_d)

        return f_list, t_list


# In[4]:


## 定义了一个自动化下载的类，对应上面的两个类有两种下载方法：1.通过下载链接直接下载；2.通过模拟输入填写表单进行下载
class pemsSelenium:
    
    url = 'https://pems.dot.ca.gov/' # 官方网址
    username = '' # 用户名
    password = '' # 密码
    download_path = '' # 下载路径
    driver = '' # 浏览器驱动，不用管。

    ## 模拟浏览器初始化并登录官网
    def __init__(self, username, password, download_path):
        ## 初始化模拟浏览器参数
        options = webdriver.ChromeOptions()
        prefs = {
            'profile.default_content_settings.popups': 0, # 设置不弹出下载文件保存地址的窗口
            'download.default_directory': download_path # 设置浏览器默认下载地址
        }
        options.add_experimental_option('prefs', prefs)
        self.driver = webdriver.Chrome(options=options)
        
        self.driver.get(self.url)  # 进入官网
        self.__login(username, password) # 使用账号密码登录
        
        ##设置可访问属性
        self.username = username 
        self.password = password
        self.download_path = download_path
    
    ## 私有的登录函数，被上面初始化时调用
    def __login(self, username, password):

        user_input = self.driver.find_element_by_id('username') #通过id找到填写用户名的输入框
        user_input.send_keys(username) #向输入框里输入用户名

        word_input = self.driver.find_element_by_id('password') #通过id找到填写密码的输入框
        word_input.send_keys(password) #向输入框里输入密码

        loginButton = self.driver.find_element_by_xpath(
            '//*[@id="top_banner"]/table/tbody/tr/td[2]/form/fieldset/div[3]/input' #通过xpath找到登录按钮
        )
        loginButton.click() #点击登录按钮

    ## 通过下载链接直接下载
    def download_by_Urls(self, dataset):

        print('开始下载...')
        count = 0
        for dirName in dataset.keys():
            os.chdir(self.download_path) 
            os.makedirs(dirName) #新建用当前sheet命名的文件夹

            data = dataset[dirName] 
            urls = data['urls']
            case_list = data['case_list']

            for i in range(len(urls)):
                self.driver.get(urls[i]) #通过下载链接下载

                time.sleep(2)

                os.rename("pems_output.xlsx", case_list[i] + '.xlsx') #重命名
                count += 1
            
            ## 文件移动到刚才新建的文件夹
            download_files = glob.glob('*.xlsx')
            for file in download_files:
                shutil.move(file, dirName)

        print(f'下载完毕! 共下载了个{count}文件')
    
    ## 通过自动填写表单下载，注释基本同上
    def download_by_Autofill(self, dataset):

        print('开始下载...')
        count = 0

        for dirName in dataset.keys():
            os.chdir(self.download_path)
            os.makedirs(dirName)

            data = dataset[dirName]
            vds_urls = data['vds_urls']
            from_list = data['from_list']
            to_list = data['to_list']
            case_list = data['case_list']

            for i in range(len(vds_urls)):
                self.driver.get(vds_urls[i])
                time_from = self.driver.find_element_by_id('s_time_id_f') #通过id找到开始时间的输入框
                time_to = self.driver.find_element_by_id('e_time_id_f') #通过id找到结束时间的输入框

                time_from.clear() #清除两个输入框里的默认值
                time_to.clear() 

                time_from.send_keys(from_list[i])  #向两个输入框输入起止日期
                time_to.send_keys(to_list[i])

                q2_select = Select(self.driver.find_element_by_id('q2')) #找到q2下拉框
                q2_select.select_by_value('speed') #在q2下拉框中找到speed

                gn_select = Select(self.driver.find_element_by_id('gn')) 
                gn_select.select_by_value('5min')

                dl = self.driver.find_element_by_xpath(
                    '//*[@id="submitButtonContainer"]/input[4]')
                dl.click() #点击下载按钮

                time.sleep(3) #休眠3秒钟

                os.rename("pems_output.xlsx", case_list[i] + '.xlsx')
                count += 1

            download_files = glob.glob('*.xlsx')
            for file in download_files:
                shutil.move(file, dirName)

        print(f'下载完毕! 共下载了个{count}文件')


# In[5]:


if __name__ == '__main__':
    ## 需要修改的信息
    file_path = 'test.xlsx'  #要读取的文件路径
    username = 'lovegoodlumos@gmail.com'  #用户名
    password = ')+rL5javak>'  #密码
    download_path = 'f:\\pemsDownload'  #下载文件的存放路径。一定要检查保证文件夹里是空的！！！最好新建一个文件夹作为存储路径。
    
    ## 实例化一个pemsSelenium类，然后浏览器会自动打开并登录，开启自动化之旅~
    ps = pemsSelenium(username, password, download_path) 
    
    ## 通过下载链接直接下载，省去填写表单的环节，速度较快
    ped = pemsUrlData(file_path) #实例化一个pemsUrlData类，对源文件中的数据处理，通过’ped.dataset‘可获取下一步需要的数据。
    ps.download_by_Urls(ped.dataset) #调用pemsSelenium类中的download_by_Urls方法进行下载
    
#     ## 通过模拟输入下载，速度较慢，但是吧，看着浏览器在那里自己操作，有一种很治愈的感觉。上面那种方式出问题的话也可以试试这个。
#     ped = pemsData(file_path) #实例化一个pemsData类，对源文件中的数据处理，通过’ped.dataset‘可获取下一步需要的数据。
#     ps.download_by_Autofill(ped.dataset) #调用pemsSelenium类中的download_by_Urls方法进行下载
    
    ## 下载完退出
    ps.driver.quit() 


    

