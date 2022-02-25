from filecmp import clear_cache
from selenium import webdriver
import time
import xlwt
import xlrd
from xlutils.copy import copy
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import json
import pyperclip
import re
import os
from urllib import parse,request
from datetime import datetime

#可自行更改的变量
keyword=''
keyword=input("Please input keyword: ")    #输入要搜索关键词
keyword_in_url=keyword.replace(" ","%20")
key_code=parse.quote(parse.quote(keyword)) #编码调整，如将 “冬奥会” 编码成 %E5%86%AC%E5%A5%A5%E4%BC%9A
path=os.path.expanduser('~/Desktop/').replace("\\","/")
file_path=path+'WeiboWeb_'+keyword+'.xls'
cookie_path=path+'WeiboCookies.txt'
login_url="https://weibo.com/login"
weibo_url="https://weibo.com"
first_time_login=False

class GetWeibo:
    def __init__(self):
        '''options = webdriver.ChromeOptions()        # 进入浏览器设置
        # 更换头部
        options.add_argument('User-Agent=Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.117 Safari/537.36')
        self.driver = webdriver.Chrome(options=options)'''
        self.driver=webdriver.Chrome()
        self.driver.maximize_window()
        if first_time_login:
            self.driver.get(login_url)
        else:
            self.driver.get(weibo_url) 
        self.url=""  
    
    def scan_code_login(self,w):   #扫码登录
        w.wait("xpath",'//*[@id="pl_login_form"]/div/div[1]/div/a[2]',1000)
        self.driver.find_element_by_xpath('//*[@id="pl_login_form"]/div/div[1]/div/a[2]').click()

    def getCookies(self,w,cookie_path):
        w.wait('xpath','//*[@id="homeWrap"]/div[1]/div',1000)
        dictCookies=self.driver.get_cookies()   #获取list
        jsonCookies=json.dumps(dictCookies)     #转换成字符串保存
        with open(cookie_path,'w') as f:
            f.write(jsonCookies)
        print('Cookies successfully obtained!')

    def add_cookies(self,cookie_path):   #向浏览器添加保存的cookies
        cookies = json.load(open(cookie_path, "rb"))
        for cookie in cookies:
            cookie_dict = {
                "domain": cookie.get('domain'),  
                'name': cookie.get('name'),
                'value': cookie.get('value'),
                "expires": "",
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'Secure': False 
            }
            self.driver.add_cookie(cookie_dict)
        time.sleep(3)
        self.driver.refresh()    #刷新网页 cookies才成功
        print("Successfully added cookies to browser!")
    
    def wait(self,method,element,seconds):   #等待直到某个元素出现，出现返回True，未出现返回False
        waits=WebDriverWait(self.driver,seconds)    #等待最多15秒
        elem=False   
        while 1:
            try:
                if method=='xpath':
                    elem = waits.until(EC.presence_of_element_located((By.XPATH, element)))   #bool
                elif method=='partial_link_text':
                    elem = waits.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, element)))
                elif method=='class_name':
                    elem = waits.until(EC.presence_of_element_located((By.CLASS_NAME, element)))
                break
            except:
                print("Timeout waiting.")
                break
        return elem

    def time_out(self,time_start,wait_seconds):
        dif_time=(datetime.now()-time_start).total_seconds()
        if dif_time>wait_seconds:
            return True
        return False

    def close_driver(self):   #退出浏览器和引擎驱动
        self.driver.delete_all_cookies()
        clear_cache()
        self.driver.close()
        self.driver.quit()

    def set_excel(self,file_path):
        wb=xlwt.Workbook()
        ws=wb.add_sheet(keyword)
        for i in range(8):
            ws.col(i).width=3500
        style = xlwt.easyxf('font: bold on, height 280;') 
        default_style = xlwt.easyxf('font: colour black, bold True, height 225; align: horiz centre')
        ws.write(0,0,'微博搜索'+keyword,style=style)
        ws.write(1,0,'用户名',style=default_style)
        ws.write(1,1,'转发数',style=default_style)
        ws.write(1,2,'评论数',style=default_style)
        ws.write(1,3,'点赞数',style=default_style)
        ws.write(1,4,'链接',style=default_style)
        ws.write(1,5,'全文',style=default_style)
        ws.write(1,6,'评论',style=default_style)
        ws.write(1,7,'截取状态',style=default_style)
        wb.save(file_path)
        print('Successfully written '+ file_path)

    def save_to_excel(self,dic,file_path):
        wb=xlrd.open_workbook(file_path,formatting_info=True)
        ws=wb.sheet_by_index(0)
        rowNum=ws.nrows
        newbook=copy(wb)
        newsheet=newbook.get_sheet(0)
        i=0
        for value in dic.values():
            newsheet.write(rowNum,i,value)
            i=i+1
        newbook.save(file_path)

    def save_detail_page_to_excel(self,row,full_content,st,bottom):
        wb=xlrd.open_workbook(file_path,formatting_info=True)
        newbook=copy(wb)
        newsheet=newbook.get_sheet(0)
        newsheet.write(row,5,full_content)
        newsheet.write(row,6,st)
        newsheet.write(row,7,bottom)
        newbook.save(file_path)

    def scroll(self,count):
        page=0
        while page<count:
            self.driver.execute_script('window.scrollTo(0,document.body.scrollHeight)')
            #time.sleep(3)
            page+=1 

    def search_topic(self):
        self.url="https://s.weibo.com/weibo?q="+keyword_in_url+"&xsort=hot&Refer=hotmore"
        self.driver.get(self.url)

    def crawl_post(self,w):
        w.wait("xpath",'//*[@id="pl_feedlist_index"]/div[1]/div[1]',1000)
        w.scroll(1)
        div=self.driver.find_elements_by_xpath('//*[@id="pl_feedlist_index"]/div[2]/div')   
        try:
            ActionChains(self.driver).move_to_element(self.driver.find_element_by_partial_link_text("第1页 ")).perform()
            time.sleep(1)
            page_list=self.driver.find_elements_by_xpath('//*[@id="pl_feedlist_index"]/div[3]/div/span/ul/li')  
        except:
            page_list=[""]
        dic={}                          #存储数据的字典
        url_ls=[]                       #存储链接的列表
        comment_count_ls=[]             #存储评论数列表
        target_page=len(page_list)      #要爬取的页数
        post_count=0
        real_count=0
        time_start=datetime.now()

        for page_count in range(target_page): 
            print("\nGetting info for page {}/{}".format(page_count+1,target_page))   #展示进度
            for index in div:
                try:   
                    #获取帖子用户名 转赞评数
                    dic['username']=index.find_element_by_xpath('./div/div[1]/div[2]/div[1]/div[2]/a').text                  
                    share_count=re.sub("[^0-9]","",index.find_element_by_xpath('./div/div[2]/ul/li[1]/a').text)
                    comment_count=re.sub("[^0-9]","",index.find_element_by_xpath('./div/div[2]/ul/li[2]/a').text)
                    like_count=re.sub("[^0-9]","",index.find_element_by_xpath('./div/div[2]/ul/li[3]/a/button/span[2]').text)
                    if share_count=="": share_count='0'
                    if comment_count=="": comment_count='0'
                    if like_count=="": like_count='0'
                    dic['share_count']=share_count
                    dic['comment_count']=comment_count
                    dic['like_count']=like_count
                    comment_count_ls.append(comment_count)

                    #获取帖子链接
                    time.sleep(0.5)
                    ActionChains(self.driver).move_to_element(index.find_element_by_xpath('./div/div[1]/div[2]/div[1]/div[1]/a/i')).click().perform()
                    time.sleep(0.5)  
                    ActionChains(self.driver).move_to_element(index.find_element_by_xpath('./div/div[1]/div[2]/div[1]/div[1]/ul/li[4]/a')).click().perform()
                    time.sleep(1)  
                    tmpUrl=pyperclip.paste().replace("?refer_flag=1001030103_","")
                    dic['url']=tmpUrl
                    url_ls.append(tmpUrl)
                    pyperclip.copy('')
                    post_count+=1
                    print("Task {} accomplished.".format(post_count),end="   ")  
                    print("Time taken: {:.2f} minutes".format((datetime.now()-time_start).total_seconds()/60))       
                    w.save_to_excel(dic,file_path)                        #存储数据

                except:
                    print("Getting task exception.")
              
            if not page_count+1 == target_page:
                self.driver.get("https://s.weibo.com/weibo?q="+keyword_in_url+"&xsort=hot&Refer=hotmore&page="+str(page_count+2))        #翻页
                w.scroll(1)
                div=self.driver.find_elements_by_xpath('//*[@id="pl_feedlist_index"]/div[2]/div') 
        
        #通过链接获取每个帖子的评论
        print('Getting post comments ......')
        post_count=len(url_ls)
        row_ls=[]
        timeout_url_ls=[]
        timeout_row_ls=[]
        for i in range(2,post_count+2):
            row_ls.append(i)
        while 1:
            for row,url,comment_count in zip(row_ls,url_ls,comment_count_ls):    
                print("\nGetting comment for task {}/{}".format(row-1,post_count),end="   ")
                print("Time taken: {:.2f} minutes".format((datetime.now()-time_start).total_seconds()/60),end="   ")
                full_content,st,bottom=w.get_detail_page(url,comment_count,w)
                w.save_detail_page_to_excel(row,full_content,st,bottom)
                if bottom=="超时限获取评论" or bottom=="超话社区" or bottom=="超时限获取链接":
                    timeout_url_ls.append(url)
                    timeout_row_ls.append(row)
                else:
                    real_count+=1

            print('\nTasks completed: {}/{}'.format(real_count,post_count))  
            if len(timeout_url_ls)==0 and len(timeout_row_ls)==0:
                break
            #break
            retrying_task=""
            retrying_task=input("Retrying uncompleted task? (y/n) ")   #由用户决定是否重试爬取失败/不完整的任务（不影响原有数据）
            if retrying_task=="n":
                break
            if w.time_out(time_start,post_count*180):
                break
            
            url_ls=timeout_url_ls
            row_ls=timeout_row_ls
            print("Retrying task ......")
            
        print('\nTasks completed: {}/{}'.format(real_count,post_count))   
        print("Total time taken: {:.2f} minutes".format((datetime.now()-time_start).total_seconds()/60))

    def get_detail_page(self,url,comment_count,w):   
        flag=1
        bottom=""
        try:
            self.driver.get(url)   
        except:
            bottom="超时限获取链接"
            full_content=""
            st=""
            return full_content,st,bottom

        time1=datetime.now()
        post_present=w.wait("class_name",'detail_wbtext_4CRf9',10)
        if post_present==False:
            bottom="该微博不存在"
            full_content="-"
            st="-"
            return full_content,st,bottom
        full_content=self.driver.find_element_by_class_name('detail_wbtext_4CRf9').text
        st=""   
        if comment_count=="0":
            st="-"
            bottom="-" 
            return full_content,st,bottom
            
        i=0
        while 1:
            try:
                self.driver.find_element_by_class_name('Bottom_text_1kFLe').is_displayed()  #已下滑到评论区底部 
                bottom=self.driver.find_element_by_class_name('Bottom_text_1kFLe').text
                break
            except:
                w.scroll(1)
                time.sleep(1)
            #if w.time_out(time1,int(int(comment_count)/10)):
            if w.time_out(time1,5):
                if bottom=="":
                    bottom="超时限获取评论"
                flag=0
                break
        if flag:
            print('Reload completed.',end="")
            
        try:
            self.driver.find_element_by_xpath('//*[@id="app"]/div[1]/div[2]/div[2]/main/div[1]/div/div[2]/div[2]/div[3]/div[2]/div').is_displayed()
            bottom="博主已开启评论精选"
        except:
            print("",end="")
        comment_ls=self.driver.find_elements_by_xpath('//*[@id="scroller"]/div[1]/div')   
        for cmt in comment_ls:
            comment=cmt.find_element_by_xpath('./div/div/div/div[1]/div[2]/div[1]').text.replace("\n","")
            if comment=="":
                continue
            i+=1
            st+=str(i)+". "+comment+"\n" 
        if st.startswith("1. 超话社区:你好") and st.endswith("超话"):
            bottom="超话社区"
        return full_content,st,bottom

    
if __name__=="__main__":
    w=GetWeibo()
    if first_time_login:
        w.scan_code_login(w)
        w.getCookies(w,cookie_path)
    w.add_cookies(cookie_path)
    w.set_excel(file_path)
    w.search_topic()
    w.crawl_post(w)
    print('Data successfully saved.')
    w.close_driver()

#python path=D:\yh\others\Project\ScrapWeibo
#weiboWeb4.py


