# 数据爬取-selenium
import os
import ssl

import psutils
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
import re
import time
import traceback
import requests
import json
from urllib import parse,request
import CtlWX
from tqdm import tqdm
from progress.bar import Bar
import ddddocr
from PIL import Image
import WXreply
import CtlYDBG

import ast
from graphviz import Digraph

# 集思录爬虫类
class EMOS_Auto:
    # 初始化变量
    # 传入is_ssl的值,1为使用ssl认证,0为禁用ssl认证
    #  chromedriver的文件位置
    EmosSheetList = []  #已处理过的工单列表

    def __init__(self):
        # 读取已处理过的EMOS工单
        with open('EmosSheetList.txt', 'r') as f:
            for line in f.readlines():
                if len(line) > 10:
                    self.EmosSheetList.append(line.replace("\n", "").rstrip().lstrip())
        print("已处理工单列表:")
        print(self.EmosSheetList)

        # 设置浏览器驱动方法
        self.browser = webdriver.Ie()
        # 隐式等待：针对全局元素生效；（讲这个）
        self.browser.implicitly_wait(20)  # 一般情况下设置30秒
        self.browser.set_script_timeout(20)
        self.browser.set_page_load_timeout(20)

    # ----------函数：保存日志-----------
    def Log(self,log):
        print(log)
        # 保存log
        with open("Log-" + time.strftime('%Y%m%d') + ".txt", "a+",encoding="utf-8") as f:
            f.write(log+"\n")

    #将emos工单内容记录下来
    def LogEmos(self,log):
        # 保存log
        with open("Emos.txt", "a+",encoding="utf-8") as f:
            f.write(log+"\n")

    def CloseBrowser(self):
        try:
            if self.browser != None:
                self.browser.close()
        except Exception as e:
            print("CloseBrowser:"+'str(Exception):\t' + str(Exception) + '\n')

    def ResetBrowser(self):
        try:
            # 杀掉原有进程
            pids = psutil.process_iter()
            for pid in pids:
                # Log(pid.name() + "  " + str(pid.pid))
                if (pid.name() == "IEDriverServer.exe"):
                    self.Log("关闭IE driver： PID=" + str(pid.pid))
                    os.kill(pid.pid, 9)  # 杀死进程
                elif (pid.name() == "iexplore.exe"):
                    self.Log("关闭IE： PID=" + str(pid.pid))
                    os.kill(pid.pid, 9)  # 杀死进程
            # 设置浏览器驱动方法
            self.browser = webdriver.Ie()
            # 隐式等待：针对全局元素生效；（讲这个）
            self.browser.implicitly_wait(20)  # 一般情况下设置30秒
            self.browser.set_script_timeout(20)
            self.browser.set_page_load_timeout(20)

        except Exception as e:
            self.Log("ResetBrowser:"+'str(Exception):\t' + str(Exception) + '\n')

    def Login_get_cookies(self):
        try:
            """获取cookies保存为txt"""
            self.browser.get("http://emip.gmcc.net")
            '''
            WebDriverWait(self.browser, timeout=20).until(
                EC.presence_of_element_located((By.ID, 'username')))
            self.browser.maximize_window()
            radio_btns = self.browser.find_elements(By.XPATH, '//input[@class="sel" and @type="radio"]')
            radio_btns[2].click()
            time.sleep(1)
            self.browser.find_element(By.ID, "username").send_keys("xujiajun2")
            time.sleep(0.2)

            iCnt = 10
            while iCnt > 0:
                iCnt -= 1
                print("输入账号密码" + str(iCnt))
                radio_btns = self.browser.find_elements(By.XPATH, '//input[@class="sel" and @type="radio"]')
                radio_btns[2].click()
                time.sleep(1)
                self.browser.find_element(By.ID,"password").send_keys("Gmcc!234")
                time.sleep(0.2)

                self.browser.save_screenshot('loginPage.png')
                element = self.browser.find_element(by=By.XPATH, value="//img[@id='captcha_id']")
                #print(element.location)  # 打印元素坐标
                #print(element.size)  # 打印元素大小
                left = element.location['x']
                top = element.location['y']
                right = element.location['x'] + element.size['width']
                bottom = element.location['y'] + element.size['height']
                im = Image.open('loginPage.png')
                im = im.crop((left, top, right, bottom))
                im.save('code.png')
                #with open("code.jpg", 'rb') as f:
                #    image = f.read()
                ocr = ddddocr.DdddOcr(show_ad=False)
                res = ocr.classification(im)
                self.browser.find_element(By.ID, "j_captcha_response").send_keys(res)
                time.sleep(1)
                ButCheck = self.browser.find_element(By.XPATH, "//*[@id=\"eoms\"]/tbody/tr[4]/td[2]/input[1]")
                self.browser.execute_script("arguments[0].click();", ButCheck)
                self.Log("输入验证码:"+res)
                time.sleep(2)
                #检查是否存在 报错元素
                loginStatus = self.browser.find_elements(By.ID, "status")
                if len(loginStatus) > 0 :
                    self.Log(loginStatus[0].text)
                else:#验证码输入正确
                    break
            '''

            # 等待主页页面加载完成
            print("等待主页加载...")
            WebDriverWait(self.browser, timeout=2000).until(
                EC.presence_of_element_located((By.ID, 'mxzyywId1')))
            print("主页加载完成")
            #input("请登陆后按Enter")

            # print("Cookies: ")
            # print(self.browser.get_cookies())
            # cookie = {}
            # for i in self.browser.get_cookies():
            #     cookie[i["name"]] = i["value"]
            # with open("cookies.txt", "w") as f:
            #     f.write(json.dumps(cookie))

            return True

        except Exception as e:
            strlog = ('str(Exception):\t' + str(Exception) + '\n')
            strlog += ('str(e):\t\t' + str(e) + '\n')
            strlog += ('repr(e):\t' + repr(e) + '\n')
            # strlog += ('e.message:\t'+ e.message+'\n')
            strlog += ('traceback.print_exc():' + str(traceback.print_exc()) + '\n')
            strlog += ('traceback.format_exc():\n' + str(traceback.format_exc()) + '\n')
            self.Log(strlog)
            return False

    #刷新工单列表
    def get_emos_list(self):
        try:
            self.Log(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())) + "打开工单列表：")
            while len(self.browser.window_handles) > 1:
                self.browser.switch_to.window(self.browser.window_handles[-1])
                strTemp = "关多余窗口" + "  ".join(self.browser.window_handles)
                self.Log(strTemp)
                time.sleep(2)
                self.browser.close()
                time.sleep(2)
            self.browser.switch_to.window(self.browser.window_handles[0])
            strTemp = "当前窗口" + self.browser.current_window_handle
            self.Log(strTemp)

            self.browser.get(
                "http://eomswf.gmcc.net/default/com.huawei.nsm.platform.query.ticketQuery.flow?code=COMM_TYDB_QUERY")
            # self.browser.get("http://eomswf.gmcc.net/default/com.huawei.nsm.platform.query.ticketQuery.flow?code=COMM_PENDING_QUERY")

            time.sleep(1)
            WebDriverWait(self.browser, timeout=20).until(
                EC.presence_of_element_located((By.ID, '__queryPageFormId__qfPageSizeSelect')))
            time.sleep(1)
            '''
            js = "op=document.getElementById('__queryPageFormId__qfPageSizeSelect');op.options.add(new Option('99999',99999));"
            self.browser.execute_script(js)  # 调用/执行js语句的方法
            time.sleep(1)
            # 等待页面加载完成  __queryPageFormId__qfPageSizeSelect
            RowCnt = self.browser.find_element(By.ID, "__queryPageFormId__qfPageSizeSelect")
            Select(RowCnt).select_by_index(3)
            time.sleep(1)
            '''
            self.Log("打开EMOS工单列表完毕")

            while True:
                self.Log(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())) + "开始刷新工单：")
                ButCheck = self.browser.find_element(By.ID,"__queryPageFormId_qfForm_query")
                #点击查询按钮
                self.browser.execute_script("arguments[0].click();", ButCheck)
                time.sleep(2)
                self.browser.switch_to.frame("__qfInnerTarget")
                #time.sleep(1)
                WebDriverWait(self.browser, timeout=20).until(EC.presence_of_element_located((By.ID, 'ext-gen72')))
                #print(self.browser.page_source)
                listID = self.browser.find_elements(By.CLASS_NAME,"x-grid3-cell-inner.x-grid3-col-TICKETID_COLID")      #工单号
                #listStatus = self.browser.find_elements(By.CLASS_NAME,"x-grid3-cell-inner.x-grid3-col-STATUSNAME_COLID")#工单状态
                listName = self.browser.find_elements(By.CLASS_NAME,"x-grid3-cell-inner.x-grid3-col-TICKETTITLE_COLID") #工单名

                CycleFlag = True  # 还有工单在处理
                for i in range(len(listID)):
                    sheetId = listID[i].text
                    if sheetId.find("TSCL")>=0 and sheetId not in self.EmosSheetList: #GZCLX  TSCL listStatus[i].text.find("等待受理")>=0 and
                        strTemp = listID[i].text+"\r\n"+ listName[i].text+"\r\n"
                        self.Log(strTemp)
                        EmosText =self.Reply(listID[i])
                        self.LogEmos("\r\n工单号:"+strTemp+EmosText)
                        if EmosText != "" and CtlYDBG.sendWxMessage("厂家值班群-安全室维护组",strTemp+EmosText, 4):
                            self.Log("发送成功")
                            self.EmosSheetList.append(sheetId)
                            WatchDogFeed()
                            #保存文件
                            with open('EmosSheetList.txt', 'w') as f:
                                for item in self.EmosSheetList:
                                    f.write(item+"\n")
                            CycleFlag = False
                            break#每次刷新页面只能打开一个投诉单
                        else:
                            self.Log("发送失败")
                            return False
                if CycleFlag:
                    break #所有工单处理完毕

            self.Log("当前投诉工单处理完毕")
            return True

        except Exception as e:
            strlog = ('str(Exception):\t' + str(Exception) + '\n')
            strlog += ('str(e):\t\t' + str(e) + '\n')
            strlog += ('repr(e):\t' + repr(e) + '\n')
            # strlog += ('e.message:\t'+ e.message+'\n')
            strlog += ('traceback.print_exc():' + str(traceback.print_exc()) + '\n')
            strlog += ('traceback.format_exc():\n' + str(traceback.format_exc()) + '\n')
            self.Log(strlog)
            return False

    def Reply(self, SheetID):
        try:
            EmosID = SheetID.text #工单号
            winCnt = len(self.browser.window_handles)
            # 对定位到的元素执行鼠标双击操作
            #self.browser.execute_script("arguments[0].click();", SheetID)
            ActionChains(self.browser).move_to_element_with_offset(SheetID,0,0).double_click().perform() #
            time.sleep(3)
            for k in range(0, 10):
                #print(self.browser.window_handles)
                time.sleep(1)
                if len(self.browser.window_handles) > winCnt:#等待窗户打开
                    winCnt = len(self.browser.window_handles)
                    break
            strTemp = ("打开工单后的全部窗口"+ "  ".join(self.browser.window_handles))
            self.Log(strTemp)
            self.browser.switch_to.window(self.browser.window_handles[-1])
            strTemp = ("当前窗口"+self.browser.current_window_handle)
            self.Log(strTemp)

            #等待点击受理按钮
            time.sleep(1)
            WebDriverWait(self.browser, timeout=20).until(EC.presence_of_element_located((By.ID, 'submit')))
            # ButtonSub = self.browser.find_element(By.ID,"submit")
            # self.browser.execute_script("arguments[0].click();", ButtonSub)
            # 点击确认按钮
            # WebDriverWait(self.browser, timeout=20).until(EC.presence_of_element_located((By.ID, 'nr_button_ok')))
            # ButtonSub = self.browser.find_element(By.ID,"nr_button_ok")
            # self.browser.execute_script("arguments[0].click();", ButtonSub)
            # time.sleep(2)
            # WebDriverWait(self.browser, timeout=20).until(EC.presence_of_element_located((By.ID, 'nr_button_ok')))

            pageHtml = self.browser.page_source
            pageText = self.extract_text_from_html(pageHtml)
            # with open(EmosID+".txt", 'w') as f:
            #    f.write(pageText)
            # self.PrintAllFrame(EmosID)

            PhoneNum = self.get_content_between_strings("* 投诉号码","* 业务分类",pageText)
            CpText = self.get_content_between_strings("* 客户投诉故障情况", "* 派单意见", pageText)
            T2Data = self.get_content_between_strings("T2处理时限", "T3处理时限", pageText)

            ButtonSub = self.browser.find_element(By.XPATH, '//*[@title="日志信息"]')
            self.browser.execute_script("arguments[0].click();", ButtonSub)
            time.sleep(3)
            #self.PrintAllFrame(EmosID)
            # 获取页面上所有iframe
            # iframes = self.browser.find_elements(By.TAG_NAME, 'iframe')
            # for i, iframe in enumerate(iframes):
            #     self.browser.switch_to.frame(i)
            #     #self.PrintAllTable()
            #     CheckList = self.GetTable(3) #获取日志处理表格
            #     if CheckList != None :
            #        print(CheckList)
            #        #break
            #     self.browser.switch_to.default_content()
            self.browser.switch_to.frame(3)
            #self.PrintAllTable()
            CheckList = self.GetProcessLogTable()  # 获取日志处理表格
            if CheckList == None:
                self.browser.switch_to.default_content()
                self.browser.switch_to.frame(1)
                CheckList = self.GetProcessLogTable()  # 获取日志处理表格
            if CheckList == None :
                self.Log("Err:报错！读取工单日志失败！")
            else:
                Checklog = "\n".join(CheckList)

            EmosText = "投诉号码:"+PhoneNum+CpText+"处理日志:\n"+Checklog+"\nT2超时:"+ T2Data
            self.browser.close()
            self.browser.switch_to.window(self.browser.window_handles[0])
            strTemp = ("当前窗口" + self.browser.current_window_handle)
            self.Log(strTemp)
            time.sleep(3)
            for k in range(0, 10):
                time.sleep(1)
                if len(self.browser.window_handles) < winCnt:  # 等待窗户打开
                    break

            return EmosText
        except Exception as e:
            strlog = ('str(Exception):\t' + str(Exception) + '\n')
            strlog += ('str(e):\t\t' + str(e) + '\n')
            # strlog += ('e.message:\t'+ e.message+'\n')
            strlog += ('traceback.print_exc():' + str(traceback.print_exc()) + '\n')
            strlog += ('traceback.format_exc():\n' + str(traceback.format_exc()) + '\n')
            self.Log(strlog)
            return ""

    def extract_text_from_html(self, html):
        cleanr = re.compile('<.*?>')
        text = re.sub(cleanr, '', html)
        text = re.sub('[\r\n]+', '\n', text)  # 正则表达式匹配连续回车/换行符，然后替换成一个换行符（顺便起到统一行尾的效果）
        text = text.rstrip() #删除尾行空白
        # soup = BeautifulSoup(html, 'html.parser')
        # text = soup.find_all(text=True)
        return text

    def extract_html_content(self, html):
        pattern = "<.*?>(.*?)</.*?>"
        content = re.findall(pattern,html,re.IGNORECASE|re.DOTALL)
        return content

    def get_content_between_strings(self,str1,str2,text):
        pattern = re.escape(str1) + "(.*?)" + re.escape(str2)
        content=re.findall(pattern,text,re.IGNORECASE|re.DOTALL)
        if len(content) == 0:
            return ""
        else:
            return content[0]

    def PrintAllFrame(self,  EmosID):
        # 打印所有页面上所有iframe
        iframes = self.browser.find_elements(By.TAG_NAME, 'iframe')
        for i, iframe in enumerate(iframes):
            print(f"iframe {i}: {iframe}")
            self.browser.switch_to.frame(i)
            pageHtml = self.browser.page_source
            pageText = self.extract_text_from_html(pageHtml)
            with open(EmosID + "-F" + str(i) + ".txt", 'w') as f:
                f.write(pageText)
            self.browser.switch_to.default_content()

    def PrintAllTable(self):
        elements = self.browser.find_elements(By.TAG_NAME,'table')  # 定位所有表格
        i = 0
        for element in elements:
            lst = []  # 将表格的内容存储为list
            tr_tags = element.find_elements(By.TAG_NAME, "tr")  # 进一步定位到表格内容所在的tr节点
            for tr in tr_tags:
                td_tags = tr.find_elements(By.TAG_NAME,'td')
                for td in td_tags:
                    lst.append(td.text)  # 不断抓取的内容新增到list当中
            print(str(i))
            print(lst)
            i += 1

    #获取网页第几张表
    def GetTable(self, Index):
        try:
            elements = self.browser.find_elements(By.TAG_NAME,'table')  # 定位所有表格
            if len(elements) == 0:
                return None
            element = elements[Index]
            lst = []  # 将表格的内容存储为list
            tr_tags = element.find_elements(By.TAG_NAME, "tr")  # 进一步定位到表格内容所在的tr节点
            for tr in tr_tags:
                td_tags = tr.find_elements(By.TAG_NAME,'td')
                for td in td_tags:
                    lst.append(td.text)  # 不断抓取的内容新增到list当中
            return lst
        except Exception as e:
            print(e)
            return None

    # 获取投诉日志的处理过程，可能是第三张表
    def GetProcessLogTable(self):
        try:
            pageHtml = self.browser.page_source
            pageText = self.extract_text_from_html(pageHtml)
            res = re.findall(r"工单处理情况&nbsp;(.+?)&nbsp;", pageText, re.DOTALL)
            if len(res) == 0 :
                return  None
            return res
        except Exception as e:
            print(e)
            return None

    #对元素所在位置截图，用于元素offset矫正
    def screenshot(self,element):
        self.browser.save_screenshot('page.png')
        print(element.location)  # 打印元素坐标
        print(element.size)  # 打印元素大小
        left = element.location['x']
        top = element.location['y']
        right = element.location['x'] + element.size['width']
        bottom = element.location['y'] + element.size['height']
        im = Image.open('page.png')
        im = im.crop((left, top, right, bottom))
        im.save('ele.png')

    # 启动EMOS http://www.emip.gmcc.net/
    def EmosStart(self):
        try:
            strlog = ""
            # emos登录
            self.browser.get(r"https://www.baidu.com")#https://temip.gmcc.net/cas/login

            print("Start!")
            # 等待主页页面加载完成
            WebDriverWait(self.browser, timeout=300).until(
                EC.presence_of_element_located((By.ID, 'mxzyywId1')))
            print("Get 1")
            time.sleep(3)
            #点击卓越运维
            elem = self.browser.find_element(By.ID, "mxzyywId1")
            elem.click()
            time.sleep(3)
            print("Get 2")
            # 等待主页页面加载完成
            WebDriverWait(self.browser, timeout=300).until(
                EC.presence_of_element_located((By.ID, 'newinfoPub')))
            print("Get 3")

            self.browser.switch_to.frame("newinfoPub")
            time.sleep(1)
            self.browser.switch_to.frame("-2222221_info")
            time.sleep(1)
            self.browser.switch_to.frame("__qfInnerTarget")
            # 等待主页页面加载完成
            WebDriverWait(self.browser, timeout=300).until(
                EC.presence_of_element_located((By.ID, 'ext-gen72')))
            print("Get 4")

            #self.browser.maximize_window()
            time.sleep(3)
            # 点击刷新
            ButtonRefresh = self.browser.find_element(By.ID, "ext-gen72")
            while True:
                self.browser.execute_script("arguments[0].click();", ButtonRefresh)#ButtonRefresh.click()
                time.sleep(3)
                #获取工单url列表

            return contents

        except Exception as e:
            strlog += ('str(Exception):\t' + str(Exception) + '\n')
            strlog += ('str(e):\t\t' + str(e) + '\n')
            strlog += ('repr(e):\t' + repr(e) + '\n')
            # strlog += ('e.message:\t'+ e.message+'\n')
            strlog += ('traceback.print_exc():' + str(traceback.print_exc()) + '\n')
            strlog += ('traceback.format_exc():\n' + str(traceback.format_exc()) + '\n')
            self.Log(strlog)

    def get_content(self):
        user_agent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36"
        headers = {
            "Accept": "*/*",
            "Accept-Encoding": "gzip,deflate",
            "Accept-Language": "zh-cn",
            "Cache-Control": "no-cache",
            "Connection": "Keep-Alive",
            "Content-Length": "538",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "Host": "eomswf.gmcc.net",
            "Referer": "http://eomswf.gmcc.net/default/com.huawei.nsm.platform.query.ticketQuery.flow?_eosFlowAction=queryresult&query/type=tydbcx&query/queryIntervalCfg=-1&criteria/_entity=com.huawei.nsm.common.common.TBL_COMM_COMMONTICKET",
            "User-Agent": "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; HCTE)"
        }
        payload = {
            "_eosFlowAction": "query",
            "_eosFlowKey": "11201098-ff94-4be8-849a-20541238b6f9.view0",
            "criteria%2F_entity": "com.huawei.nsm.common.common.TBL_COMM_COMMONTICKET",
            "pageCond%2Fbegin": "0",
            "pageCond%2FisCount": "true",
            "pageCond%2Flength": "20",
            "query%2FchangeQView": "N",
            "query%2FisSetCSStyle": "Y",
            "query%2FpageCode": "COMM_TYDB_QUERY",
            "query%2FprocessNotCompletedStatus": "N",
            "query%2FqueryByOfflineView": "N",
            "query%2FqueryByParticipantView": "N",
            "query%2FqueryIntervalCfg": "-1",
            "query%2FqueryPageSize": "20",
            "query%2FsupportSetPageSize": "Y",
            "query%2Ftype": "tydbcx",
            "query%2Funique": "%2B_qf_kc16tqyk1550sv"
        }
        with open("cookies.txt", "r")as f:
            cookies = f.read()
            cookies = json.loads(cookies)
        session = requests.session()
        html = session.post("https://eomswf.gmcc.net/default/com.huawei.nsm.platform.query.ticketQuery.flow?code=COMM_TYDB_QUERY", headers=headers,data=payload,
                           cookies=cookies)
        print(html.text)


    def Post_Url(self):
        textmod = {"jsonrpc": "2.0", "method": "user.login", "params": {"user": "admin", "password": "zabbix"},
                   "auth": None, "id": 1}
        # json串数据使用
        textmod = json.dumps(textmod).encode(encoding='utf-8')
        # 普通数据使用
        #textmod = parse.urlencode(textmod).encode(encoding='utf-8')
        print(textmod)
        # 输出内容:b'{"params": {"user": "admin", "password": "zabbix"}, "auth": null, "method": "user.login", "jsonrpc": "2.0", "id": 1}'
        header_dict = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko',
                       "Content-Type": "application/json"}
        url = 'http://192.168.199.10/api_jsonrpc.php'
        req = request.Request(url=url, data=textmod, headers=header_dict,cookies=self.browser.get_cookies())
        res = request.urlopen(req)
        res = res.read()
        print(res)
        # 输出内容:b'{"jsonrpc":"2.0","result":"37d991fd583e91a0cfae6142d8d59d7e","id":1}'
        print(res.decode(encoding='utf-8'))
        # 输出内容:{"jsonrpc":"2.0","result":"37d991fd583e91a0cfae6142d8d59d7e","id":1}

    def Json(self):
        Res = {"result":[{"PROCESSINSTID":44319448,"PROCESSNAME":"Y","BIZ_WORKITEMID":220721311,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20200611-01154","CREATETIME":"2020-06-11 18:02:29","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"网络安全室","CREATEPERSONNAME":"刘峥","BIZ_LASTOPERATORNAME":"刘峥","TICKETTITLE":"5月份各地市/专业室已完成入网安全评估项目的验收情况收集(核心网工程室)","PROCESSID":100007,"BIZ_OVERTIME":"2020-06-17 18:29:30"},{"PROCESSINSTID":44278296,"PROCESSNAME":"Y","BIZ_WORKITEMID":220590039,"CREATEDEPTFULLNAME":"广东公司-网络管理中心-安全与增值业务支撑室-安全组","TICKETID":"CMCC-GD-TYGZ-20200529-00496","CREATETIME":"2020-05-29 14:37:07","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"网络安全室","CREATEPERSONNAME":"全俊斌","BIZ_LASTOPERATORNAME":"李彬","TICKETTITLE":"请协助查询疑似诈骗号码的短信话单","PROCESSID":100007,"BIZ_OVERTIME":"2020-06-04 15:35:14"},{"PROCESSINSTID":44246933,"PROCESSNAME":"Y","BIZ_WORKITEMID":220481283,"CREATEDEPTFULLNAME":"广东公司-佛山分公司-集团客户部-物联网营销中心-业务拓展室","TICKETID":"CMCC-FS-TYGZ-20200518-00385-001","CREATETIME":"2020-05-18 11:24:07","STATUSNAME":"等待受理","BIZ_PARTICIPANT":"网络安全室","CREATEPERSONNAME":"杨健林","BIZ_LASTOPERATORNAME":"杨健林","TICKETTITLE":"【业务数据核查】(网络安全室)","PROCESSID":100007,"BIZ_OVERTIME":"2020-05-18 20:29:08"},{"PROCESSINSTID":44234851,"PROCESSNAME":"Y","BIZ_WORKITEMID":220440575,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20200513-00413","CREATETIME":"2020-05-13 11:38:00","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"网络安全室","CREATEPERSONNAME":"刘峥","BIZ_LASTOPERATORNAME":"刘峥","TICKETTITLE":"4月份各地市/专业室已完成入网安全评估项目的验收情况收集(核心网工程室)","PROCESSID":100007,"BIZ_OVERTIME":"2020-05-15 18:30:02"},{"PROCESSINSTID":44132072,"PROCESSNAME":"Y","BIZ_WORKITEMID":220108633,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20200403-00592","CREATETIME":"2020-04-03 15:52:12","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"陈桂文","BIZ_LASTOPERATORNAME":"陈桂文","TICKETTITLE":"关于推荐更新2020年全省网络安全虚拟专家团队成员及收集参加众测圆活动人员的通知","PROCESSID":100007,"BIZ_OVERTIME":"2020-04-10 16:59:28"},{"PROCESSINSTID":44126705,"PROCESSNAME":"Y","BIZ_WORKITEMID":220090221,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20200401-00663","CREATETIME":"2020-04-01 17:10:51","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"王松","BIZ_LASTOPERATORNAME":"王松","TICKETTITLE":"广东公司关于召开集团“铸盾2020”攻防演习工作部署视讯会的通知","PROCESSID":100007,"BIZ_OVERTIME":"2020-04-03 18:08:07"},{"PROCESSINSTID":43565225,"PROCESSNAME":"Y","BIZ_WORKITEMID":218205001,"CREATEDEPTFULLNAME":"广东公司-网络管理中心-安全与增值业务支撑室","TICKETID":"CMCC-GD-TYGZ-20190905-01008","CREATETIME":"2019-09-05 16:46:04","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"xujiajun2","CREATEPERSONNAME":"徐家俊","BIZ_LASTOPERATORNAME":"徐家俊","TICKETTITLE":"请各地市提供综合网格对应的小区信息","PROCESSID":100007,"BIZ_OVERTIME":"2019-09-13 17:45:46"},{"PROCESSINSTID":43269644,"PROCESSNAME":"Y","BIZ_WORKITEMID":217236184,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20190617-01298-002","CREATETIME":"2019-06-17 17:14:28","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"王松","BIZ_LASTOPERATORNAME":"值班手机","TICKETTITLE":"请地市公司梳理并确认向互联网开放可疑高危端口情况（第1-1批）(综合维护组)","PROCESSID":100007,"BIZ_OVERTIME":""},{"PROCESSINSTID":43207923,"PROCESSNAME":"Y","BIZ_WORKITEMID":217023103,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20190530-01111-003","CREATETIME":"2019-05-30 16:46:17","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"王松","BIZ_LASTOPERATORNAME":"邹俊","TICKETTITLE":"关于对网管网开放OA网访问的网站（网页）认领及清理的通知(综合维护组)","PROCESSID":100007,"BIZ_OVERTIME":""},{"PROCESSINSTID":43193682,"PROCESSNAME":"Y","BIZ_WORKITEMID":216965964,"CREATEDEPTFULLNAME":"广东公司-网络管理中心-安全与增值业务支撑室","TICKETID":"CMCC-GD-TYGZ-20190527-00428-002","CREATETIME":"2019-05-27 11:21:10","STATUSNAME":"等待受理","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"杨薇琦","BIZ_LASTOPERATORNAME":"杨薇琦","TICKETTITLE":"请核对2019年HP设备维保项目第一季度工作量，谢谢。(安全与增值业务支撑室)","PROCESSID":100007,"BIZ_OVERTIME":"2019-05-29 11:19:13"},{"PROCESSINSTID":43070535,"PROCESSNAME":"Y","BIZ_WORKITEMID":216561356,"CREATEDEPTFULLNAME":"省公司-信息系统部-渠道支持室","TICKETID":"CMCC-GD-TYGZ-20190423-00511-001","CREATETIME":"2019-04-23 11:29:28","STATUSNAME":"已关闭","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"戴匠心","BIZ_LASTOPERATORNAME":"杨薇琦","TICKETTITLE":"协请网管IOD协助(网络监控室)","PROCESSID":100007,"BIZ_OVERTIME":"2019-04-26 00:00:00"},{"PROCESSINSTID":42648684,"PROCESSNAME":"Y","BIZ_WORKITEMID":215440276,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20181224-01546-010","CREATETIME":"2018-12-24 17:50:43","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"徐家俊","BIZ_LASTOPERATORNAME":"李榆鑫","TICKETTITLE":"21地市Top1000活跃小区FDD终端支持情况(网络监控室)","PROCESSID":100007,"BIZ_OVERTIME":""},{"PROCESSINSTID":40460595,"PROCESSNAME":"Y","BIZ_WORKITEMID":207264864,"CREATEDEPTFULLNAME":"广东公司-网络管理中心","TICKETID":"CMCC-GD-TYGZ-20171204-00936-001","CREATETIME":"2017-12-04 14:29:59","STATUSNAME":"确认任务","BIZ_PARTICIPANT":"安全与增值业务支撑室","CREATEPERSONNAME":"王松","BIZ_LASTOPERATORNAME":"王爱红","TICKETTITLE":"请资配室新增三个地市关口局网间来话全量话务触发虚假主叫监控平台数据配置(资源配置室)","PROCESSID":100007,"BIZ_OVERTIME":""},{"PROCESSINSTID":31905674,"PROCESSNAME":"Y","BIZ_WORKITEMID":167226238,"CREATEDEPTFULLNAME":"广东公司-汕头分公司-网络管理中心-综合维护组","TICKETID":"CMCC-GD-SCGK-20140701-01125-001","CREATETIME":"2014-07-29 17:56:14","STATUSNAME":"待确认","BIZ_PARTICIPANT":"广东公司-网络管理中心-增值网络维护室-终端组|广东公司...","CREATEPERSONNAME":"黄建安","BIZ_LASTOPERATORNAME":"王娟","TICKETTITLE":"【投诉PDCA评估】[数据]广东东莞：-测试 (网投系统) 上网类广义投诉预警-未知","PROCESSID":100026,"BIZ_OVERTIME":""},{"PROCESSINSTID":31928771,"PROCESSNAME":"Y","BIZ_WORKITEMID":167127397,"CREATEDEPTFULLNAME":"广东公司-汕头分公司-网络管理中心-综合维护组","TICKETID":"CMCC-GD-SCGK-20140701-01127-001","CREATETIME":"2014-07-30 09:18:58","STATUSNAME":"待确认","BIZ_PARTICIPANT":"广东公司-网络管理中心-增值网络维护室-终端组|广东公司...","CREATEPERSONNAME":"黄建安","BIZ_LASTOPERATORNAME":"罗志远","TICKETTITLE":"【投诉PDCA评估】[数据]广东梅州：-测试 (网投系统) 上网类广义投诉预警-未知","PROCESSID":100026,"BIZ_OVERTIME":""},{"PROCESSINSTID":31928772,"PROCESSNAME":"Y","BIZ_WORKITEMID":167123653,"CREATEDEPTFULLNAME":"广东公司-汕头分公司-网络管理中心-综合维护组","TICKETID":"CMCC-GD-SCGK-20140701-01126-001","CREATETIME":"2014-07-30 09:19:10","STATUSNAME":"待确认","BIZ_PARTICIPANT":"广东公司-网络管理中心-增值网络维护室-终端组|广东公司...","CREATEPERSONNAME":"黄建安","BIZ_LASTOPERATORNAME":"黄建安","TICKETTITLE":"【投诉PDCA评估】[数据]广东汕头：-测试 (网投系统) 上网类广义投诉预警-未知","PROCESSID":100026,"BIZ_OVERTIME":""}],"totalCount":36}
        print(Res)
        print(type(Res))

        Res1 = Res["result"]
        print(Res1)
        print(type(Res1))

        for line in Res1:
            print(line)
            print(type(line))

        res2 = json.dumps(Res, ensure_ascii=False)  # 先把字典转成json
        print(res2)
        print(type(res2))

    # 启动EMOS http://www.emip.gmcc.net/
    def EmosStart2(self):
        strlog = ""
        try:

            contents = []  # 定义list,用于存储匹配的数据,这里保存所有etf的url

            # emos登录
            self.browser.get(r"https://www.baidu.com")  # https://temip.gmcc.net/cas/login

            print("Start!")
            # 等待主页页面加载完成
            WebDriverWait(self.browser, timeout=300).until(
                EC.presence_of_element_located((By.ID, 'mxzyywId1')))
            print("Get 1")
            time.sleep(3)

            return contents

        except Exception as e:
            strlog += ('str(Exception):\t' + str(Exception) + '\n')
            strlog += ('str(e):\t\t' + str(e) + '\n')
            strlog += ('repr(e):\t' + repr(e) + '\n')
            # strlog += ('e.message:\t'+ e.message+'\n')
            strlog += ('traceback.print_exc():' + str(traceback.print_exc()) + '\n')
            strlog += ('traceback.format_exc():\n' + str(traceback.format_exc()) + '\n')
            self.Log(strlog)

    def DelayProgressBar2(self, Sec, WxReply):
        # 可以通过fill设置进度条填充符号，默认“#”
        # 可以通过suffix设置成百分比显示
        bar = Bar("Delay" + str(Sec) + "s", fill='>', max=100, suffix='%(percent)d%%')
        lastTime = 0
        dlySec = float(Sec) / 100
        lastTime2=0

        for i in bar.iter(range(100)):
            time.sleep(dlySec)
            # if (dlySec * i) - lastTime2 > 10:#处理微信信息
            #     lastTime2 = dlySec * i
            #     WxReply.WxMsgReply("厂家值班群-安全室维护组")

            # 每分钟发一次心跳
            if (dlySec * i) - lastTime > 60:
                lastTime = dlySec * i
                WatchDogFeed()

#延时进度条
def DelayProgressBar(Sec):
    # 加上进度,就是将range(N)放到ProgressBar()中
    for i in tqdm(range(Sec)):
        time.sleep(1)
#延时进度条


#看门狗喂狗
def WatchDogFeed():
    try:
        url = "http://127.0.0.1:9999/"
        params = {
            "ProcessName": "EMOS-Auto.py",
            "DlyTimeSec": "120"
        }
        r = requests.get(url, params=params,timeout=5)
        r.close()
        # print("Feed Watch Dog...")
        # print(r)
    except Exception as e:
        print("WatchDogFeed:" + 'str(Exception):\t' + str(Exception) + '\n')

if __name__ == '__main__':  # 注意main后面没有括号
    WatchDogFeed()
    # 传入is_ssl的值,1为使用ssl认证,0为禁用ssl认证
    Auto = EMOS_Auto()
    WxReply = WXreply.WX_Reply()
    # 调用start方法
    iLoginFailCnt = 0
    while True:
        if Auto.Login_get_cookies() :
            WatchDogFeed()
            iFailCnt = 0
            iLoginFailCnt = 0
            while iFailCnt < 3: #连续出错3次重新登录
                if Auto.get_emos_list():
                    WatchDogFeed()
                    iFailCnt = 0
                    Auto.DelayProgressBar2(300,WxReply)
                else:
                    iFailCnt += 1
                    print("连续失败次数："+str(iFailCnt))
            #出错超过3次 重启浏览器
            Auto.ResetBrowser()
        else:#登录失败
            iLoginFailCnt +=1
            print("连续登录失败次数：" + str(iLoginFailCnt))
            if iLoginFailCnt > 5:
                # 出错超过5次 重启浏览器
                Auto.ResetBrowser()
                iLoginFailCnt = 0

