from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import uiautomation as auto
from win32com.client import Dispatch
import re
import sys

old_t = time.time()
def log(msg):
    global old_t
    t = time.time()
    strt = time.strftime("%H:%M:%S",time.localtime(t)) 
    print(f'{strt}({round(t-old_t)})',sys._getframe(1).f_lineno,str(msg))
    old_t = t

browser = webdriver.Edge(executable_path='msedgedriver.exe')

# browser.get('http://www.gangju5.cc/2021tvb/aimeilikuangxiangquTVBban/')
# browser.get('http://www.gangju5.cc/2021tvb/zhinenairenyueyu/')
# url = "http://www.gangju5.cc/2021tvb/nitianqianlubanyueyu/"
url = "http://www.gangju5.cc/2021tvb/xingzhenrijilubanyueyu/"
url = "http://www.gangju5.cc/2021tvb/dashudeai2021yueyu/"
url = "http://www.gangju5.cc/2021tvb/qigongzhuyueyu/"

#下载的集数，逗号隔开，可用-表示范围，为空表示全部下载，如num = "1,3-10,15"
num = "" 

log("browser.get")
browser.get(url)
log("browser.window_handles")
# assert '百度' in browser.title
# elem = browser.find_element_by_class_name('yunpan')  # Find the search box
# elem.send_keys('seleniumhq' + Keys.RETURN)
log(browser.window_handles)
raw_win = browser.window_handles[0]
log("browser.find_element_by_xpath")
elem = browser.find_element_by_xpath("//h2[@class='yunpan']/following-sibling::ul")
lis = elem.find_elements_by_tag_name("li")
log(len(lis))
if num == "":
    real_num = range(len(lis)-1,-1,-1)
else:
    tmpnum = []
    for i in num.split(","): #展开
        if i.find("-") != -1:
            m,n = i.split("-")
            if m == "":
                m = 1
            if n == "":
                n = len(lis)
            for j in range(int(m),int(n)+1):
                tmpnum.append(j)
        else:
            tmpnum.append(int(i))
    log(tmpnum)
    real_num = list(map(lambda x:len(lis)-x,tmpnum))
log(real_num)
for i in real_num:
    log(f"{i}--------------------------------------------------")
    lis[i].click()
    log(browser.window_handles)
    browser.switch_to.window(browser.window_handles[1])
    # browser.implicitly_wait(30)
    log(browser.title)
    log(browser.current_url)
    while True:
        #加载失败，刷新
        while browser.find_elements_by_id("main-frame-error"):
            log("refresh")
            browser.refresh()
        frames = browser.find_elements_by_tag_name("iframe")
        log(len(frames))
        ii = 0
        while len(frames)<2:
            ii += 1
            if ii >= 10: #10s刷新一次
                ii = 0
                log("noframe refresh")
                browser.refresh()
            time.sleep(1)
            frames = browser.find_elements_by_tag_name("iframe")
            log(len(frames))

        browser.switch_to.frame(frames[1])
        #另一个加载失败
        log("check error")
        if browser.find_elements_by_id("main-frame-error"):
            log("frame refresh")
            browser.switch_to.window(browser.window_handles[1])
            browser.refresh()
        else:
            break
    
    log("browser.find_element(By.PARTIAL_LINK_TEXT")
    try:
        link = browser.find_element(By.PARTIAL_LINK_TEXT,"pan.baidu.com")
        code = browser.find_element_by_xpath("//button[text()='复制密码']").get_attribute("data-clipboard-text")
        log(f'提取码:{code}')
        log("click  pan.baidu.com ")
        link.click()
    except Exception as e:
    # if not link:
        log(e)
        log(browser.find_element_by_class_name("plink").text)
        browser.switch_to.window(browser.window_handles[1])
        browser.close()
        browser.switch_to.window(raw_win)
        continue
    
    # browser.find_element(By.PARTIAL_LINK_TEXT,"pan.baidu.com").click()
    log(browser.window_handles)
    browser.switch_to.window(browser.window_handles[2])
    browser.implicitly_wait(30)
    log(browser.title)
    
    browser.find_element_by_id("accessCode").send_keys(code + Keys.RETURN)
    log("sent code")
    downicon = WebDriverWait(browser, 100).until(lambda x: x.find_element_by_css_selector(".icon-download"))
    log("found downicon")
    try:
        browser.find_element_by_class_name("module-share-file-main") #单文件
    except Exception as e:
        log(e)
        browser.find_element_by_class_name("share-list") #多文件
        browser.find_element_by_xpath("//span[text()='文件名']/preceding-sibling::div").click() 
    log("click download")
    downicon.click()
    time.sleep(1)

    log("find pop")
    chromeWindow = auto.WindowControl(searchDepth=1, ClassName='Chrome_WidgetWin_1')
    pop = chromeWindow.PaneControl(ClassName='BrowserRootView')
    log(f'弹框标题:{pop.Name}')
    if pop.Name == "此站点正在尝试打开 YunDetectService。":
        log("click confirm")
        pop.ButtonControl(ClassName='MdTextButton').Click()
    else: #再找一次
        time.sleep(3)
        log("find pop")
        pop = chromeWindow.PaneControl(ClassName='BrowserRootView')
        log(f'弹框标题:{pop.Name}')
        if pop.Name == "此站点正在尝试打开 YunDetectService。":
            log("click confirm")
            pop.ButtonControl(ClassName='MdTextButton').Click()
        
    found = False
    while not found:
        time.sleep(1)
        try:
            log("find save")
            save = auto.PaneControl(ClassName='BaseGui')
            log(save.Name)
            log(type(save.BoundingRectangle))
            log(save.BoundingRectangle)
            if save.Name == "设置下载存储路径":
                save.Click(392,188)
                found = True
        except Exception as e:
            log(e)
            
    # ai = Dispatch("AutoItX3.Control")
    # mobj = re.match("\((\d+),(\d+),(\d+),(\d+)\)\[.*\]",str(save.BoundingRectangle))
    # if mobj:
        # x1 = mobj.group(1)
        # y1 = mobj.group(2)
        # x2 = mobj.group(3)
        # y2 = mobj.group(4)
        # log(f'x1:{x1},y1:{y1},x2:{x2},y2:{y2}')
    # ai.MouseClick("left", int(x1)+392, int(y1)+188)
    
    for w in browser.window_handles:
        if w != raw_win:
            log(w)
            browser.switch_to.window(w)
            browser.close()
    browser.switch_to.window(raw_win)

browser.quit()