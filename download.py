from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import uiautomation as auto
'''
# time.sleep(10)
browser = webdriver.Edge(executable_path='msedgedriver.exe')

browser.get("https://pan.baidu.com/s/1ZgxbwKQE23RBvklNfM-lTw")

browser.find_element_by_id("accessCode").send_keys('8888' + Keys.RETURN)

WebDriverWait(browser, 100).until(lambda x: x.find_element_by_css_selector(".icon-download")).click()
# time.sleep(50)
chromeWindow = auto.WindowControl(searchDepth=1, ClassName='Chrome_WidgetWin_1')
try:
    pop = chromeWindow.PaneControl(ClassName='BrowserRootView')
except:
    pop = None
if pop is not None:
    print(pop.Name)
    pop.ButtonControl(ClassName='MdTextButton').Click()

save = auto.PaneControl(ClassName='BaseGui')
print(save.Name)
save.Click(392,188)

exit()
browser.switch_to.alert
time.sleep(1)
print(browser.text)
print(lis[0].click())
# print(elem.text)
browser.quit()
exit()
'''


browser = webdriver.Edge(executable_path='msedgedriver.exe')

# browser.get('http://www.gangju5.cc/2021tvb/aimeilikuangxiangquTVBban/')
browser.get('http://www.gangju5.cc/2021tvb/zhinenairenyueyu/')

# assert '百度' in browser.title
# elem = browser.find_element_by_class_name('yunpan')  # Find the search box
# elem.send_keys('seleniumhq' + Keys.RETURN)
print(browser.window_handles)
raw_win = browser.window_handles[0]

elem = browser.find_element_by_xpath("//h2[@class='yunpan']/following-sibling::ul")
lis = elem.find_elements_by_tag_name("li")

for i in range(len(lis)-1,-1,-1):
    # if i<24:
        # continue
    lis[i].click()
    print(browser.window_handles)
    browser.switch_to.window(browser.window_handles[1])
    browser.implicitly_wait(30)
    print(browser.title)
    print(browser.current_url)
    while browser.find_elements_by_id("main-frame-error"):
        print("refresh")
        browser.refresh()
    frame = browser.find_elements_by_tag_name("iframe")
    browser.switch_to.frame(frame[1])
    browser.find_element(By.PARTIAL_LINK_TEXT,"pan.baidu.com").click()
    print(browser.window_handles)
    browser.switch_to.window(browser.window_handles[2])
    browser.implicitly_wait(30)
    print(browser.title)

    browser.find_element_by_id("accessCode").send_keys('8888' + Keys.RETURN)

    WebDriverWait(browser, 100).until(lambda x: x.find_element_by_css_selector(".icon-download")).click()
    # time.sleep(50)
    chromeWindow = auto.WindowControl(searchDepth=1, ClassName='Chrome_WidgetWin_1')
    try:
        pop = chromeWindow.PaneControl(ClassName='BrowserRootView')
    except:
        pop = None
    if pop is not None:
        print(pop.Name)
        pop.ButtonControl(ClassName='MdTextButton').Click()

    save = auto.PaneControl(ClassName='BaseGui')
    print(save.Name)
    print(save.BoundingRectangle)
    save.Click(392,188)
    for w in browser.window_handles:
        if w != raw_win:
            print(w)
            browser.switch_to.window(w)
            browser.close()
    browser.switch_to.window(raw_win)

browser.quit()