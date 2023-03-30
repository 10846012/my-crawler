from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl


def launchBrowser():
    PATH = Service(r"C:\Users\Yun\Desktop\me\crawler\chromedriver.exe")

    chrome = webdriver.Chrome(service=PATH)
    chrome.get("https://www.instagram.com/")
    time.sleep(2)
    username = chrome.find_element(
        By.XPATH, '//input[@name="username"]')
    username.send_keys('')
    password = chrome.find_element(
        By.XPATH, '//input[@name="password"]')
    password.send_keys('')
    login_click = chrome.find_element(
        By.XPATH, '//*[@id="loginForm"]/div/div[3]')
    login_click.click()

    time.sleep(3)
    chrome.get("https://www.instagram.com/nba/")
    time.sleep(5)
    data = []

    actions = ActionChains(chrome)
    for i in range(5):
        post = chrome.find_elements(
            By.XPATH, '//div[@class="_aabd _aa8k  _al3l"]/a')
        href = post[i].get_attribute('href')
        actions.move_to_element(post[i]).perform()
        time.sleep(3)
        likesAndComments = chrome.find_elements(
            By.XPATH, '//div[@class="_aacl _aacp _adda _aad3 _aad6 _aade"]/span')
        likes = likesAndComments[0].text
        likes = likes.encode("utf8").decode("cp950", "ignore")
        comments = likesAndComments[1].text
        comments = comments.encode("utf8").decode("cp950", "ignore")

        dic = {'href': href, 'likes': likes, 'comments': comments}
        data.append(dic)
        if (len(data) >= 5):
            break
    print(data)
    commentList = []
    for k in range(5):

        chrome.get(data[k]['href'])
        time.sleep(5)
        postTime = chrome.find_element(
            By.XPATH, '//time[@class="_aaqe"]')
        postTime = postTime.get_attribute('datetime')
        data[k]['time'] = postTime.encode("utf8").decode("cp950", "ignore")
        content = chrome.find_elements(
            By.XPATH, '//h1[@class="_aacl _aaco _aacu _aacx _aad7 _aade"]')[0].text
        data[k]['content'] = content.encode("utf8").decode("cp950", "ignore")
        conment = chrome.find_elements(
            By.XPATH, '//span[@class="_aacl _aaco _aacu _aacx _aad7 _aade"]')
        conment_len = len(conment)
        for i in range(conment_len-1):
            conment = chrome.find_elements(
                By.XPATH, '//span[@class="_aacl _aaco _aacu _aacx _aad7 _aade"]')[i].text
            commentList.append(conment.encode(
                "utf8").decode("cp950", "ignore"))
    data[k]['conment'] = commentList
    print('---------------')

    print(data)

    while (True):
        pass


def export(self, stocks):
    wb = openpyxl.Workbook()
    sheet = wb.create_sheet("內文", "日期", "按讚數", "回覆數", "回覆內容",)


switch = True

driver = launchBrowser()
