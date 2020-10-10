from docx import Document
from docx.shared import Inches
from bs4 import BeautifulSoup
import requests
import os
import selenium
from selenium import webdriver
import time
from PIL import Image
import io
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import ElementClickInterceptedException
import docformat

def main():
    opts=webdriver.ChromeOptions()
    opts.headless=True
    driver = webdriver.Chrome(ChromeDriverManager().install())
    Search_url=input("URL?")
    driver.get(Search_url)
    #Scroll to the end of the page
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(5)#sleep_between_interactions
    tmp = driver.find_element_by_xpath("//div[@class='talk-load-all-comments']")
    driver.execute_script("arguments[0].click();", tmp)
    time.sleep(10)
    iframe = driver.find_element_by_xpath("//*[@id='comment_stream_iframe']")
    soup = BeautifulSoup(driver.page_source, 'lxml')
    driver.switch_to.frame(iframe)
    time.sleep(10)
    eles = driver.find_elements_by_xpath("//div[@class='talk-load-more']")
    while len(eles)>0:
        for ele in eles:
            tmp = ele.find_element_by_tag_name('button')
            driver.execute_script("arguments[0].click();", tmp)
            time.sleep(3)
        eles = driver.find_elements_by_xpath("//div[@class='talk-load-more']")
    ch = driver.find_element_by_xpath("//div[@id='stream']")
    chats = ch.find_element_by_xpath("//div[@class='embed__stream']")
    chat = chats.get_attribute('outerHTML')
    soup2 = BeautifulSoup(chat, 'lxml')
    zhtitle = soup.find(id="block-zerohedge-page-title")
    zhcontent = soup.find(id="block-zerohedge-content")
    zhtalk = soup2.find_all('div', class_="talk-stream-comment")

    t1 = zhcontent.find_all('p')

    print(zhtitle.get_text())
    print(len(t1))
    #print(soup2.prettify())
    print(len(eles))
    for c in zhtalk:
        pass
        #print(comauth.get_text())
        #print(c['class'])
    kwargs = {'source':Search_url, 'title':zhtitle, 'body':zhcontent, 'comments':zhtalk}
    document = docformat.buildZHdoc(**kwargs)

    driver.quit()
    path = zhtitle.get_text().replace(" ","-").replace(":","").replace(".","").replace("'","").replace('"',"")
    path = path.replace("\n","")
    path = path.strip()
    document.save('C:\\Users\\etdeh\\Desktop\\{}.docx'.format(path))

    
if __name__ == "__main__":
    main()

