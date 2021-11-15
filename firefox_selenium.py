# -*- coding:utf-8 -*-
import requests
import time
from selenium import webdriver
import openpyxl
from lxml import etree

# 前序搜索条件的录入使用webdriver，其实也可以手动
'''
    优化：校外的账号密码登录没看到人机验证
    且校内cnki主页进去直接处于登录状态
    这里需要优化一下，后面可以自动下载了
'''
driver = webdriver.Firefox(executable_path='geckodriver')
#driver.execute_script()
driver.get('https://kns.cnki.net/kns8/AdvSearch?dbprefix=SCDB&&crossDbcodes=CJFQ%2CCDMD%2CCIPD%2CCCND%2CCISD%2CSNAD%2CBDZK%2CCCVD%2CCJFN%2CCCJD')  # 这个直接就是高级检索的网页
# driver.get('https://cnki.net')
time.sleep(2)

# advanced_search = driver.find_element_by_class_name('btn-grade-search')
# advanced_search.click()  # 跳转进入高级检索，后期可维护
# time.sleep(1)

major_search = driver.find_element_by_name('majorSearch')
major_search.click()    # 跳转进入专业检索，后期可维护
time.sleep(2)

# 输入检索语句，后期可维护
input_key_words = driver.find_element_by_class_name('search-middle').find_element_by_tag_name('textarea')
# key_words = input("输入检索语句：")
shuru = input("输入检索语句：")
key_words = "SU %=" + shuru 
input_key_words.send_keys(key_words)

'''
    勾一些checkbox
    Xpath选中元素右键复制就行，这里勾一下同义词扩展
'''
#check_something = driver.find_element_by_xpath('/html/body/div[2]/div/div[2]/div/div[1]/div[1]/div[2]/div[1]/div[1]/span[3]/label[1]/input')
#check_something.click()

'''
    时间范围的选择——通过修改js脚本：
    移除控件的只读属性
    然后直接把日期赋值给日历框
'''
js = "$('input:eq(1)').removeAttr('readonly')"
driver.execute_script(js)
js_value = 'document.getElementById("datebox0").value="2020-01-01"'
driver.execute_script(js_value)
time.sleep(1)

js2 = "$('input:eq(2)').removeAttr('readonly')"
driver.execute_script(js2)
date_today = time.strftime('%Y-%m-%d')
js_value2 = 'document.getElementById("datebox1").value=' + '"' + date_today + '"'
driver.execute_script(js_value2)
time.sleep(1)

search_button = driver.find_element_by_class_name('search-buttons').find_element_by_tag_name('input')
search_button.click()
time.sleep(2)


def is_chinese(uchar):
    if u'\u4e00' <= uchar <=u'\u9fff':
        return True
    else:
        return False



# 现在开始抓取每篇文献的url
results = driver.find_element_by_class_name('pagerTitleCell').find_element_by_tag_name('em').text
if len(results) > 3:
    result = str(results)
    num_results = int(result.replace(',', ''))
else:
    num_results = int(results)
print(num_results)
num_pages = int(num_results/20) + 1
url_list = []
j = 1
for j in range(1, num_pages):
    files = driver.find_elements_by_class_name('seq')
    oks = driver.find_elements_by_class_name('source')
    k = 0
    for i in files:     
        target_str = files[k].find_element_by_tag_name('input').get_attribute('value') 
        # print(target_str)
        unchar = oks[k].get_attribute('textContent')
        text = unchar.lstrip()
        rel = is_chinese(text)
        # print(rel)
        k += 1
        # 加入判别中英文的语句
        if rel == False:                # 英文url
            # print(target_str)
            en_url = 'https://schlr.cnki.net/en/Detail/index/' + target_str.split('!')[0] + '/' + target_str.split('!')[1]
            url_list.append(en_url)
        else:                           # 中文url
            # print(target_str)
            # 这里的value以“！”分割dbname和filename
            dbcode = target_str[0:4]
            split_string = target_str.split("!")
            dbname = split_string[0]
            filename = split_string[1]
            cn_url = 'https://kns.cnki.net/kcms/detail/detail.aspx?dbcode=' + dbcode + '&dbname=' + dbname + '&filename=' + filename
            # 上述索取到的url 为中文论文的，而英文论文的url并不一致
        # print(url)  
            url_list.append(cn_url)  
    turn_page = driver.find_element_by_xpath('//*[@id="PageNext"]') 
    turn_page.click()
    time.sleep(1)
# for url in url_list:
#     print(url)

# 抓取页面详情的信息，存入excel
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:90.0) Gecko/20100101 Firefox/90.0'}
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = shuru
ColumnName = ['文献名', '作者', '机构', '摘要', '来源', '关键词', 'url']
sheet.append(ColumnName)
m = 0
for address in url_list:
    
    m += 1
    print('第%d篇开始爬取' %(m))
    res = requests.get(url=address, headers=headers).content
    get_page = etree.HTML(res)
    append_list = []
    # 中英文详情页的区别，用一个用于区别就可以
    rel = address.split('/')[3]
    if rel == 'kcms':                   # 中文详情页的信息提取
        # 题目
        try:
            cn_title = get_page.xpath('/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h1/text()')[0]
            append_list.append(cn_title)
        except:
            pass
        # 作者 取第一和通讯（第二）作者
        try:
            spans = get_page.xpath('//*[@id="authorpart"]/span')
            authors = []
            for a in spans:
                author = a.xpath('./a/text()')[0]
                authors.append(author)
            append_list.append(''.join(authors))
        except:
            pass
        # 机构
        institutes = []
        try:
            institute = get_page.xpath('/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h3[2]/span/a/text()')[0]
            #　print(institute)
            institutes.append(institute)
        except:
            h3 = get_page.xpath('/html/body/div[2]/div[1]/div[3]/div/div/div[3]/div/h3[2]/a')
            for b in h3:
                institute = b.xpath('./text()')[0]
                institutes.append(institute)
        append_list.append(''.join(institutes))
        # 摘要
        try:
            abstract = get_page.xpath('//*[@id="ChDivSummary"]/text()')[0]
            append_list.append(abstract)
        except:
            pass
        # 期刊来源
        try:
            source = get_page.xpath('/html/body/div[2]/div[1]/div[3]/div/div/div[1]/div[1]/span/a[1]/text()')[0]
            append_list.append(source)
        except:
            pass
        # 关键词
        try:
            p = get_page.xpath('/html/body/div[2]/div[1]/div[3]/div/div/div[5]/p/a')
            key_words = []
            for c in p:
                keyword = c.xpath('./text()')[0]
                key_words.append(keyword)
            append_list.append(''.join(key_words))
        except:
            pass
        # 详情页链接
        append_list.append(address)
    else:
        # title
        try:
            en_title = get_page.xpath('//*[@id="doc-title"]/text()')[0]
            append_list.append(en_title)
        except:
            pass
        # 作者 取第一和通讯（第二）作者 or 遍历所有作者
        try:
            spans = get_page.xpath('//*[@id="doc-author-text"]/a')
            authors = []
            for a in spans:
                author = a.xpath('./text()')[0]
                authors.append(author)
            append_list.append(''.join(authors))
        except:
            pass
        # 机构
        try:
            h3 = get_page.xpath('//*[@id="doc-affi-text"]/span[2]/a')
            institutes = []
            for b in h3:
                institute = b.xpath('./@title')[0]
                institutes.append(institute)
            append_list.append(''.join(institutes))
        except:
            pass
        # 摘要
        try:
            abstract = get_page.xpath('//*[@id="doc-summary-content-text"]/text()')[0]
            append_list.append(abstract)
        except:
            pass
        # 期刊来源
        try:
            source = get_page.xpath('//*[@class="detail_journal_name__b1mas"]/a/text()')[0]
            append_list.append(source)
        except:
            pass
        # 关键词
        try:
            p = get_page.xpath('//*[@id="doc-keyword-text"]/a')
            key_words = []
            for c in p:
                keyword = c.xpath('./text()')[0]
                key_words.append(keyword)
            append_list.append(''.join(key_words))
        except:
            pass
        # 详情页链接
        append_list.append(address)
    # 写入excel
    sheet.append(append_list) 
    print('第%d篇爬取成功'%m)
    
wb.save(shuru + '.xlsx')
wb.close() 


driver.close()
