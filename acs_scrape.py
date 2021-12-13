# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""


import openpyxl
import bs4
import time
import pandas
import re
import random

from selenium import webdriver
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException

browser = webdriver.Chrome('C:/Users/Liam McDonald/chromedriver_win32/chromedriver.exe')
random.seed()

decades = range(192, 203)
years = range(0,10)
url_base = 'https://pubs.acs.org' 

issue_df = pandas.DataFrame(columns = ['Date', 'Number', 'Link'])

for j in range(0, len(decades)):
    
    decade = str(decades[j])
    
    for k in range(0, len(years)):
        if((decade == '192' and k < 4) or (decade == '202' and k > 1)):
            continue
        year = str(years[k])
        
        search_term = '/loi/jceda8/group/d' + decade + '0' + '.y' + decade + year

        browser.get(url_base + search_term)
        
        time.sleep(1)
        
        page = browser.page_source
        soup = bs4.BeautifulSoup(page)

        elems = soup.select('div[class="loi__issue col-lg-9 col-md-8 col-sm-9 col-xs-12"]')

        for i in range(0, len(elems)):
            links = elems[i].select('a')
    
            issue_date = links[0].getText()
            issue_num = links[1].getText().replace('I', ' I')
            issue_link = url_base + links[0]['href']
    
            issue_data = pandas.DataFrame(data = [[issue_date, issue_num, issue_link]], columns = ['Date', 'Number', 'Link'])
            issue_df = issue_df.append(issue_data, ignore_index = True)
            
wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = 'Issues'

sheet['A1'] = 'Date'
sheet['B1'] = 'Number'
sheet['C1'] = 'Link'

for i in range(0, len(issue_df)):
    sheet['A' + str(i+2)] = issue_df.at[i, 'Date']
    sheet['B' + str(i+2)] = issue_df.at[i, 'Number']
    sheet['C' + str(i+2)] = issue_df.at[i, 'Link']
    


articles_df = pandas.DataFrame(columns = ['Link', 'Title', 'Authors', 'Volume', 'Issue', 'Type', 'Date'])
for i in range(len(issue_df) - 1, -1, -1):
    url = issue_df.at[i, 'Link']
    
    browser.get(url)
    time.sleep(1)
    
    page = browser.page_source
    soup = bs4.BeautifulSoup(page)
    
    elems = soup.select('div[class="issue-item clearfix"]')
    
    for j in  range(0, len(elems)):
        links = elems[j].select('h5[class="issue-item_title"] > a')
        
        
        article_link = url_base + links[0]['href']
        article_title = links[0]['title']
        
        if article_title == 'Issue Editorial Masthead' or article_title == 'Issue Publication Information':
            continue
        
        try:
            article_authors = elems[j].select('ul[class="issue-item_loa"]')[0].getText()
        except IndexError:
            article_authors = ''
        
        article_vol_num = elems[j].select('span[class="issue-item_vol-num"]')[0].getText()
        article_issue_num = elems[j].select('span[class="issue-item_issue-num"]')[0].getText()
        
        type_form = re.compile('\(([\s\w\W]+)\)')
        article_type = type_form.findall(elems[j].select('span[class="issue-item_type"]')[1].getText())[0]
        
        article_pub_date = elems[j].select('span[class="pub-date-value"]')[0].getText()
        
        article = pandas.DataFrame(data = [[article_link, article_title, article_authors, article_vol_num, article_issue_num,
                                            article_type, article_pub_date]],
                                            columns = ['Link', 'Title', 'Authors', 'Volume', 'Issue', 'Type', 'Date'])

        articles_df = articles_df.append(article, ignore_index = True)




index_status = articles_df['Citations'].isnull()
for i in range(0, len(articles_df)):
    if index_status[i] == False:
       continue
    url = articles_df.at[i, 'Link']
    
    browser.get(url)
    
    # browser.get('https://pubs.acs.org/doi/10.1021/ed006p9')
    time.sleep(3)
    
    try:
        browser.find_element_by_id('gdpr-con-btn').click()
        time.sleep(.01)
    except ElementNotInteractableException:
        time.sleep(.01)
    
    browser.execute_script("window.scrollTo(0, 200)")
    
    try:
        click_elements = browser.find_elements_by_class_name('read-more')
        time.sleep(.25)
        for j in range(0, len(click_elements)):
            click_elements[j].click()
            
    except ElementClickInterceptedException:
        click_elements = browser.find_elements_by_class_name('read-more')
        time.sleep(.25)
        for j in range(0, len(click_elements)):
            click_elements[j].click()
    
    page = browser.page_source
    soup = bs4.BeautifulSoup(page)
    
    try:
        article_subjects = soup.select('ul[class="rlist--inline loa"]')[0].getText()
    except IndexError:
        article_subjects = ''
    
    
    keywords = soup.select('a[class="keyword"]')
    article_keywords = ''
    for j in range(0, len(keywords)):
        article_keywords = article_keywords + keywords[j].getText() 
        if j != len(keywords) - 1:
            article_keywords = article_keywords + ', '
            
    article_views = soup.select('div[class="articleMetrics-val"]')[0].getText().replace('-', '0')
    article_citations = soup.select('div[class="articleMetrics-val"]')[1].getText().replace('-', '0')
        
    articles_df.at[i, 'Subjects'] = article_subjects
    articles_df.at[i, 'Keywords'] = article_keywords
    articles_df.at[i, 'Views'] = article_views
    articles_df.at[i, 'Citations'] = article_citations
    time.sleep(random.choice((1, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4, 1.45, 1.5, 1.55, 1.6, 1.65, 1.7, 1.75, 1.8, 1.85, 1.9, 1.95, 2)))
    


wb.create_sheet(index = 0, title = 'Articles')
sheet = wb['Articles']

sheet['A1'] = 'Link'
sheet['B1'] = 'Title'
sheet['C1'] = 'Authors'
sheet['D1'] = 'Volume'
sheet['E1'] = 'Issue'
sheet['F1'] = 'Type'
sheet['G1'] = 'Date'
sheet['H1'] = 'Subjects'
sheet['I1'] = 'Keywords'
sheet['J1'] = 'Views'
sheet['K1'] = 'Citations'


for i in range(0, len(articles_df)):
    sheet['A' + str(i+2)] = articles_df.at[i, 'Link']
    sheet['B' + str(i+2)] = articles_df.at[i, 'Title']
    sheet['C' + str(i+2)] = articles_df.at[i, 'Authors']
    sheet['D' + str(i+2)] = articles_df.at[i, 'Volume']
    sheet['E' + str(i+2)] = articles_df.at[i, 'Issue']
    sheet['F' + str(i+2)] = articles_df.at[i, 'Type']
    sheet['G' + str(i+2)] = articles_df.at[i, 'Date']
    sheet['H' + str(i+2)] = articles_df.at[i, 'Subjects']
    sheet['I' + str(i+2)] = articles_df.at[i, 'Keywords']
    sheet['J' + str(i+2)] = articles_df.at[i, 'Views']
    sheet['K' + str(i+2)] = articles_df.at[i, 'Citations']    

wb.save('acs_articles_final.xlsx')  

