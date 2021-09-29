# %%

import geckodriver_autoinstaller
import selenium
import bs4
import requests
import os
from selenium import webdriver
import re
from datetime import datetime as dt
from time import sleep
import pandas as pd

geckodriver_autoinstaller.install()


# %%

def comfinspb_parse(thresh_date):
    # parses site of Financial Committee of SPB
    thresh_date = dt.strptime(thresh_date, '%d.%m.%Y')  # the date after that parser will stop
    last_auc_date = dt.today()  # initialization of minimum date on the page
    comfin_url = 'https://комфинспб.рф/committees/news/'
    key_message = 'об итогах депозитного'  # pattern of necessary text
    driver = webdriver.Firefox()
    driver.get(comfin_url)
    sleep(10)  # time to load page

    def process_text(text):
        # first column is string
        return text.replace('\n', ''). \
            replace('\xa0', ''). \
            replace(',', '.'). \
            replace('.00', '')

    def process_nums(nums):
        # second column is numbers
        new_nums = []
        nums = nums.split(' – ')
        for n in nums:
            n = n.replace('\n', ''). \
                replace('\xa0', ''). \
                replace(' ', ''). \
                replace(',', '.'). \
                replace('.00', '')
            try:
                n = float(n)
            except ValueError:
                pass
            new_nums.append(n)
        return new_nums

    def get_last_auc_date():
        # get last date from necessary messages on page (that is not scrolled down)
        news_list = []
        for x in driver.find_elements_by_tag_name('h3'):
            if key_message.upper() in x.text:
                news_list.append(x)
        last = news_list[-1]
        return last.text.split('-')[-1][1:]

    def make_df(all_tables):
        # makes pd
        df = pd.DataFrame(all_tables,
                          columns=['Дата отбора', 'Срок', 'Объем, млн',
                                   'Минимальная ставка', 'Масимальная ставка',
                                   'Ставка отсеченя'])
        df['Вид средств'] = "Бюджет Ст.Петербурга"
        df['Площадка'] = "СПВБ"
        df = df[['Дата отбора', 'Вид средств', 'Площадка', 'Объем, млн', 'Срок',
                 'Минимальная ставка', 'Масимальная ставка', 'Ставка отсеченя']]
        return df

    # main
    all_tables = []  # init list with data
    key_message = 'об итогах депозитного аукциона'

    # regexp patterns
    term_repattern = '[0-9]{1,3} Д'
    fact_value_pattern = '[0-9].[0-9]{3}.[0-9]{3}.[0-9]{3}'
    rate_pattern = '[0-9]{1,2},[0-9]{1,2}'

    # Scroll to thresh_date pressing button More
    while last_auc_date >= thresh_date:
        last_auc_date = dt.strptime(get_last_auc_date(), '%d.%m.%Y')
        driver.find_element_by_css_selector(
            'html body div#root div.appWrapper div.contentWrapper div.pageContentWrapper-0-59 div div.conditionalWrapper-0-99 div.fullContentField div.rightContentField div div div div div.react-swipeable-view-container div div div.backButtonContainer-0-422.backButtonContainer-0-273 button div div span').click()

    for x in driver.find_elements_by_tag_name('h3'):
        if key_message.upper() in x.text:
            auc_date = x.text.split('-')[-1][1:]
            #             print(x.text)
            terms = re.search(term_repattern, x.text).group().split()[0]
            x.click()
            html = driver.page_source
            soup = bs4.BeautifulSoup(html)

            # Case of table
            if len(soup.find_all('table', attrs={'class': 'MsoNormalTable'})) >= 1:
                #                 print('Table')
                table = soup.find_all('table', attrs={'class': 'MsoNormalTable'})[0]
                tab_data = {}

                for tr in table.find_all('tr'):
                    temp_data = []
                    for td in tr.find_all('td'):
                        temp_data.append(td.text)
                    temp_data[0] = process_text(temp_data[0])
                    temp_data[1] = process_nums(temp_data[1])
                    tab_data[temp_data[0]] = temp_data[1]
                #                 print(tab_data)

                try:
                    fact_value = \
                        tab_data.get(
                            'Фактический  объем размещения Средств бюджета на текущий Процентный период. рублей')[0]
                except TypeError:
                    fact_value = '-'
                try:
                    min_rate = \
                        tab_data.get('Диапазон  предложенных Ставок депозита. процентов годовых')[0]
                except TypeError:
                    min_rate = '-'
                try:
                    max_rate = \
                        tab_data.get('Диапазон  предложенных Ставок депозита. процентов годовых')[1]
                except TypeError:
                    max_rate = '-'
                try:
                    cutoff_rate = \
                        tab_data.get('Средневзвешенная  Ставка депозита по удовлетворенным Заявкам. процентов годовых')[
                            0]
                except TypeError:
                    cutoff_rate = '-'

            else:
                #                 print('List')
                tab_data = {'fact_value': None,
                            'cutoff_rate': None,
                            'min_rate': None,
                            'max_rate': None
                            }

                for x in soup.find_all('p', attrs={'class': 'MsoNormal'}):
                    x = x.text

                    if 'объем размещенных' in x.lower():
                        tab_data['fact_value'] = \
                            float(re.search(fact_value_pattern, x).group().replace('\xa0', ''))
                        fact_value = tab_data['fact_value']

                    elif 'отсечения' in x.lower():
                        tab_data['cutoff_rate'] = float(re.search(rate_pattern, x).group().replace('%', ''). \
                                                        replace(',', '.'))
                        cutoff_rate = tab_data['cutoff_rate']

                    elif 'минимальная' in x.lower():
                        tab_data['min_rate'] = float(re.search(rate_pattern, x).group().replace('%', ''). \
                                                     replace(',', '.'))
                        min_rate = tab_data['min_rate']

                    elif 'максимальная' in x.lower():
                        tab_data['max_rate'] = float(re.search(rate_pattern, x).group().replace('%', ''). \
                                                     replace(',', '.'))
                        max_rate = tab_data['max_rate']

                    if min_rate == max_rate:
                        cutoff_rate = max_rate

            all_tables.append([auc_date, terms, fact_value, min_rate, max_rate, cutoff_rate])
    driver.close()
    df = make_df(all_tables)
    df.to_excel('aucs.xlsx', index=None, sheet_name='comfin')
    return df


# main
if __name__ == __main__:
    df_comfinspb = comfinspb_parse('1.01.2021')