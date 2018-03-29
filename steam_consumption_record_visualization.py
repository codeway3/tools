import re
from bs4 import BeautifulSoup

soup = BeautifulSoup(open('./tmp/record.html'), 'html.parser')
div1 = soup.find(class_='page_content_ctn')
wallet_table_rows = div1.find_all('tr', class_='wallet_table_row')
sum = 0
dates = []
for wallet_table_row in wallet_table_rows:
    wht_date = wallet_table_row.find('td', class_='wht_date').stripped_strings
    wht_items = wallet_table_row.find('td', class_='wht_items').stripped_strings
    tmp = wallet_table_row.find('td', class_='wht_wallet_change')
    if tmp:
        tp = list(tmp.stripped_strings)
    wht_total = wallet_table_row.find('td', class_='wht_total').stripped_strings
    dates.append(list(wht_date)[0])
    wht_num = re.search(r'(.*)¥ (.*)', list(wht_total)[0]).groups()
    # print(wht_num)
    if not (tp and tp[0][0] == '+'):
        sum += float(wht_num[1])
    # for i in wht_date:
    #     for j in wht_total:
    #         print(i, re.search(r'¥ (.*)', j).groups()[0], list(wht_items))
print('{} ~ {}'.format(dates[-1], dates[0]))
print('SUM: {:.2f}'.format(sum))
print('AVG: {:.2f}'.format(sum/118))
