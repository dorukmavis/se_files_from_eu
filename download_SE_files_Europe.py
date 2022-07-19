import os, time, requests, lxml, re
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

euronext_data = {'France': ['#edit-market-18--2 > div:nth-child(1) > label:nth-child(2)',
                            '#edit-market-36--2 > div:nth-child(1) > label:nth-child(2)'],
                 'Belgium': ['#edit-market-07--2 > div:nth-child(1) > label:nth-child(2)'],
                 'Ireland': ['#edit-market-08--2 > div:nth-child(1) > label:nth-child(2)',
                             '#edit-market-33--2 > div:nth-child(1) > label:nth-child(2)'],
                 'Netherlands': ['#edit-market-06--2 > div:nth-child(1) > label:nth-child(2)'],
                 'Norway': ['#edit-market-16--2 > div:nth-child(1) > label:nth-child(2)',
                            '#edit-market-48--2 > div:nth-child(1) > label:nth-child(2)',
                            '#edit-market-35--2 > div:nth-child(1) > label:nth-child(2)'],
                 'Portugal': ['#edit-market-12--2 > div:nth-child(1) > label:nth-child(2)']}

euronext_url = "https://live.euronext.com/en/products/equities/list"
euronext_counter = 1

data_rest = {'Germany': {'url': 'https://www.xetra.com/xetra-en/instruments/instruments',
                         'click': ['li.linkItem:nth-child(1) > a:nth-child(1) > span:nth-child(1)']},
             'Israel': {'url': 'https://info.tase.co.il/eng/marketdata/stocks/marketdata/Pages/MarketData.aspx',
                        'click': ['.close-button-cookies',
                                  '#tblCloseBtnucGridAllSharesExportButtonUC1 > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2)',
                                  '#divExportucGridAllSharesExportButtonUC1 > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(2) > td:nth-child(1) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(2) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(3) > a:nth-child(1)']},
             'Switzerland': {'url': 'https://www.six-group.com/en/products-services/the-swiss-stock-exchange/market-data/shares/share-explorer.html',
                             'click': ['#onetrust-accept-btn-handler',
                                       '.table-header-meta > li:nth-child(1) > a:nth-child(1)']},
             'UK': {'url': 'https://www.londonstockexchange.com/reports?tab=instruments',
                    'click': ['#ccc-notify-accept > span',
                              # '#filter-toggle > div.index-filter.second-level > div > ul > li:nth-child(2) > a',
                              '#download-single > div > div > a > span']}}
urls_shared = ('http://www.nasdaqomxnordic.com/shares/listed-companies/first-north',
               'http://www.nasdaqomxnordic.com/shares/listed-companies/first-north-premier')
#######################################################################################################################
path = os.getcwd()
options = Options()
# options.headless = True
options.add_argument('--disable_gpu')
options.add_argument('--ignore-certificate-errors')
options.add_argument('--ignore-ssl-errors')
options.add_argument('--start-maximized')
options.add_experimental_option("prefs", {'download.prompt_for_download': False,
                                          'download.default_directory': path})


def save_file(country, df):
    global path
    writer = pd.ExcelWriter(path + f'\\download Stock Exchange files\\{country}\\{country} SE.xlsx')
    df.to_excel(writer, index=False)
    writer.save()


def move_latest_file(country, i=1):
    global path
    latest = max(os.listdir(), key=os.path.getctime)
    os.rename(path + f'\\{latest}',
              path + f'\\download Stock Exchange files\\{country}\\{i} - {latest}\\')


def download_country_from_euronext(country):
    global driver, euronext_url, euronext_data, euronext_counter
    print(f'\nStarting {country}')
    driver.get(euronext_url)
    time.sleep(2)
    if euronext_counter == 1:
        driver.find_element_by_css_selector(
            '#eu-cookie-compliance-categories > div.eu-cookie-compliance-categories-buttons > button').click()
        time.sleep(2)
        euronext_counter += 1
    i = 1
    for market in euronext_data[country]:
        driver.get(euronext_url)
        time.sleep(5)
        print(f'loaded website for {country}')
        time.sleep(3)
        driver.find_element_by_css_selector('button.btn-lg:nth-child(2)').click()
        time.sleep(3)
        driver.find_element_by_css_selector(
            'div.card:nth-child(1) > div:nth-child(1) > div:nth-child(2) > button:nth-child(1)').click()
        time.sleep(3)
        driver.find_element_by_css_selector(market).click()
        time.sleep(3)
        driver.find_element_by_css_selector('#edit-awl-pd-filters-submit--2').click()
        time.sleep(3)
        driver.find_element_by_css_selector('.dt-buttons > button:nth-child(2)').click()
        time.sleep(3)
        driver.find_element_by_css_selector('input.btn:nth-child(6)').click()
        print(f'clicked the link to download the file')
        time.sleep(10)
        move_latest_file(country, i)
        print('file was moved')
        i += 1


def download_country_from_euronext_2(country):
    global driver, euronext_url, euronext_data, euronext_counter, wait
    print(f'\nStarting {country}')
    driver.get(euronext_url)
    time.sleep(2)
    if euronext_counter == 1:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                               '#eu-cookie-compliance-categories > div.eu-cookie-compliance-categories-buttons > button'))).click()
        # driver.find_element_by_css_selector(
        #     '#eu-cookie-compliance-categories > div.eu-cookie-compliance-categories-buttons > button').click()
        # time.sleep(2)
        euronext_counter += 1
    i = 1
    for market in euronext_data[country]:
        driver.get(euronext_url)
        print(f'loaded website for {country}')
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn-lg:nth-child(2)'))).click()
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                               'div.card:nth-child(1) > div:nth-child(1) > div:nth-child(2) > button:nth-child(1)'))).click()
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, market))).click()
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#edit-awl-pd-filters-submit--2'))).click()
        time.sleep(3)  # need to pause for rows to filter and display and be ready to download
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '.dt-buttons > button:nth-child(2)'))).click()
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.btn:nth-child(6)'))).click()
        print(f'clicked the link to download the file')
        time.sleep(10)
        move_latest_file(country, i)
        print('file was moved')
        i += 1


def ge_is_sw_uk(country):
    global driver, data_rest, wait
    print(f'\nStarting {country}')
    driver.get(data_rest[country]['url'])
    time.sleep(5)
    print(f'loaded website for {country}')
    if country == 'UK':
        i = 0
        while i < len(data_rest[country]['click']):
            click = data_rest[country]['click'][i]
            try:
                print('UK click - ', i)
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, click))).click()
                i += 1
            except Exception as e:
                print('error, retrying')

    else:
        for click in data_rest[country]['click']:
            # print(click)
            # driver.execute_script("window.scrollTo(0, 0);")
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, click))).click()
            # driver.execute_script("document.body.style.zoom='50%'")
            # driver.find_element_by_css_selector(click).click()
            # time.sleep(2)
    print(f'clicked the link to download the file')
    time.sleep(10)
    move_latest_file(country)
    print('file was moved')


def austria(country='Austria'):
    print(f'\nStarting {country}')
    tables = pd.read_html('https://www.wienerborse.at/en/issuers/shares/companies-list/', match='ISIN')
    print('Table for Austria is extracted')
    au_df = tables[0]
    save_file(country, au_df)


def denmark(country='Denmark'):
    global urls_shared, driver
    print(f'\nStarting {country}')
    driver.get('http://www.nasdaqomxnordic.com/shares/listed-companies/copenhagen')
    html = driver.page_source
    table = pd.read_html(html, attrs={'id': 'listedCompanies'})
    DE_df = table[0]
    DE_df = DE_df.drop(0).drop(columns=DE_df.columns[4:])
    for url in urls_shared:
        driver.get(url)
        html = driver.page_source
        table = pd.read_html(html, attrs={'id': 'listedCompanies'})
        df = table[0]
        df = df.drop(0).drop(columns=df.columns[4:])
        DE_df = pd.concat([DE_df, df], ignore_index=True)
    DE_df.drop_duplicates(ignore_index=True, inplace=True)
    save_file(country, DE_df)


def finland(country='Finland'):
    global urls_shared, driver
    print(f'\nStarting {country}')
    driver.get('http://www.nasdaqomxnordic.com/shares/listed-companies/helsinki')
    html = driver.page_source
    table = pd.read_html(html, attrs={'id': 'listedCompanies'})
    FN_df = table[0]
    FN_df = FN_df.drop(0).drop(columns=FN_df.columns[4:])
    for url in urls_shared:
        driver.get(url)
        html = driver.page_source
        table = pd.read_html(html, attrs={'id': 'listedCompanies'})
        df = table[0]
        df = df.drop(0).drop(columns=df.columns[4:])
        FN_df = pd.concat([FN_df, df], ignore_index=True)
    FN_df.drop_duplicates(ignore_index=True, inplace=True)
    save_file(country, FN_df)


def spain(country='Spain'):
    global driver, wait
    print(f'\nStarting {country}')
    spain_url = "http://www.bolsamadrid.es/ing/aspx/Empresas/Empresas.aspx"
    spain_list = list()
    driver.get(spain_url)
    print('Starting with the bolsamadrid website\n')
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#CookiesOk"))).click()
    # driver.find_element_by_css_selector("#CookiesOk").click()

    # extract table data: companies names and links
    for i in range(7):  # the table consists of 6 tabs
        driver.implicitly_wait(100)
        html = driver.page_source
        soup_object = BeautifulSoup(html, 'lxml')
        table = soup_object.select_one("#ctl00_Contenido_tblEmisoras").findAll('a')  # extract all links for companies
        for tag in table:
            company_name = tag.text
            link = 'https://www.bolsamadrid.es' + tag.get('href')
            #time.sleep(20)
            second_soup = BeautifulSoup(requests.get(link).text, 'html.parser')

            # below is a one-liner to extract ticker from the company webpage: soup -> find "Ticker" cell -> get next value
            try:
                ticker = second_soup.find("td", id="ctl00_Contenido_TickerDat").text.strip()
            except AttributeError:
                try:
                    ticker = second_soup.find("table", class_="TblPort SinTH").find("td", id="ctl00_Contenido_TickerDat")
                except AttributeError:
                    try:
                        ticker = second_soup.find("table", class_="TblPort SinTH").find_all("td")[3]
                    except AttributeError:
                        try:
                            ticker = second_soup.find('td', string=re.compile('Ticker')).fetchNextSiblings("td")[0]
                        except AttributeError:
                            ticker = None

            # below is extraction of the close price. It is either first or second row of a corresp. table
            try:
                close_table = second_soup.select_one('#ctl00_Contenido_tblPrecios').find_all('tr')[1]
            except AttributeError:
                close = None
                spain_list.append((company_name, close, '', ticker))
                print(company_name, ticker, close, sep=' | ')
                continue
            close_row = close_table.find_all('td')  # first row

            if close_row[2].text.strip() == '-':
                close_row = close_table.find_next('tr').find_all('td')  # second row
                close = close_row[2].text.strip()
            else:
                close = close_row[2].text.strip()
            print(company_name, ticker, close, sep=' | ')
            spain_list.append((company_name, close, '', ticker))
        try:
            driver.find_element_by_css_selector(
                "#ctl00_Contenido_SiguientesArr").click()  # load the next part of the table
        except:
            pass
        print(f"Scraped [{i+1}/7] pages\n")

    print('\nDone with the bolsamadrid website, going to bmegrowth\n')
    spain_url = 'https://www.bmegrowth.es/ing/Precios.aspx'
    driver.get(spain_url)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#CookiesOk"))).click()
    # driver.find_element_by_css_selector('#CookiesOk').click()

    # extract table data: companies names and links
    for i in range(4):  # the table consists of 6 tabs
        driver.implicitly_wait (100)
        html = driver.page_source
        soup_object = BeautifulSoup(html, 'html.parser')
        table = soup_object.select_one("#Contenido_Tbl").findAll('a')  # extract all links for companies
        table = [x for x in table if not x.get('href').startswith('java')]  # remove header
        for tag in table:
            company_name = tag.text
            link = 'https://www.bmegrowth.es/' + tag.get('href')
            second_soup = BeautifulSoup(requests.get(link).text, 'html.parser')
            time.sleep(1)  # otherwise the next lines fail to find the required info. Need time to load response?
            ticker = second_soup.find('h3', string='Ticker').find_next('p').text.strip()
            try:
                close = second_soup.find('small', string="Last price").find_previous('p').find_previous_sibling(
                    'p').text.strip()
            except AttributeError:
                close = second_soup.find('small', string="Close Price").find_previous('p').find_previous_sibling(
                    'p').text.strip()
            if close == '-':
                close = ''
            print(company_name, ticker, close, sep=' | ')
            spain_list.append((company_name, close, '', ticker))

        try:
            driver.find_element_by_css_selector("#Contenido_Siguiente").click()  # next page of the table
        except:
            pass  # the last click is meaningless
        print (f"Scraped [{i + 1}/4] pages\n")
    print('\nDone with the bmegrowth website, saving')
    spain_df = pd.DataFrame(data=spain_list, columns=['Company', 'Close', '', 'Ticker'])
    save_file(country, spain_df)


def sweden(country='Sweden'):
    global driver
    print(f'\nStarting {country}')
    print ('\nStarting with nasdaqomxnordic.com')
    driver.get('http://www.nasdaqomxnordic.com/shares/listed-companies/stockholm')
    html = driver.page_source

    table = pd.read_html(html, attrs={'id': 'listedCompanies'})
    SW_df = table[0]
    SW_df = SW_df.drop(0).drop(columns=SW_df.columns[4:])

    urls_shared_sw = ('http://www.nasdaqomxnordic.com/shares/listed-companies/first-north',
                      'http://www.nasdaqomxnordic.com/shares/listed-companies/first-north-premier',
                      'http://www.nasdaqomxnordic.com/shares/listed-companies/norwegian-listed-shares')

    for url in urls_shared_sw:
        driver.get(url)
        html = driver.page_source
        table = pd.read_html(html, attrs={'id': 'listedCompanies'})
        df = table[0]
        df = df.drop(0).drop(columns=df.columns[4:])
        SW_df = pd.concat([SW_df, df], ignore_index=True)

    print('\nFinishing with nasdaqomxnordic.com')
    print(SW_df)

    print ('\nStarting with spotlightstockmarket.com')
    driver.get('https://www.spotlightstockmarket.com/en/market-overview/share-prices/')
    driver.maximize_window()
    driver.implicitly_wait(10)
    html = driver.page_source
    print("html parsed")
    soup = BeautifulSoup(html, 'lxml')
    driver.minimize_window()
    # extracting all names and links at once based on alt attribute which they all have
    links = list()
    names = list()
    print("starting finding links...")
    tags = soup.find_all('a', alt=True)
    for tag in tags:
        names.append(tag.text)
        links.append(tag.get('href'))

    print('Names from three tables are extracted, now attempting for ISINs and short names on spotlightstockmarket.com')
    ISINs = list()
    symbols = list()

    list_counter = 1
    for link in links:
        try:
            list_start = time.monotonic()
            driver.get(f'https://www.spotlightstockmarket.com{link}')
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,
                                                   'body > main > div.page-stripe.header--colored.no-padding-top.no-padding-bottom > div > div > div > div.component-tabs__desktop > nav > ul > li:nth-child(2) > a'))).click()
            # driver.find_element_by_css_selector(
            #     'body > main > div.page-stripe.header--colored.no-padding-top.no-padding-bottom > div > div > div > div.component-tabs__desktop > nav > ul > li:nth-child(2) > a').click()
            html = driver.page_source
            soup = BeautifulSoup(html, 'lxml')
            ISIN = soup.find('p', text='ISIN-Code').find_next('li').text.strip()
            short_name = soup.find('p', text='Short Name').find_next('li').text.strip()
            ISINs.append(ISIN)
            symbols.append(short_name)
            print(f"Scraped [{list_counter}/{len(links)}]\n Time spent: {time.monotonic() - list_start}")
            list_counter += 1
        except Exception as ex:
            ISINs.append (None)
            symbols.append (None)
            print(f"\n{ex} for https://www.spotlightstockmarket.com{link}, skipping\n")

    print('Extracted successfully')
    print ('\nFinishing with spotlightstockmarket.com')

    # scraping from ngm.se website
    print('\nStarting with the ngm.se website')
    driver.get('https://www.ngm.se/en/market/main-regulated-equity')
    driver.find_element_by_css_selector('#hs-eu-confirmation-button').click()
    driver.switch_to.frame(0)
    driver.find_element_by_css_selector(
        '#quotesDiv > table > tbody > tr:nth-child(2) > td > div > div:nth-child(2) > div > table > tbody > tr:nth-child(2) > td > div > div:nth-child(1) > div > div > div:nth-child(1) > table > tbody > tr > td:nth-child(2) > div > button').click()
    time.sleep(2)
    txt = driver.find_element_by_tag_name('body').find_elements_by_tag_name('a')
    txt = txt[5:-4]  # ignore header and navigation
    i = 0
    links = list()
    while i < len(txt):
        try:
            link = txt[i].get_attribute('href')
            name = txt[i].text.strip()
        except StaleElementReferenceException:
            print('---------')
            print(f'fail, line {i}, retrieving table again and continuing')
            txt = driver.find_element_by_tag_name('body').find_elements_by_tag_name('a')
            txt = txt[5:-4]  # ignore header and navigation
            continue
        links.append(link)
        names.append(name)
        i += 1
    print(f'Company names are extracted, now attempting for ISINs and symbols')

    i = 0
    while i < len(links):
        ngm_time = time.monotonic()
        driver.get(links[i])
        driver.switch_to.frame(0)

        try:
            ISIN = driver.find_element_by_css_selector(
                '#detailviewDiv > table > tbody > tr:nth-child(2) > td > div > div:nth-child(1) > div > table > tbody > tr:nth-child(2) > td > table > tbody > tr > td:nth-child(1) > table > tbody > tr:nth-child(2) > td > div > table > tbody > tr:nth-child(2) > td > table > tbody > tr:nth-child(3) > td:nth-child(2) > div').text.strip()
        except:
            print(f'\nFailed to extract ISIN from {links[i]}, retrying\n')
            ISIN = None

        ISINs.append(ISIN)
        symbols.append(links[i].split('symbol=')[1].replace('%20', ' '))
        print(f"Scraped [{i}/{len(links)}] links\nTime spent: {time.monotonic() - ngm_time}\n")
        i += 1

    print('Symbols and ISINs are extracted')
    print('\nCombining and saving as Sweden SE')

    df = pd.DataFrame(data={'Name': names,
                            'Symbol': symbols,
                            'ISIN': ISINs})

    SW_df = pd.concat([SW_df, df], ignore_index=True)
    save_file(country, SW_df)
    print ('\nSweden was scraped successfully!')


def download_all():
    global download_dict
    failed = list()
    for country in list(download_dict.keys()):
        try:
            download_dict[country](country)
        except Exception as e:
            failed.append(country)
            print(e)
    if len(failed) > 0:
        print(f'Failed to download for the following countries: \n{failed}')


def download_one():
    global download_dict
    print('\nPlease select one of the following countries and type its name:')
    for i, country in enumerate(list(download_dict.keys())):
        print(i + 1, country, sep=' - ')
    user_input = str(input('I want to download the file for [type in country name] - '))
    try:
        download_dict[user_input](user_input)
    except Exception as ex:
        print(ex)
        print(f'Failed to download the file for {user_input}')


download_dict = {'Austria': austria,
                 'Denmark': denmark,
                 'Finland': finland,
                 'Germany': ge_is_sw_uk,
                 'Israel': ge_is_sw_uk,
                 'Switzerland': ge_is_sw_uk,
                 'UK': ge_is_sw_uk,
                 'France': download_country_from_euronext,
                 'Belgium': download_country_from_euronext,
                 'Ireland': download_country_from_euronext,
                 'Netherlands': download_country_from_euronext,
                 'Norway': download_country_from_euronext,
                 'Portugal': download_country_from_euronext,
                 'Spain': spain,
                 'Sweden': sweden}

print('Please choose one of the following options:')
print('1 - download files for all countries')
print('2 - download file(s) for one specific country')
print()
user_input_1 = str(input('My choice [1/2] - '))

input_dict = {'1': download_all,
              '2': download_one}

print('Starting Chrome...\n')
# driver = webdriver.Chrome(executable_path="D:/chromedriver.exe", options=options)
driver = webdriver.Chrome(executable_path=path + "\\chromedriver.exe", options=options)


driver.implicitly_wait(10)
wait = WebDriverWait(driver, 20)
# driver.maximize_window()
time.sleep(1)

input_dict[user_input_1]()
driver.quit()
del driver
print('\nQuitting')
