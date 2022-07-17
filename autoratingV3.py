#import all needed module for automation
from cmath import exp
from distutils.log import info
from ntpath import join
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from datetime import date, datetime
import time, openpyxl, logging, os, pandas as pd, win32com.client as win32, concurrent.futures, multiprocessing, shutil, sys

logging.basicConfig(filename='logfile.log',format='%(asctime)s (%(levelname)s) %(message)s',level=logging.INFO)

#................General....................
def evaluate_path(filename):
    filePath = input('輸入 '+ filename +' 路徑: ')
    valid_path = os.path.exists(filePath) and (filename in filePath)

    while True:
        if valid_path:
            filebackupname = filename + 'backup' + '.xlsx'
            if os.path.exists(r'backup'):
                shutil.copy2(filePath,r'backup/%s' % filebackupname)
            else:
                os.makedirs(r'backup')
                shutil.copy2(filePath,r'backup/%s' % filebackupname)
            logging.info('%s backup file completed' % filebackupname)
            break
        
        filePath = input(filename+' 路徑錯誤或檔案不存在或文件不包含 %s 請重新輸入: ' % filename)
        valid_path = os.path.exists(filePath) and (filename in filePath)

    return filePath

def evaluate_digit():
    digit = input('請輸入搜尋頁數(最大5): ')

    while True:
        if digit.isdigit() and int(digit) > 0 and int(digit) < 6:
            break
        digit = input('請檢查數值並重新輸入: ')

    return int(digit)

def url_concate(basicurl,postfix):
    wholeurl = basicurl.split('ie=UTF8')[0]+postfix
    return wholeurl

#................CHROME....................
def chrome_config():
    #get chrome driver path and option ready
    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')
    #options.add_argument("--headless")
    options.add_argument('--log-level=1')

    return options

def activate_session(url,options):
    #create driver session
    logging.info('Scrapping_Start.....'+url)

    try:
        driver = webdriver.Chrome(chrome_options=options)
    except:
        driver = webdriver.Chrome(ChromeDriverManager.install(),chrome_options=options)
    driver.get(url)
    html = driver.page_source
    logging.info("driver for %s is opened" % url)
    #bs4 parse data
    soup = BeautifulSoup(html, 'html.parser')

    #close Driver session    
    driver.quit()

    return soup

#...............EXCELTOOL....................
def adjust_excel_width(filepath,sheetname):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(filepath)
    except:
        # determine if application is a script file or frozen exe
        relativepath = os.path.dirname(sys.executable)+'/'+filepath if getattr(sys, 'frozen', False) else os.path.dirname(__file__) + '/' + filepath
        wb = excel.Workbooks.Open(relativepath)

    ws = wb.Worksheets(sheetname)
    ws.Columns.AutoFit()
    wb.Save()
    logging.info(filepath + " Adjusted")
    excel.Application.Quit()

def check_worksheet(workbookpath, sheetname, config):
    logging.info("workbook_check...")
    wb = openpyxl.load_workbook(workbookpath)

    if sheetname in wb.sheetnames:
        sheetindex = wb.sheetnames.index(sheetname)
    else:
        if config == 0:
            df = pd.read_excel('Template_Rating.xlsx')
        elif config == 1:
            df = pd.read_excel('Template_Review.xlsx')
        with pd.ExcelWriter(workbookpath, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer,sheet_name=sheetname,index=False)
            writer.save()

        sheetindex = len(wb.sheetnames)
        logging.info("newSheet_created: "+sheetname)

    logging.info("workbook_check_finish")
    return sheetindex

#................Scraping..................
def get_oneTofiveStars(url):
    #chrome config
    chromeoption = chrome_config()
       
    try:
        #Scaping session
        soup = activate_session(url,chromeoption)    

        review = ['0'] if soup.find('div', {'data-hook': 'cr-filter-info-review-rating-count'}) is None else soup.find('div', {'data-hook': 'cr-filter-info-review-rating-count'}).text.strip().split()
        logging.info("bs4 for %s is parsed" % url)

        #format our list 
        str_value = review[0].replace(",","")
        value = int(str_value)

    except Exception as e:
        logging.error(e)
        raise(e)

    logging.info('Scrapping_Finish')
    return value

def get_comments(url,counter):
    #chrome config
    chromeoption = chrome_config()
    
    reviewslist = []

    try:
        #Scaping session
        soup = activate_session(url,chromeoption)
        reviews = soup.find_all('div', {'data-hook': 'review'})
        for item in reviews:
            str_date = 'NoneValue' if item.find('span', {'data-hook': 'review-date'}) is None else item.find('span', {'data-hook': 'review-date'}).text.strip().split('on ')[1]
            datetime_obj = datetime.strptime(str_date, '%B %d, %Y')
            review = {
            'DATE': str(datetime_obj).split(" ")[0],
            'CUSTOMER': 'NoneValue' if item.find('span', {'class': 'a-profile-name'}) is None else item.find('span', {'class': 'a-profile-name'}).text.strip(),
            'RATING' : counter,
            'COMMENT': 'NoneValue' if item.find('span', {'data-hook': 'review-body'}) is None else item.find('span', {'data-hook': 'review-body'}).text.strip(),
            'OVERALL': ('NoneValue' if item.find('span', {'data-hook': 'review-title'}) is None else item.find('span', {'data-hook': 'review-title'}) .text.strip()) if item.find('a', {'data-hook': 'review-title'}) is None else item.find('a', {'data-hook': 'review-title'}).text.strip(),
            }
            reviewslist.append(review)

        logging.info("bs4 for %s is parsed" % url)

    except Exception as e:
        logging.error(e)
        raise(e)

    logging.info('Scrapping_Finish')
    return reviewslist

def get_both(url,counter,pages):
    #chrome config
    chromeoption = chrome_config()

    reviewslist = []
    totalstarlist = []

    try:
        #Scaping session
        soup = activate_session(url,chromeoption)

        if pages == 1:
            #format rating list
            rating = soup.find('div', {'class': 'a-row a-spacing-base a-size-base'}).text.strip().split()
            str_value = rating[0].replace(",","")
            int_value = int(str_value)
            totalstarlist.append(int_value)

        else:
            totalstarlist.append('FlagToSkip')

        if counter < 4:
            #format review list 
            reviews = soup.find_all('div', {'data-hook': 'review'})
            for item in reviews:
                str_date = 'NoneValue' if item.find('span', {'data-hook': 'review-date'}) is None else item.find('span', {'data-hook': 'review-date'}).text.strip().split('on ')[1]
                datetime_obj = datetime.strptime(str_date, '%B %d, %Y')
                review = {
                'DATE': str(datetime_obj).split(" ")[0],
                'CUSTOMER': 'NoneValue' if item.find('span', {'class': 'a-profile-name'}) is None else item.find('span', {'class': 'a-profile-name'}).text.strip(),
                'RATING' : counter,
                'COMMENT': 'NoneValue' if item.find('span', {'data-hook': 'review-body'}) is None else item.find('span', {'data-hook': 'review-body'}).text.strip(),
                'OVERALL': ('NoneValue' if item.find('span', {'data-hook': 'review-title'}) is None else item.find('span', {'data-hook': 'review-title'}) .text.strip()) if item.find('a', {'data-hook': 'review-title'}) is None else item.find('a', {'data-hook': 'review-title'}).text.strip(),
                }
                reviewslist.append(review)

        else:
            reviewslist.append('FlagToSkip')

        logging.info("bs4 for %s is parsed" % url)

    except Exception as e:
        logging.error(e)
        raise(e)

    logging.info('Scrapping_Finish')
    return totalstarlist, reviewslist

#................EXCEL.....................
def write_to_excel(config):
    logging.info('Loading Product...')
    configdata = pd.read_excel(r"config.xlsx")
    logging.info(configdata)

    if config == 0:
        filePath = evaluate_path('Rating')
    elif config == 1:
        filePath = evaluate_path('Review')
        numberofpages = evaluate_digit()
    elif config == 2:
        filePath_Rating = evaluate_path('Rating')
        filePath_Review = evaluate_path('Review')
        numberofpages = evaluate_digit()
    
    basepostfix = ['ie=UTF8&reviewerType=all_reviews']*5
    basestars = ['&filterByStar=one_star','&filterByStar=two_star','&filterByStar=three_star','&filterByStar=four_star','&filterByStar=five_star']
    
    for i in range(len(configdata.index)):
        #calling scrapping function
        sheetname = configdata['merchandise'][i]

        if config == 0:
            #create url list
            urllist = [configdata['url'][i]]*5
            postfix = [m+n for m,n in zip(basepostfix,basestars)]
            starsurl = list(map(url_concate, urllist, postfix))

            logging.info(starsurl)

            #multiprocessing
            with concurrent.futures.ProcessPoolExecutor() as executor:
                templist = list(executor.map(get_oneTofiveStars, starsurl))

            logging.info('Assembly DataFrame...')
            #create df list
            mylist = templist    
            tempreviewavg = 0
            for i, rating in enumerate(mylist):
                tempreviewavg += (i+1)*rating
            totalreviews = sum(mylist)
            mylist.extend([totalreviews, 0 if totalreviews == 0 else round(tempreviewavg/totalreviews,3)])

            #mylist = get_oneTofiveStars(configdata['url'][i])
            sheetindex = check_worksheet(filePath, sheetname, config)
            df = pd.read_excel(filePath, sheet_name=sheetindex)
            today = date.today()
            df[today] = mylist

        elif config == 1:
            #create url list
            postfix = []
            counterlist = []
            urllist = [configdata['url'][i]]*3*numberofpages
            for i in range(1,numberofpages+1):
                pageurl = ['&sortBy=recent&pageNumber=%d' % i]*3
                counterlist.extend([1,2,3])
                postfix.extend([m+n+o for m,n,o in zip(basepostfix[:3],pageurl,basestars[:3])])
            starsurl = list(map(url_concate, urllist, postfix))

            logging.info(starsurl)

            templist=[]
            #multiprocessing
            with concurrent.futures.ProcessPoolExecutor() as executor:
                templist = list(executor.map(get_comments, starsurl, counterlist))

            logging.info('Assembly DataFrame...')
            #create df list
            mylist = []
            for data in templist:
                mylist.extend(data)
            notuse_index = check_worksheet(filePath, sheetname, config)  
            df = pd.DataFrame(mylist)  

        elif config == 2:
            #create url list
            postfix = []
            counterlist = []
            pagenumlist = []
            urllist = [configdata['url'][i]]*5*numberofpages
            for i in range(1,numberofpages+1):
                pageurl = ['&sortBy=recent&pageNumber=%d' % i]*5
                counterlist.extend([1,2,3,4,5])
                pagenumlist.extend([i]*5)
                postfix.extend([m+n+o for m,n,o in zip(basepostfix,pageurl,basestars)])
            starsurl = list(map(url_concate, urllist, postfix))

            logging.info(starsurl)

            #multiprocessing
            with concurrent.futures.ProcessPoolExecutor() as executor:
                temptuple = list(executor.map(get_both, starsurl, counterlist, pagenumlist))
            
            logging.info('Assembly DataFrame...')
            #create df rating list
            ratinglist = []   
            tempreviewavg = 0
            for data in temptuple:
                if len(data[0]) > 0 and data[0][0] != 'FlagToSkip':
                    ratinglist.append(data[0][0])

            for i, rating in enumerate(ratinglist):
                tempreviewavg += (i+1)*rating
            totalreviews = sum(ratinglist)
            ratinglist.extend([totalreviews,round(tempreviewavg/totalreviews,3)])
            print(ratinglist) 
            sheetindex_Rating = check_worksheet(filePath_Rating, sheetname, 0)
            df1 = pd.read_excel(filePath_Rating, sheet_name=sheetindex_Rating)
            today = date.today()
            df1[str(today)] = ratinglist

            #create df review list
            reviewlist = []
            for data in temptuple:
                if data[1][0] != 'FlagToSkip':
                    reviewlist.extend(data[1])
            print(reviewlist)    
            notuse_index = check_worksheet(filePath_Review, sheetname, 1)
            df2 = pd.DataFrame(reviewlist)

        logging.info('Start Writing excel...')

        if config < 2:
            logging.info('\t'+ df.to_string().replace('\n', '\n\t'))
            with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheetname, index=False)
            #adjust width
            adjust_excel_width(filePath,sheetname)

        elif config == 2:
            #Write first rating file
            logging.info('\t'+ df1.to_string().replace('\n', '\n\t'))
            with pd.ExcelWriter(filePath_Rating, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df1.to_excel(writer, sheet_name=sheetname, index=False)
            #adjust width
            adjust_excel_width(filePath_Rating,sheetname)

            #Write second review file
            logging.info('\t'+ df2.to_string().replace('\n', '\n\t'))
            with pd.ExcelWriter(filePath_Review, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df2.to_excel(writer, sheet_name=sheetname, index=False)
            #adjust width
            adjust_excel_width(filePath_Review,sheetname)

        logging.info('Write Complete')

if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        str_config = input("0:Rating Only 只生成星數表格\n1:Revew Only 只生成評論表格\n2:All 生成表格\n請選擇模式: ")

        while True:
            if str_config == "0" or str_config == "1"  or str_config == "2" :
                break
            str_config = input("0:Rating Only 只生成星數表格\n1:Revew Only 只生成評論表格\n2:All 生成表格\n錯誤請重新選擇模式: ")
            
        int_config = int(str_config)

        #start Timing
        start_time = time.time()
        write_to_excel(int_config)

        os.system('pause')

    except Exception as e:
        logging.error(e)
        raise(e)        
    logging.info("--- %s seconds ---" % (time.time() - start_time))

