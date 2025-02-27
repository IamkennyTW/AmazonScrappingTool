#import all needed module for automation
from cmath import exp
from distutils.log import info
from ntpath import join
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
from datetime import date, datetime
import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time, openpyxl, logging, os, pandas as pd, win32com.client as win32, concurrent.futures, multiprocessing, shutil, sys

logging.basicConfig(filename='logfile.log',format='%(asctime)s (%(levelname)s) %(message)s',level=logging.INFO)

#................General....................
def evaluate_path(filename):
    #filePath = input('輸入 '+ filename +' 路徑: ')
    filePath = filename
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
        
        filePath = input(filename+' 路徑錯誤或檔案不存在或文件名不包含 %s 請重新輸入: ' % filename)
        valid_path = os.path.exists(filePath) and (filename in filePath)

    return filePath

def evaluate_digit():
    digit = input('請輸入搜尋頁數(最大5): ')

    while True:
        if digit.isdigit() and int(digit) > 0 and int(digit) < 6:
            break
        digit = input('請檢查數值並重新輸入: ')

    return int(digit)

def email_sender(receivers,filename='excel.xlsx'):
    # Define email sender and receiver
    email_sender = '13764905041@163.com'
    email_password = 'QMUUQKIAIJTYUJFX'

    # Set the subject and body of the email
    subject = 'Autosend Email for Rating/Review files'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = ", ".join(receivers)
    msg['Subject'] = subject

    body = 'Hi there, sending this email from AutoratingV5!'
    msg.attach(MIMEText(body,'plain'))
    if '\\' in filename:
        filename =  filename[filename.rfind('\\')+1:]
    attachment = open(filename,'rb')

    part = MIMEBase('application','octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition',"attachment; filename= "+filename)

    msg.attach(part)
    # Add SSL (layer of security)
    context = ssl.create_default_context()

    # Log in and send the email
    with smtplib.SMTP_SSL('smtp.163.com', 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, receivers, msg.as_string())

    logging.info('Email sent to :' + str(receivers))

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
        driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
    driver.get(url)
    html = driver.page_source
    logging.info("driver for %s is opened" % url)
    #bs4 parse data
    soup = BeautifulSoup(html, 'html.parser')

    #close Driver session    
    driver.quit()

    return soup

def url_concate(basicurl,postfix):
    wholeurl = basicurl.split('ie=UTF8')[0]+postfix
    return wholeurl

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

        logging.info("bs4 for %s is parsed" % url)

    except Exception as e:
        logging.error(e)
        raise(e)

    logging.info('Scrapping_Finish')
    return totalstarlist, reviewslist

def creat_url_list(merchandiseurl, config, numberofpages=0, basepostfix= ['ie=UTF8&reviewerType=all_reviews']*5, basestars= ['&filterByStar=one_star','&filterByStar=two_star','&filterByStar=three_star','&filterByStar=four_star','&filterByStar=five_star']):

    postfix = []
    counterlist = []
    pagenumlist = []
    #create url list
    if config == 0:
        urllist = [merchandiseurl]*5
        postfix = [m+n for m,n in zip(basepostfix,basestars)]
        starsurl = list(map(url_concate, urllist, postfix))
    
    if config == 1:
        urllist = [merchandiseurl]*3*numberofpages
        for i in range(1,numberofpages+1):
            pageurl = ['&sortBy=recent&pageNumber=%d' % i]*3
            counterlist.extend([1,2,3])
            postfix.extend([m+n+o for m,n,o in zip(basepostfix[:3],pageurl,basestars[:3])])
        starsurl = list(map(url_concate, urllist, postfix))

    if config == 2:
        urllist = [merchandiseurl]*5*numberofpages
        for i in range(1,numberofpages+1):
            pageurl = ['&sortBy=recent&pageNumber=%d' % i]*5
            counterlist.extend([1,2,3,4,5])
            pagenumlist.extend([i]*5)
            postfix.extend([m+n+o for m,n,o in zip(basepostfix,pageurl,basestars)])
        starsurl = list(map(url_concate, urllist, postfix))

    logging.info(starsurl)

    return starsurl, counterlist, pagenumlist

#...............EXCELTOOL_PANDAS....................
def add_worksheet(workbookpath, sheetsname, config):
    logging.info("Open Excel File: "+workbookpath)
    excel = win32.dynamic.Dispatch('Excel.Application')
    excel.Visible = True
    try:
        wb = excel.Workbooks.Open(workbookpath)
    except:
        # determine if application is a script file or frozen exe
        relativepath = os.path.dirname(sys.executable)+'/'+workbookpath if getattr(sys, 'frozen', False) else os.path.dirname(__file__) + '/' + workbookpath
        wb = excel.Workbooks.Open(relativepath)
    
    if config == 0:
        templatepath = os.path.dirname(sys.executable)+'/'+'Template_Rating.xlsx' if getattr(sys, 'frozen', False) else os.path.dirname(__file__) + '/' + 'Template_Rating.xlsx'
    elif config == 1:
        templatepath = os.path.dirname(sys.executable)+'/'+'Template_Review.xlsx' if getattr(sys, 'frozen', False) else os.path.dirname(__file__) + '/' + 'Template_Review.xlsx'

    wb1 = excel.Workbooks.Open(templatepath)
    templateWS = wb1.Worksheets(1)
    for sheet in sheetsname:
        logging.info("NewSheet created start... "+sheet)
        templateWS.Copy(Before=None, After=wb.Sheets(wb.Sheets.Count))
        wb.Sheets('Template').Name = sheet
        logging.info("NewSheet createded... "+sheet)

    wb.Close(True)
    wb1.Close(True)    
    excel.Application.Quit()

def save_hyperlink(workbookpath, sheetsname):
    logging.info("Saving Hyperlink...")
    excel = win32.dynamic.Dispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(workbookpath)
    except:
        # determine if application is a script file or frozen exe
        relativepath = os.path.dirname(sys.executable)+'/'+workbookpath if getattr(sys, 'frozen', False) else os.path.dirname(__file__) + '/' + workbookpath
        wb = excel.Workbooks.Open(relativepath)

    # links = [x.Hyperlink.Address for x in ws.Shapes]
    sheets_to_obj_dic = {}
    for s in sheetsname:
        ws = wb.Sheets[s]
        obj_to_url = {}
        for x in ws.Shapes:
            try:
                obj_to_url[x.ID] = x.Hyperlink.Address
            except:
                obj_to_url[x.ID] = ""

        sheets_to_obj_dic[s] = obj_to_url

    wb.Close(True)
    print(sheets_to_obj_dic)
    return sheets_to_obj_dic

def check_worksheet(workbookpath, sheetsname):
    logging.info("Workbook_checking...")
    wb = openpyxl.load_workbook(workbookpath)

    newsheet = []
    allsheetnames = wb.sheetnames

    for sheet in sheetsname:
        if sheet not in wb.sheetnames:
            newsheet.append(sheet)            
    wb.close() 
    logging.info("Workbook_check_finish")
    return newsheet, allsheetnames

def adjust_excel_width(workbookpath,sheetsname,hyperlink_dic,firstcolumnwidth=0):
    logging.info("Adjust worksheet column start!")
    excel = win32.dynamic.Dispatch('Excel.Application')
    try:
        wb = excel.Workbooks.Open(workbookpath)
    except:
        # determine if application is a script file or frozen exe
        relativepath = os.path.dirname(sys.executable)+'/'+workbookpath if getattr(sys, 'frozen', False) else os.path.dirname(__file__) + '/' + workbookpath
        wb = excel.Workbooks.Open(relativepath)
    try:
        for sheet in sheetsname:
            ws = wb.Worksheets(sheet)
            for x in ws.Shapes:
                if hyperlink_dic[sheet][x.ID] == '':
                    pass
                else:
                    x.Hyperlink.Address = hyperlink_dic[sheet][x.ID]                
            ws.Columns.AutoFit()
            if firstcolumnwidth > 0:
                ws.Columns('A').ColumnWidth = firstcolumnwidth
            logging.info(workbookpath + " WorkSheet: " + sheet + " Adjusted")

        wb.Close(True)        
    except Exception as e:
        logging.error(e)

    excel.Application.Quit()
    
def create_datalist(templist, filePath, sheetname,  config):

    logging.info('Assembly DataFrame...')
    datalist = []

    #create df list
    if config == 0:
        datalist = templist    
        # tempreviewavg = 0
        # for i, rating in enumerate(datalist):
        #     tempreviewavg += (i+1)*rating
        # totalreviews = sum(datalist)

        df1 = pd.read_excel(filePath, sheet_name=sheetname)
        columnindex = df1.shape[1]
        columnletter = openpyxl.utils.cell.get_column_letter(columnindex+1)

        sum_formula = '=SUM({}2:{}6)'.format(columnletter,columnletter)
        average_formula = '=ROUND(SUMPRODUCT({}2:{}6,$B$2:$B$6)/SUM({}2:{}6),3)'.format(columnletter,columnletter,columnletter,columnletter)
        datalist.extend([sum_formula,average_formula])
        #datalist.extend([totalreviews, 0 if totalreviews == 0 else round(tempreviewavg/totalreviews,3)])

    elif config == 1:
        for data in templist:
            datalist.extend(data)

    return datalist

def create_dataFrame(filePath,sheetindex):
    df = pd.read_excel(filePath, sheet_name=sheetindex, usecols='B:OO')
    df.columns = df.columns.astype(str)
    oldindex = [ i for i in df.columns]
    newindex = [ i.split()[0] for i in df.columns]
    res = dict(zip(oldindex, newindex))
    df.rename(columns=res,inplace=True)

    return df

def write_to_excel(datalist, filePath, sheetname, config=0):
    if config == 0:
        #df = create_dataFrame(filePath, sheetindex)
        df1 = pd.read_excel(filePath, sheet_name=sheetname)
        columnindex = df1.shape[1]
        del df1
        today = date.strftime(date.today(), "%Y/%m/%d")
        today = today.replace("/0", "/") 
        df = pd.DataFrame({today:datalist})
    elif config == 1:
        df = pd.DataFrame(datalist)

    logging.info('Start Writing excel...')
    logging.info('\t'+ df.to_string().replace('\n', '\n\t'))

    pd.io.formats.excel.ExcelFormatter.header_style = None

    if config == 0:
        with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=False, startcol=columnindex)
    elif config == 1:
        with pd.ExcelWriter(filePath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df.to_excel(writer, sheet_name=sheetname, index=False)

def main_logic(config):
    logging.info('Loading Product...')
    configdata = pd.read_excel(r"config.xlsx")
    if (configdata.shape[0] > 0) and (configdata.shape[1] == 2):
        pass
    else:
        print("[ERROR] 表格錯誤請檢查config表格格式! 必須至少有一列產品")
        logging.error(" (%d,%d) 表格錯誤請檢查config表格格式! 必須至少有一列產品" % (configdata.shape[0], configdata.shape[1]))
        os.system('pause')
        sys.exit()   

    if config == 0:
        filePath = evaluate_path('Rating.xlsx')
        newsheetlist, allsheetnams = check_worksheet(filePath, list(configdata['merchandise']))
        hyperlink_dic = save_hyperlink(filePath, allsheetnams)
        add_worksheet(filePath,newsheetlist,config)

    elif config == 1:
        filePath = evaluate_path('Review')
        numberofpages = evaluate_digit()
        newsheetlist, allsheetnams = check_worksheet(filePath, list(configdata['merchandise']))
        add_worksheet(filePath,newsheetlist,config)

    elif config == 2:
        filePath_Rating = evaluate_path('Rating')
        filePath_Review = evaluate_path('Review')
        numberofpages = evaluate_digit()
        newsheetlistRating, allsheetnams  = check_worksheet(filePath_Rating, list(configdata['merchandise']))
        newsheetlistReview, allsheetnams  = check_worksheet(filePath_Review, list(configdata['merchandise']))
        add_worksheet(filePath_Rating,newsheetlistRating,0)
        add_worksheet(filePath_Review,newsheetlistReview,1)
    
    for i in range(len(configdata.index)):
        #calling scrapping function
        sheetname = configdata['merchandise'][i]
        #start main process
        templist=[]
        if config == 0:
            #gather all urls & lists
            myurl = creat_url_list(configdata['url'][i],config)
            starsurl = myurl[0]
            del myurl
            #multiprocessing
            with concurrent.futures.ProcessPoolExecutor() as executor:
                templist = list(executor.map(get_oneTofiveStars, starsurl))
            #Get all data to excel
            mylist = create_datalist(templist, filePath, sheetname, config)
            write_to_excel(mylist, filePath, sheetname, config)

        elif config == 1:
            #gather all urls & lists
            myurl = creat_url_list(configdata['url'][i],config,numberofpages)
            starsurl = myurl[0]
            counterlist = myurl[1]
            del myurl
            #multiprocessing
            with concurrent.futures.ProcessPoolExecutor() as executor:
                templist = list(executor.map(get_comments, starsurl, counterlist))
            #Get all data to excel
            mylist = create_datalist(templist,config)
            write_to_excel(mylist, filePath, sheetname, config) 

        elif config == 2:
            #gather all urls
            myurl = creat_url_list(configdata['url'][i],config,numberofpages)
            starsurl = myurl[0]
            counterlist = myurl[1]
            pagenumlist = myurl[2]
            del myurl
            #multiprocessing
            with concurrent.futures.ProcessPoolExecutor() as executor:
                temptuple = list(executor.map(get_both, starsurl, counterlist, pagenumlist))
            
            logging.info('Assembly DataFrame...')
            #create df rating list
            ratinglist = []   
            tempreviewavg = 0
            for data in temptuple:
                ratinglist.append(data[0][0])

            for i, rating in enumerate(ratinglist):
                tempreviewavg += (i+1)*rating
            totalreviews = sum(ratinglist)
            ratinglist.extend([totalreviews,round(tempreviewavg/totalreviews,3)])
            write_to_excel(ratinglist, filePath_Rating, sheetname, 0)

            #create df review list
            reviewlist = []
            for data in temptuple:
                reviewlist.extend(data[1])    
            write_to_excel(reviewlist,filePath_Review,sheetname,1)

        logging.info('Write Complete')

    #adjust width
    if config == 0:
        adjust_excel_width(filePath,allsheetnams,hyperlink_dic,15)
        email_sender(['vicky885365@gmail.com','kennyhuang14@yahoo.com.tw'],filePath)
    elif config == 1:
        adjust_excel_width(filePath,allsheetnams)
        email_sender(['vicky885365@gmail.com','kennyhuang14@yahoo.com.tw'],filePath) 
    elif config == 2:
         adjust_excel_width(filePath_Rating,allsheetnams,15)
         adjust_excel_width(filePath_Review,allsheetnams)
         email_sender(['vicky885365@gmail.com','kennyhuang14@yahoo.com.tw'],filePath_Rating)
         email_sender(['vicky885365@gmail.com','kennyhuang14@yahoo.com.tw'],filePath_Review)

    

if __name__ == "__main__":
    multiprocessing.freeze_support()
    try:
        # str_config = input("0:Rating Only 只生成星數表格\n1:Revew Only 只生成評論表格\n2:All 生成表格\n請選擇模式: ")

        # while True:
        #     if str_config == "0" or str_config == "1"  or str_config == "2" :
        #         break
        #     str_config = input("0:Rating Only 只生成星數表格\n1:Revew Only 只生成評論表格\n2:All 生成表格\n錯誤請重新選擇模式: ")
            
        int_config = 0 #int(str_config)

        #start Timing
        start_time = time.time()
        main_logic(int_config)
        
        #os.system('pause')

    except Exception as e:
        print(e)
        logging.error(e)
        
        #os.system('pause')
        raise(e)
        
       
    logging.info("--- %s seconds ---" % (time.time() - start_time))
    os._exit(0)
