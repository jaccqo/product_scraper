import undetected_chromedriver as uc
import time
from termcolor import colored
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import random
from selenium.common.exceptions import TimeoutException,ElementClickInterceptedException,ElementNotInteractableException,StaleElementReferenceException,NoSuchElementException
import bs4
import requests
import colorama

# import openpyxl module
import openpyxl
import os
colorama.init()

class Excel:
    def __init__(self,img_url,product_title,price,part_number,description,status):
        file_name="new_product.xlsx"
        if os.path.exists(file_name):
            wb=openpyxl.load_workbook(file_name)
            sheet=wb.active

           
        
        else:
           
            wb = openpyxl.Workbook()

            sheet = wb.active
            sheet.append(("IMAGE URL","PRODUCT TITLE","PRICE","PART NUMBER","DESCRIPTION","STATUS"))


        sheet.append((img_url,product_title,price,part_number,description,status))

        wb.save(file_name)



class shopify_scraper:

    def __init__(self):
        options = uc.ChromeOptions()
        # setting profile
        # options.user_data_dir = "c:\\temp\\profile"

        # use specific (older) version
        self.driver = uc.Chrome(options=options) 

    def start(self):

        url="https://www.quadratec.com/categories/jeep-overland-camping-gear"

        print(colored(f"Getting {url}","green"))

        self.driver.get(url)

        selected_elements=WebDriverWait(self.driver, 20).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                        "li[class='col-xs-4 col-sm-3 col-md-2 text-center']")))
        
        
        print(selected_elements,end="\n")
        elems=0
        # loop through each camping item in website 
       
        while elems<len(selected_elements):
            k=elems+1
            try:
                selected_elements[elems].click()

                time.sleep(0.5)

            
                print(colored(f"collecting data  {round(k/len(selected_elements)*100,4)}%","green"))
                elems+=1

                # get each item inside the collection

                # collection len
                current_url=""
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")
                try:

                    pages_count=WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR,
                                                "ul[class='pagination']")))
                    
                except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                    pages_count=1
               
                    
            except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                selected_elements[elems].click()

                time.sleep(0.5)

            
                print(colored(f"collecting data  {round(k/len(selected_elements)*100,4)}%","green"))
                elems+=1

                # get each item inside the collection

                # collection len
                current_url=""
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight)")

                try:

                    pages_count=WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR,
                                                "ul[class='pagination']")))
                    
                except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                    pages_count=1
                    
                
            
            if pages_count==1:
                pages_len=1
            else:
                pages_count=bs4.BeautifulSoup(pages_count.get_attribute("innerHTML"),"lxml")
                pages_len=pages_count.find_all("li")
                pages_len=len(pages_len)-2

            print(colored(f" total pages found {pages_len}","red"))
            page=1
            while page<=pages_len:
                print(colored(f" pages progress {round(page/pages_len*100,4)}","red"))
                try:
                    if page==1:
                        collection_len=WebDriverWait(self.driver, 20).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                    "div[class='row row-autowrap']")))[1]
                        
                        self.driver.execute_script("arguments[0].scrollIntoView();", collection_len)


                        soup=bs4.BeautifulSoup(collection_len.get_attribute("innerHTML"),"lxml")

                        collection_articles = soup.find_all("article")

                        current_collection_len=len(collection_articles)

                        print(colored(f"This collection len {current_collection_len}","cyan"))

                    else:
                        collection_len=WebDriverWait(self.driver, 20).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                    "div[class='row row-autowrap']")))[0]
                        
                        self.driver.execute_script("arguments[0].scrollIntoView();", collection_len)


                        soup=bs4.BeautifulSoup(collection_len.get_attribute("innerHTML"),"lxml")

                        collection_articles = soup.find_all("article")

                        current_collection_len=len(collection_articles)

                        print(colored(f"This collection len {current_collection_len}","cyan"))

                except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                    
                    if page==1:
                        collection_len=WebDriverWait(self.driver, 20).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                    "div[class='row row-autowrap']")))[1]
                        
                        self.driver.execute_script("arguments[0].scrollIntoView();", collection_len)


                        soup=bs4.BeautifulSoup(collection_len.get_attribute("innerHTML"),"lxml")

                        collection_articles = soup.find_all("article")

                        current_collection_len=len(collection_articles)

                        print(colored(f"This collection len {current_collection_len}","cyan"))

                    else:
                        collection_len=WebDriverWait(self.driver, 20).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                                    "div[class='row row-autowrap']")))[0]
                        
                        self.driver.execute_script("arguments[0].scrollIntoView();", collection_len)


                        soup=bs4.BeautifulSoup(collection_len.get_attribute("innerHTML"),"lxml")

                        collection_articles = soup.find_all("article")

                        current_collection_len=len(collection_articles)

                        print(colored(f"This collection len {current_collection_len}","cyan"))
                
                item=1
            
                #get each item in columns
                while item<current_collection_len:
                    try:
                        try:
                            collection_item=WebDriverWait(self.driver,10).until(
                            EC.presence_of_element_located((By.XPATH,f"/html/body/div[4]/div/div/section/div[2]/section[3]/div/div/div/div[3]/div/div/div[1]/article[{item}]")))
                        except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                            print("retrying..")
                            if page==1:
                                collection_item=WebDriverWait(self.driver,5).until(
                                EC.presence_of_element_located((By.XPATH,f"/html/body/div[4]/div/div/section/div[2]/section[3]/div/div/div/div[2]/div/div/div[1]/article[{item}]")))
                            else:
                                collection_item=WebDriverWait(self.driver,5).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[4]/div/div/section/div[2]/section[3]/div/div/div/div/div/div/div[1]/article[{item}]")))
                 
                                                                    
                                                                   
                    
                    except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                        try:
                            collection_item=WebDriverWait(self.driver,10).until(
                            EC.presence_of_element_located((By.XPATH,f"/html/body/div[4]/div/div/section/div[2]/section[3]/div/div/div/div[3]/div/div/div[1]/article[{item}]")))
                        except (TimeoutException, ElementClickInterceptedException, ElementNotInteractableException,
                        StaleElementReferenceException, NoSuchElementException) as e:
                            print("retrying..")
                            if page==1:
                                collection_item=WebDriverWait(self.driver,5).until(
                                EC.presence_of_element_located((By.XPATH,f"/html/body/div[4]/div/div/section/div[2]/section[3]/div/div/div/div[2]/div/div/div[1]/article[{item}]")))
                            else:
                                collection_item=WebDriverWait(self.driver,5).until(EC.presence_of_element_located((By.XPATH, f"/html/body/div[4]/div/div/section/div[2]/section[3]/div/div/div/div/div/div/div[1]/article[{item}]")))
                 
                                           
                    self.driver.execute_script("arguments[0].scrollIntoView();", collection_item)

                    item_attributes=bs4.BeautifulSoup(collection_item.get_attribute("innerHTML"),"lxml")
                    

                    image_link=item_attributes.find('img')["src"]            
                    title=item_attributes.find('div',{"class":"title"})
                    price=item_attributes.find('span',{"class":"product-price"})
                    suggested_status=item_attributes.find('span',{"class":"suggested-price"})
                    stock_status=item_attributes.find('div',{"class":"serp-stock-status"})
                    part_number_url=item_attributes.find('a')['href']

                    part_number_url=f"https://www.quadratec.com{part_number_url}"

                    part_num=requests.get(part_number_url).text

                    parsed_num=bs4.BeautifulSoup(part_num,"lxml")

                    img_url=parsed_num.find("div",{"class":"slide__content"}).find("a")["href"]

                    if ".jpg" not in img_url:
                         
                        img_url=parsed_num.find("div",{"class":"slick__slide"}).find("a")["href"]


                    product_description=parsed_num.find('div',{"itemprop":"description"})
                    
                    parsed_num=parsed_num.find('span',{"class":"num-value"})

                    product_title=title.get_text()
                    product_price=price.get_text()
                    if parsed_num:
                        product_num=parsed_num.get_text()
                    else:
                        product_num="none"


                   
                    print(colored(product_title,"yellow"))
                    print(product_price)
                    if suggested_status:
                        print(suggested_status.get_text())
                    print(stock_status.get_text())
                    print(product_num)

                    print(img_url)
            
                  
                    if product_description:
                        product_desc=product_description.get_text()

                        
                    else:
                        product_desc="no description found"

                    print(colored(product_desc,"cyan"))

                    print("\n\n")
                    image_link=img_url

                    Excel(image_link,product_title,product_price,product_num,product_desc,stock_status.get_text())
                
                    time.sleep(2)
                    item+=1

                if page==1:

                    current_url=self.driver.current_url

                print(page,"||",pages_len)

                if page==pages_len:
                    print("reached end of resuslts")
                    break
                else:
                    self.driver.get(f"{current_url}?page={page}")

                    page+=1

    
            print(colored("going back to main page","cyan"))
            self.driver.get(url)

            selected_elements=WebDriverWait(self.driver, 20).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR,
                                        "li[class='col-xs-4 col-sm-3 col-md-2 text-center']")))
                
        
        

        self.driver.close()

if __name__=="__main__":
  
  
    scrape=shopify_scraper()
    scrape.start()
    
