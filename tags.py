from selenium import webdriver
from time import sleep
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

data1 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','Reg_TC')
bugs1 = list(data1.BugLinks)
print('No. of Reg_TC Bugs:', len(bugs1))

data2 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','Reg_Adhoc')
bugs2 = list(data2.BugLinks)
print('No. of Reg_Adhoc Bugs:', len(bugs2))

data3 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','NF_Adhoc')
bugs3 = list(data3.BugLinks)
print('No. of NF_Adhoc Bugs:', len(bugs3))

data4 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','NF_TC')
bugs4 = list(data4.BugLinks)
print('No. of NF_TC Bugs:', len(bugs4))

Bug_FTags1 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','Bug_FTags')
tag1 = list(Bug_FTags1.Reg_TC_Tags)
print('No. of Reg_TC_Tags Tags:', len(tag1))

Bug_FTags2 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','Bug_FTags')
tag2 = list(Bug_FTags2.Reg_Adhoc_Tags)
print('No. of Reg_Adhoc_Tags Tags:', len(tag2))

Bug_FTags3 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','Bug_FTags')
tag3 = list(Bug_FTags3.NF_Adhoc_Tags)
print('No. of NF_Adhoc_Tags Tags:', len(tag3))

Bug_FTags4 = pd.read_excel('C:/Users/hdavisvi/Desktop/Workpy/Tags.xlsx','Bug_FTags')
tag4 = list(Bug_FTags4.NF_TC_Tags)
print('No. of NF_TC_Tags Tags:', len(tag4))

class freevee():
    driver = webdriver.Firefox (service= Service (GeckoDriverManager().install()))
    action = ActionChains(driver)
    url = 'https://sim.amazon.com/issues/search?q=requester%3A(hdavisvi)&sort=lastUpdatedConversationDate+desc'
    userid_locator = '//*[@id="user_name_field"]'
    username = "hdavisvi"
    signin_locator = '//*[@id="user_name_btn"]'

    def login(self):
        self.driver.get(self.url)
        self.driver.find_element(by=By.XPATH, value=self.userid_locator).send_keys(self.username)
        self.driver.find_element(by=By.XPATH, value=self.signin_locator).click()
        sleep(35)

    def add_tags(self):
        self.login()
        for i in bugs1:
            self.driver.get(i)
            self.driver.implicitly_wait(35)
            info = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[3]/div/div/ul/li[2]/a')
            sleep(2)
            self.action.move_to_element(info).click(info).perform()
            self.driver.implicitly_wait(20)
            self.driver.find_element(by=By.TAG_NAME, value="body").send_keys(Keys.END)

            for j in tag1:
                tagi = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/input')
                tagi.send_keys(j)
                add = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/button')
                add.click()
                sleep(2)
            sleep(2)
        print('Reg_TC_Tags added successfully for the listed Bugs')

        for i in bugs2:
            self.driver.get(i)
            self.driver.implicitly_wait(35)
            info = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[3]/div/div/ul/li[2]/a')
            sleep(2)
            self.action.move_to_element(info).click(info).perform()
            self.driver.implicitly_wait(20)
            self.driver.find_element(by=By.TAG_NAME, value="body").send_keys(Keys.END)

            for j in tag2:
                tagi = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/input')
                tagi.send_keys(j)
                add = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/button')
                add.click()
                sleep(2)
            sleep(2)
        print('Reg_Adhoc_Tags added successfully for the listed Bugs')

        for i in bugs3:
            self.driver.get(i)
            self.driver.implicitly_wait(35)
            info = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[3]/div/div/ul/li[2]/a')
            sleep(2)
            self.action.move_to_element(info).click(info).perform()
            self.driver.implicitly_wait(20)
            self.driver.find_element(by=By.TAG_NAME, value="body").send_keys(Keys.END)

            for j in tag3:
                tagi = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/input')
                tagi.send_keys(j)
                add = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/button')
                add.click()
                sleep(2)
            sleep(2)
        print('NF_Adhoc_Tags added successfully for the listed Bugs')

        for i in bugs4:
            self.driver.get(i)
            self.driver.implicitly_wait(35)
            info = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[3]/div/div/ul/li[2]/a')
            sleep(2)
            self.action.move_to_element(info).click(info).perform()
            self.driver.implicitly_wait(20)
            self.driver.find_element(by=By.TAG_NAME, value="body").send_keys(Keys.END)

            for j in tag4:
                tagi = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/input')
                tagi.send_keys(j)
                add = self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div[1]/div/div/div[1]/div[2]/div[4]/div[1]/div[6]/section/div[2]/form/fieldset/div[1]/span/button')
                add.click()
                sleep(2)
            sleep(2)
        print('NF_TC_Tags added successfully for the listed Bugs')
        self.driver.close()

fv = freevee()
fv.add_tags()