import json
import logging
import sqlite3
from random import randint
from time import time, sleep

from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager as CM
import requests
import openpyxl
import _thread

import re

f = open('infos/config.json', )
config = json.load(f)

DEFAULT_IMPLICIT_WAIT = 1

userInfo = {}


class InstaDM(object):

    def __init__(self, username, password, headless=True, instapy_workspace=None, profileDir=None):
        self.selectors = {
            "search_bar": "//input[@id='search']",
            "search_button": "//button[@id='search-icon-legacy']",
            "filter_menu": "//div[@id='filter-menu']//div[@id='container']//a",
            "year": "//div[@title='Search for This year'] | //div[@title='搜索“今年”']",
            "month": "//div[@title='Search for This month'] | //div[@title='搜索“本月”']",
            "day": "//div[@title='Search for This day'] | //div[@title='搜索“今天”']",
            "week": "//div[@title='Search for This week'] | //div[@title='搜索“本周”']",
            "channel_link": "//a[@id='channel-thumbnail']",
            "sub_count": "//yt-formatted-string[@id='subscriber-count']",
            "desc": "//yt-formatted-string[@id='description']",
            "location": "//*[@id='details-container']/table/tbody/tr[2]/td[2]/yt-formatted-string",
            "title": "//*[@id='text']",
            "noMore": "//*[text()='No more results'] | //*[text()='无更多结果']",
            "extLink": "//*[@id='link-list-container']/a",
            "accept_cookies": "//button[text()='Accept']",
            "home_to_login_button": "//button[text()='Log In']",
            "username_field": "username",
            "password_field": "password",
            "button_login": "//button/*[text()='Log In']",
            "login_check": "//*[@aria-label='Home'] | //button[text()='Save Info'] | //button[text()='Not Now']",
            "search_user": "queryBox",
            "select_user": '//div[text()="{}"]',
            "name": "((//div[@aria-labelledby]/div/span//img[@data-testid='user-avatar'])[1]//..//..//..//div[2]/div[2]/div)[1]",
            "next_button": "//button/*[text()='Next']",
            "textarea": "//textarea[@placeholder]",
            "send": "//button[text()='Send']",
            "button_not_now": "//button[text()='Not Now']"
        }

        self.excelData = []
        self.excelData.append(["title", "fans num", "link", "country", "desc", "extra_link"])
        self.is_no_more_result = False
        self.fetchInterval = config["fetch_interval"]
        self.minFans = config["min_fans_num"]

        # Selenium config
        options = webdriver.ChromeOptions()

        if profileDir:
            options.add_argument("user-data-dir=profiles/" + profileDir)

        if headless:
            options.add_argument("--headless")

        options.add_argument("--log-level=3")

        self.driver = webdriver.Chrome(
            executable_path=CM().install(), options=options)
        self.driver.set_window_position(0, 0)
        self.driver.maximize_window()

        _thread.start_new_thread(self.__waiting_for_stop_bot, ())

        self.mainLoop()

    def login(self, username, password):
        # homepage
        self.driver.get('https://instagram.com/?hl=en')
        self.__random_sleep__(3, 5)
        if self.__wait_for_element__(self.selectors['accept_cookies'], 'xpath', 10):
            self.__get_element__(
                self.selectors['accept_cookies'], 'xpath').click()
            self.__random_sleep__(3, 5)
        if self.__wait_for_element__(self.selectors['home_to_login_button'], 'xpath', 10):
            self.__get_element__(
                self.selectors['home_to_login_button'], 'xpath').click()
            self.__random_sleep__(5, 7)

        # login
        logging.info(f'Login with {username}')
        self.__scrolldown__()
        if not self.__wait_for_element__(self.selectors['username_field'], 'name', 10):
            print('Login Failed: username field not visible')
        else:
            self.driver.find_element_by_name(
                self.selectors['username_field']).send_keys(username)
            self.driver.find_element_by_name(
                self.selectors['password_field']).send_keys(password)
            self.__get_element__(
                self.selectors['button_login'], 'xpath').click()
            self.__random_sleep__()
            if self.__wait_for_element__(self.selectors['login_check'], 'xpath', 10):
                print('Login Successful')
            else:
                print('Login Failed: Incorrect credentials')

    def createCustomGreeting(self, greeting):
        # Get username and add custom greeting
        if self.__wait_for_element__(self.selectors['name'], "xpath", 10):
            user_name = self.__get_element__(
                self.selectors['name'], "xpath").text
            if user_name:
                greeting = greeting + " " + user_name + ", \n\n"
        else:
            greeting = greeting + ", \n\n"
        return greeting

    def typeMessage(self, user, message):
        # Go to page and type message
        if self.__wait_for_element__(self.selectors['next_button'], "xpath"):
            self.__get_element__(
                self.selectors['next_button'], "xpath").click()
            self.__random_sleep__()

        for msg in message:
            if self.__wait_for_element__(self.selectors['textarea'], "xpath"):
                self.__type_slow__(self.selectors['textarea'], "xpath", msg)
                self.__random_sleep__()

            if self.__wait_for_element__(self.selectors['send'], "xpath"):
                self.__get_element__(self.selectors['send'], "xpath").click()
                self.__random_sleep__(3, 5)
                print('Message sent successfully')

    def sendMessage(self, user, message, greeting=None):
        logging.info(f'Send message to {user}')
        print(f'Send message to {user}')
        self.driver.get('https://www.instagram.com/direct/new/?hl=en')
        self.__random_sleep__(2, 4)

        try:
            self.__wait_for_element__(self.selectors['search_user'], "name")
            self.__type_slow__(self.selectors['search_user'], "name", user)
            self.__random_sleep__(1, 2)

            if greeting != None:
                greeting = self.createCustomGreeting(greeting)

            # Select user from list
            elements = self.driver.find_elements_by_xpath(
                self.selectors['select_user'].format(user))
            if elements and len(elements) > 0:
                elements[0].click()
                self.__random_sleep__()

                if greeting != None:
                    self.typeMessage(user, message)
                else:
                    self.typeMessage(user, message)

                if self.conn is not None:
                    self.cursor.execute(
                        'INSERT INTO message (username, message) VALUES(?, ?)', (user, message[0]))
                    self.conn.commit()
                self.__random_sleep__(5, 10)

                return True

            # In case user has changed his username or has a private account
            else:
                print(f'User {user} not found! Skipping.')
                return False

        except Exception as e:
            logging.error(e)
            return False

    def sendGroupMessage(self, users, message):
        logging.info(f'Send group message to {users}')
        print(f'Send group message to {users}')
        self.driver.get('https://www.instagram.com/direct/new/?hl=en')
        self.__random_sleep__(5, 7)

        try:
            usersAndMessages = []
            for user in users:
                if self.conn is not None:
                    usersAndMessages.append((user, message))

                self.__wait_for_element__(
                    self.selectors['search_user'], "name")
                self.__type_slow__(self.selectors['search_user'], "name", user)
                self.__random_sleep__()

                # Select user from list
                elements = self.driver.find_elements_by_xpath(
                    self.selectors['select_user'].format(user))
                if elements and len(elements) > 0:
                    elements[0].click()
                    self.__random_sleep__()
                else:
                    print(f'User {user} not found! Skipping.')

            self.typeMessage(user, message)

            if self.conn is not None:
                self.cursor.executemany("""
                    INSERT OR IGNORE INTO message (username, message) VALUES(?, ?)
                """, usersAndMessages)
                self.conn.commit()
            self.__random_sleep__(50, 60)

            return True

        except Exception as e:
            logging.error(e)
            return False

    def sendGroupIDMessage(self, chatID, message):
        logging.info(f'Send group message to {chatID}')
        print(f'Send group message to {chatID}')
        self.driver.get('https://www.instagram.com/direct/inbox/')
        self.__random_sleep__(5, 7)

        # Definitely a better way to do this:
        actions = ActionChains(self.driver)
        actions.send_keys(Keys.TAB * 2 + Keys.ENTER).perform()
        actions.send_keys(Keys.TAB * 4 + Keys.ENTER).perform()

        if self.__wait_for_element__(f"//a[@href='/direct/t/{chatID}']", 'xpath', 10):
            self.__get_element__(
                f"//a[@href='/direct/t/{chatID}']", 'xpath').click()
            self.__random_sleep__(3, 5)

        try:
            usersAndMessages = [chatID]

            if self.__wait_for_element__(self.selectors['textarea'], "xpath"):
                self.__type_slow__(
                    self.selectors['textarea'], "xpath", message)
                self.__random_sleep__()

            if self.__wait_for_element__(self.selectors['send'], "xpath"):
                self.__get_element__(self.selectors['send'], "xpath").click()
                self.__random_sleep__(3, 5)
                print('Message sent successfully')

            if self.conn is not None:
                self.cursor.executemany("""
                    INSERT OR IGNORE INTO message (username, message) VALUES(?, ?)
                """, usersAndMessages)
                self.conn.commit()
            self.__random_sleep__(50, 60)

            return True

        except Exception as e:
            logging.error(e)
            return False

    def __get_element__(self, element_tag, locator):
        """Wait for element and then return when it is available"""
        try:
            locator = locator.upper()
            dr = self.driver
            if locator == 'ID' and self.is_element_present(By.ID, element_tag):
                return WebDriverWait(dr, 15).until(lambda d: dr.find_element_by_id(element_tag))
            elif locator == 'NAME' and self.is_element_present(By.NAME, element_tag):
                return WebDriverWait(dr, 15).until(lambda d: dr.find_element_by_name(element_tag))
            elif locator == 'XPATH' and self.is_element_present(By.XPATH, element_tag):
                return WebDriverWait(dr, 15).until(lambda d: dr.find_element_by_xpath(element_tag))
            elif locator == 'CSS' and self.is_element_present(By.CSS_SELECTOR, element_tag):
                return WebDriverWait(dr, 15).until(lambda d: dr.find_element_by_css_selector(element_tag))
            else:
                logging.info(f"Error: Incorrect locator = {locator}")
        except Exception as e:
            logging.error(e)
        logging.info(f"Element not found with {locator} : {element_tag}")
        return None

    def is_element_present(self, how, what):
        """Check if an element is present"""
        try:
            self.driver.find_element(by=how, value=what)
        except NoSuchElementException:
            return False
        return True

    def __wait_for_element__(self, element_tag, locator, timeout=30):
        """Wait till element present. Max 30 seconds"""
        result = False
        self.driver.implicitly_wait(0)
        locator = locator.upper()
        for i in range(timeout):
            initTime = time()
            try:
                if locator == 'ID' and self.is_element_present(By.ID, element_tag):
                    result = True
                    break
                elif locator == 'NAME' and self.is_element_present(By.NAME, element_tag):
                    result = True
                    break
                elif locator == 'XPATH' and self.is_element_present(By.XPATH, element_tag):
                    result = True
                    break
                elif locator == 'CSS' and self.is_element_present(By.CSS_SELECTORS, element_tag):
                    result = True
                    break
                else:
                    logging.info(f"Error: Incorrect locator = {locator}")
            except Exception as e:
                logging.error(e)
                print(f"Exception when __wait_for_element__ : {e}")

            sleep(1 - (time() - initTime))
        else:
            print(
                f"Timed out. Element not found with {locator} : {element_tag}")
        self.driver.implicitly_wait(DEFAULT_IMPLICIT_WAIT)
        return result

    def __type_slow__(self, element_tag, locator, input_text=''):
        """Type the given input text"""
        try:
            self.__wait_for_element__(element_tag, locator, 5)
            element = self.__get_element__(element_tag, locator)
            actions = ActionChains(self.driver)
            actions.click(element).perform()
            # element.send_keys(input_text)
            for s in input_text:
                element.send_keys(s)
                # sleep(uniform(0.005, 0.02))

        except Exception as e:
            logging.error(e)
            print(f'Exception when __typeSlow__ : {e}')

    def __random_sleep__(self, minimum=2, maximum=7):
        t = randint(minimum, maximum)
        logging.info(f'Wait {t} seconds')
        sleep(t)

    def __scrolldown__(self):
        self.driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);")

    def __scrollToBottom__(self):
        self.driver.execute_script(
            "document.documentElement.scrollTo(0, document.documentElement.scrollHeight);")

    def teardown(self):
        self.driver.close()
        self.driver.quit()

    def extract_cookies(self, cookie=""):
        """从浏览器或者request headers中拿到cookie字符串，提取为字典格式的cookies"""
        cookies = {i.split("=")[0]: i.split("=")[1] for i in cookie.split("; ")}
        return cookies

    def __find_and_click(self, name, locator):
        locator = locator.upper()
        if not self.__wait_for_element__(name, locator, 10):
            print(name + ' field not visible')
        else:
            self.__get_element__(
                name, locator).click()

    def mainLoop(self):
        try:
            self.driver.get('https://youtube.com')
            if not self.__wait_for_element__(self.selectors['search_bar'], 'xpath', 10):
                print('search bar field not visible')
            else:
                self.__get_element__(
                    self.selectors['search_bar'], "xpath").click()
                self.__get_element__(
                    self.selectors['search_bar'], "xpath").send_keys(config["search_word"])
                self.__get_element__(
                    self.selectors['search_button'], 'xpath').click()

            self.__random_sleep__(5, 10)

            self.__find_and_click(self.selectors['filter_menu'], 'xpath')

            self.__random_sleep__()

            if config["recent_time"] == "year":
                self.__find_and_click(self.selectors["year"], "xpath")
            elif config["recent_time"] == "month":
                self.__find_and_click(self.selectors["month"], "xpath")
            elif config["recent_time"] == "week":
                self.__find_and_click(self.selectors["week"], "xpath")
            elif config["recent_time"] == "day":
                self.__find_and_click(self.selectors["day"], "xpath")

            channelEles = self.driver.find_elements_by_xpath(self.selectors["channel_link"])

            links = {}
            count = 0
            index = 0
            while count < config["nums"]:
                self.__scrollToBottom__()
                self.__random_sleep__()
                self.__scrollToBottom__()

                channelEles = self.driver.find_elements_by_xpath(self.selectors["channel_link"])

                for ele in channelEles[index:]:
                    index += 1
                    url = ele.get_attribute("href")
                    if url in links:
                        print(url)
                    else:
                        count += 1
                        links[str(url)] = 1

                if self.is_element_present('xpath', self.selectors["noMore"]):
                    self.is_no_more_result = True
                    print("arrive page bottom, no more result")
                    break

            print("link collect finish, start to fetch")
            for link in links:
                self.__random_sleep__(self.fetchInterval, self.fetchInterval)

                detailLink = f'{link}/about'

                desc = ""
                subCount = ""
                title = ""
                location = ""
                # open new tab for fetch info
                newWindow = f'window.open("{detailLink}")'
                self.driver.execute_script(newWindow)
                handles = self.driver.window_handles
                self.driver.switch_to.window(handles[len(handles) - 1])
                self.__random_sleep__(5, 10)
                try:
                    subCountNode = self.__get_element__(self.selectors["sub_count"], "xpath")
                    if subCountNode is not None:
                        subCount = self.transformNum(subCountNode.text)
                        if subCount < self.minFans:
                            self.driver.close()
                            self.driver.switch_to.window(handles[0])
                            continue

                    descNode = self.__get_element__(self.selectors["desc"], "xpath")
                    if descNode is not None:
                        desc = descNode.text

                    locationNode = self.__get_element__(self.selectors["location"], "xpath")
                    if locationNode is not None:
                        location = locationNode.text
                    titleNode = self.__get_element__(self.selectors["title"], "xpath")
                    if titleNode is not None:
                        title = titleNode.text
                    extraLinks = self.driver.find_elements_by_xpath(self.selectors["extLink"])
                    extraLinksStr = ""
                    if extraLinks is not None:
                        for extraLink in extraLinks:
                            extraLinksStr += f'{extraLink.text}: {extraLink.get_attribute("href")}\n'
                    self.excelData.append(
                        [title, subCount, link, location, desc, extraLinksStr])
                    self.driver.close()
                    self.driver.switch_to.window(handles[0])
                    print(f'fetch success, current process: {len(self.excelData) - 1}|title: {title}')

                except Exception as e:
                    logging.error(e)

            self.__save_excel(self.excelData)
            if self.is_no_more_result:
                print("no more result stop|fetch data nums: " + str(len(self.excelData)) + "|target nums: " + str(
                    config["nums"]))


        except Exception as e:
            print(str(e))
            self.__save_excel(self.excelData)
            logging.error(e)

    def transformNum(self, numStr):
        try:
            num = re.search(r'[.\d]+', numStr)
            unit = re.search(r'[a-zA-Z]', numStr)
            if unit is not None:
                if unit.group().lower() == 'k':
                    return float(num.group()) * 1000
                elif unit.group().lower() == 'm':
                    return float(num.group()) * 1000000
                elif unit.group().lower() == 'b':
                    return float(num.group()) * 1000000000
                else:
                    return int(num.group())
            else:
                return int(num.group())

        except Exception as e:
            print(str(e))

    def __save_excel(self, data):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for row in data:
            sheet.append(row)
        workbook.save('userinfo_' + str(time()) + '.xlsx')
        print("save into excel finished")


    def __waiting_for_stop_bot(self):
        cmd = input()
        if cmd == "stop":
            self.__save_excel(self.excelData)
            print("stop the bot!")
            exit(0)
