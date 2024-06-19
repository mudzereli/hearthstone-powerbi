from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime
from colorama import just_fix_windows_console
from colorama import Fore, Back, Style
import csv
import win32gui
import win32process
import win32con
import sys

# Set up Colors
just_fix_windows_console()
FY = Fore.YELLOW
FG = Fore.GREEN
FR = Fore.RED
FC = Fore.CYAN
FB = Fore.BLUE
FW = Fore.WHITE
FM = Fore.MAGENTA
FD = Fore.BLACK
SB = Style.BRIGHT
SN = Style.NORMAL
SD = Style.DIM
#timestamp = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
timestamp = datetime.now().strftime('%m/%d/%Y')

# Print with Timestamp
def tprint(message):
    tstamp = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
    print(f'{SB}{FW}[{SD}{FD}{tstamp}{SB}{FW}] {SN}{message}')

# Exception Hook
def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    input(f'{FR}ERROR: Press key to exit.{FW}')
    sys.exit(-1)

# Class Definitions
def get_window_handle(window_title):
    """
    Get the window handle by its title.
    """
    handle = win32gui.FindWindow(None, window_title)
    return handle

def bring_window_to_front(window_handle):
    """
    Bring the window associated with the given handle to the front.
    """
    win32gui.ShowWindow(window_handle, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(window_handle)

# Boolean Input Helper
def get_bool(prompt):
    while True:
        try:
           return {"true":True,"false":False,"yes":True,"no":False,"y":True,"n":False}[input(f'{prompt}').lower()]
        except KeyError:
           tprint(f'{FR}{SB}Invalid input please enter True or False!{SN}')

# Define deck class
class Deck:
    def dump(self):
        tprint('---------------------------')
        tprint(f'Deck Timestamp = {self.timestamp}')
        tprint(f'Format = {self.format}')
        tprint(f'RankRange = {self.rankrange}')
        tprint(f'Archetype = {self.archetype}')
        tprint(f'Class = {self.classname}')
        tprint(f'Win_Rate = {self.winrate}')
        tprint(f'Games = {self.games}')
        tprint(f'Duration = {self.duration}')
        tprint(f'Dust = {self.dust}')
        tprint(f'AllCards = {self.cardlist}')
        tprint(f'Link = {self.link}')

# Hook Exceptions to Not Close on Error
sys.excepthook = show_exception_and_exit
# Set up Variables
tprint(f'{FY}{SB}Setting Up Variables{SN}')
decks = []
csvpath = "C:\\Users\\mudz\\Documents\\pyhsreplay.csv"
# Prompt for Full Read
fullread = get_bool(FC + "Read all Rank Ranges? " + FW)
tprint(f'{FM}{SB}Reading All Rank Ranges: {FW}{fullread}{SN}')
# Set up URLs
urls = ["https://hsreplay.net/decks/","https://hsreplay.net/decks/#gameType=RANKED_WILD"]
if fullread:
    urls.append("https://hsreplay.net/decks/#rankRange=GOLD")
    urls.append("https://hsreplay.net/decks/#rankRange=SILVER")
    urls.append("https://hsreplay.net/decks/#rankRange=BRONZE")
    urls.append("https://hsreplay.net/decks/#gameType=RANKED_WILD&rankRange=GOLD")
    urls.append("https://hsreplay.net/decks/#gameType=RANKED_WILD&rankRange=SILVER")
    urls.append("https://hsreplay.net/decks/#gameType=RANKED_WILD&rankRange=BRONZE")
window_title = "C:\\WINDOWS\\py.exe"
window_handle = get_window_handle(window_title)
timeout = 30

# Load Selenium Driver
tprint(f'{FY}Loading Selenium Chrome Driver')
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument('--blink-settings=imagesEnabled=false')
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("user-data-dir=D:\\tmp\\Selenium")
chrome_options.add_extension('C:\\Users\\mudz\\AppData\\Roaming\\Opera Software\\Opera GX Stable\\Extensions\\kccohkcpppjjkkjppopfnflnebibpida\\1.58.0_0.crx')
#chrome_options.add_argument("--headless")
#chrome_options.add_argument("--disable-ads")
#chrome_options.add_argument("--disable-tracking")
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()
#if window_handle:
#    bring_window_to_front(window_handle)
#else:
#    tprint(FR + "Window not found.")

# Loop through URLs
for url in urls:
    # Check if Wild or Standard
    if "RANKED_WILD" in url:
        gameformat = "Wild"
    else:
        gameformat = "Standard"
    # Load and Setup Website in Selenium
    tprint(f'{FC}Loading Website: {FW}{url}')
    driver.execute_script("window.history.pushState('', '', '/')")
    driver.get(url)
    # Actually Wait for Page to Load
    try:
        # Using CSS_SELECTOR with a more robust selector that doesn't depend on dynamic class names
        element_present = EC.presence_of_element_located((By.CSS_SELECTOR, '#decks-container > main > div.page-with-banner-container > div > div > div > div > section > div.paging.text-center > nav > ul > li.visible-lg-inline.active'))
        WebDriverWait(driver, timeout).until(element_present)
    except TimeoutException:
        tprint(f'Timed out waiting for page to load')
    # driver.implicitly_wait(implicit_wait_time)
    # Find Group Elements
    rankrange = driver.find_element(By.CSS_SELECTOR, "#rank-range-filter > div > ul > li.selectable.selected.no-deselect").text
    lastUpdated = driver.find_element(By.XPATH, "//li[starts-with(normalize-space(),'Last updated')]/span[@class='infobox-value']/div[@class='tooltip-wrapper']/time").text
    try:
        page_count = int(driver.find_element(By.CLASS_NAME, "pagination").find_elements(By.CLASS_NAME, "visible-lg-inline")[-1].text)
    except NoSuchElementException:
        page_count = 1
    # Loop Through All Pages for Current Group
    tprint(f'{FG}Starting Scrape of {FW}{gameformat}{FG} - {FW}{rankrange}{FG} Data')
    tprint(f'{FM}Last Updated {FW}{lastUpdated}')
    for p in range(1,page_count+1):
        # Find Decks
        deck_tiles = driver.find_elements(By.CLASS_NAME, "deck-tile")
        # Loop Through Each Deck Tile
        for tile in deck_tiles:
            deck = Deck()
            deck.timestamp = timestamp
            deck.format = gameformat
            deck.rankrange = rankrange
            deck.archetype = tile.find_element(By.CLASS_NAME, "deck-name").text
            deck.link = tile.get_attribute("href")
            deck.classname = tile.get_attribute("data-card-class")
            deck.dust = tile.find_element(By.CLASS_NAME, "dust-cost").text
            deck.winrate = tile.find_element(By.CLASS_NAME, "win-rate").text
            deck.games = tile.find_element(By.CLASS_NAME, "game-count").text
            deck.duration = tile.find_element(By.CLASS_NAME, 'duration').text
            deck.cardlist = ''
            for card in tile.find_elements(By.XPATH, ".//div/div[6]/ul/li/div/a/div"):
                deck.cardlist += card.get_attribute("aria-label")
                deck.cardlist += ";"
            #deck.tprint()
            decks.append(deck)
            rj = len(str(page_count))
        tprint(f'{SB}{FB}page # {FW}{str(p).rjust(rj)}{FB} of {FW}{page_count} {FB}{SN}/{SB}{FB} format: {FW}{gameformat} {FB}{SN}/{SB}{FB} rank: {FW}{rankrange} {FB}{SN}/{SB}{FB} # of decks read: {FW}{len(decks)}{SN}')
        # Find and Click Next Page Button
        try:
            button = driver.find_element(By.XPATH, "//a[@class='weight-normal' and @title='Next page']")
            button.click()
        except NoSuchElementException:
            tprint(FR + "Next Page not found... Moving On")
driver.quit()

# Write Results to CSV
tprint(f'{FM}Writing results to: {FW}{csvpath}')
with open(csvpath, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    # Write Header Row
    writer.writerow(['Timestamp', 'Format', 'RankRange', 'Archetype', 'Class', 'Win_Rate', 'Games', 'Duration', 'Dust', 'AllCards', 'Link'])
    # Write All Decks
    for deck in decks:
        writer.writerow([deck.timestamp, deck.format, deck.rankrange, deck.archetype, deck.classname, deck.winrate, deck.games, deck.duration, deck.dust, deck.cardlist, deck.link])
# All Done!
tprint(f'{FC}{SB}All Done! # of decks written: {FW}{len(decks)}{SN}')
input(f'{FY}Press Enter key to close...')
