from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import StaleElementReferenceException
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
import time
import os

# Set up Colors
just_fix_windows_console()
FY = Fore.YELLOW
FG = Fore.GREEN
FR = Fore.RED
FC = Fore.CYAN
FB = Fore.BLUE
FW = Fore.WHITE
FM = Fore.MAGENTA
FX = Fore.BLACK
SB = Style.BRIGHT
SN = Style.NORMAL
SD = Style.DIM
timestamp = datetime.now().strftime('%m/%d/%Y')
deck_load = 31  # number of decks loaded at a time
deck_height = 86  # pixel height of each deck
scroll_size = deck_load * deck_height  # how much to scroll on each page


# Print with Timestamp
def tprint(message):
    tstamp = datetime.now().strftime('%m/%d/%Y %H:%M:%S')
    print(f'{SB}{FW}[{SD}{FX}{tstamp}{SB}{FW}] {SN}{message}')


# Exception Hook
def show_exception_and_exit(exc_type, exc_value, tb):
    import traceback
    traceback.print_exception(exc_type, exc_value, tb)
    input(f'{FR}ERROR: Press key to exit.{FW}')
    sys.exit(-1)


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
            return {"true": True, "false": False, "yes": True, "no": False, "y": True, "n": False}[input(f'{prompt}').lower()]
        except KeyError:
            tprint(f'{FR}{SB}Invalid input please enter True or False!{SN}')


# Class Definitions
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


# Function to scroll incrementally and capture decks
def scroll_and_capture_decks_incrementally(driver, decks, scroll_pause_time=0, scroll_increment=scroll_size, eocscroll=897):
    last_height = driver.execute_script("return window.scrollY")
    total_height = driver.execute_script("return document.body.scrollHeight")-eocscroll

    while True:
        # Capture current deck tiles
        capture_deck_tiles(driver, decks)
        rj = len(str(total_height))
        tprint(f'{SB}{FB}scroll: {FW}{str(last_height).rjust(rj)}{FB} of {FW}{total_height} {FB}{SN}/{SB}{FB} format: {FW}{gameformat} {FB}{SN}/{SB}{FB} rank: {FW}{rankrange} {FB}{SN}/{SB}{FB} # of decks read: {FW}{len(decks)}{SN}')

        # Remove duplicates
        decks, duplicates_count = remove_duplicates(decks)
        if duplicates_count > 0:
            tprint(f'{SB}{FC}duplicates removed: {FW}{duplicates_count}{SN}')

        # Scroll down by a small increment
        driver.execute_script(f"window.scrollBy(0, {scroll_increment});")
        time.sleep(scroll_pause_time)  # Wait for content to load

        # Check if the new scroll height is the same as before (i.e., reached the end)
        new_height = driver.execute_script("return window.scrollY")
        if new_height == last_height:
            break
        last_height = new_height
    return decks


# Function to capture deck tiles
def capture_deck_tiles(driver, decks):
    try:
        # Wait for deck tiles to be present
        deck_tiles = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "deck-tile"))
        )
        # tprint(f'{SB}{FM}# of deck tiles = {FW}{len(deck_tiles)}')
        for tile in deck_tiles:
            try:
                deck = Deck()
                deck.timestamp = timestamp
                deck.format = gameformat
                deck.rankrange = rankrange

                # Use explicit waits for each element
                deck.archetype = WebDriverWait(tile, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "deck-name"))
                ).text

                deck.link = tile.get_attribute("href")
                deck.classname = tile.get_attribute("data-card-class")

                deck.dust = WebDriverWait(tile, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "dust-cost"))
                ).text

                deck.winrate = WebDriverWait(tile, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "win-rate"))
                ).text

                deck.games = WebDriverWait(tile, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "game-count"))
                ).text

                deck.duration = WebDriverWait(tile, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "duration"))
                ).text

                # Handle card list with explicit wait
                cards = WebDriverWait(tile, 5).until(
                    EC.presence_of_all_elements_located((By.XPATH, ".//div/div[6]/ul/li/div/a/div"))
                )
                deck.cardlist = ';'.join(card.get_attribute("aria-label") for card in cards if card.get_attribute("aria-label"))

                decks.append(deck)

            except (TimeoutException, StaleElementReferenceException) as e:
                tprint(f'{FR}Failed to capture deck tile: {str(e)}{FW}')
                continue

    except TimeoutException:
        tprint(f'{FR}Timeout waiting for deck tiles{FW}')

    return decks


# Function to remove duplicates from decks list
def remove_duplicates(decks):
    seen_links = set()
    unique_decks = []
    duplicates_count = 0
    for deck in decks:
        if deck.link not in seen_links:
            seen_links.add(deck.link)
            unique_decks.append(deck)
        else:
            duplicates_count += 1
    return unique_decks, duplicates_count

# Hook Exceptions to Not Close on Error
sys.excepthook = show_exception_and_exit
# Set up Variables
tprint(f'{FY}{SB}Setting Up Variables{SN}')
decks = []
# Get the directory where the script is located
script_dir = os.path.dirname(__file__)
# Create the full path for the CSV file in the same directory as the script
csvpath = os.path.join(script_dir, "pyhsreplay.csv")
# Prompt for Full Read
fullread = get_bool(FC + "Read all Rank Ranges? " + FW)
tprint(f'{FM}{SB}Reading All Rank Ranges: {FW}{fullread}{SN}')
# Set up URLs
urls = ["https://hsreplay.net/decks/", "https://hsreplay.net/decks/#gameType=RANKED_WILD"]
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
driver = webdriver.Chrome(options=chrome_options)
driver.maximize_window()

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
        element_present = EC.presence_of_element_located((By.CSS_SELECTOR, '#decks-container > main > div.deck-list-wrapper.page-with-banner-container > div > div:nth-child(2) > div > section > ul'))
        WebDriverWait(driver, timeout).until(element_present)
    except TimeoutException:
        tprint(f'Timed out waiting for page to load')

    # Find Group Elements
    rankrange = driver.find_element(By.CSS_SELECTOR, "#rank-range-filter > div > ul > li.selectable.selected.no-deselect").text
    lastUpdated = driver.find_element(By.XPATH, "//*[@id=\"side-bar-data\"]/dl/div/dd/div/time").text

    # Scrape all data for current Group
    tprint(f'{FG}Starting Scrape of {FW}{gameformat}{FG} - {FW}{rankrange}{FG} Data')
    tprint(f'{FM}Last Updated {FW}{lastUpdated}')
    decks = scroll_and_capture_decks_incrementally(driver, decks)
    tprint(f'{SB}{FB}format: {FW}{gameformat} {FB}{SN}/{SB}{FB} rank: {FW}{rankrange} {FB}{SN}/{SB}{FB} # of decks read: {FW}{len(decks)}{SN}')
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
