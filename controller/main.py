# *************************************************
# LOOKER STUDIO - RETRIEVING DATA
# *************************************************

# Importing libraries
import pandas as pd
import os
import time
import sys

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from datetime import datetime
from selenium.common.exceptions import NoSuchElementException

# Avoiding creation of .pyc file
sys.dont_write_bytecode = True


# *************************************************
# Defining functions
# *************************************************
def log_content(_content, _filename):
    f = open(_filename, "a")
    f.write("{0} -- {1}\n".format(datetime.now().strftime("%Y-%m-%d %H:%M"), _content))
    print(_content)
    f.close()


def set_driver(_driver_path, _profile, _desired_capabilities, _options):
    _driver = webdriver.Firefox(executable_path=_driver_path,
                                firefox_profile=_profile,
                                desired_capabilities=_desired_capabilities,
                                options=_options)
    return _driver


def set_window_action(_webpage, _driver):
    _driver.get(_webpage)
    _driver.switch_to.window(_driver.current_window_handle)

    action = ActionChains(_driver)
    return action


def click_button(_path_to_element, __action, _driver):
    _driver = _driver
    _action = __action
    _action.move_to_element(_driver.find_element(by=By.XPATH, value=_path_to_element)).click().perform()


def click_button_css_selector(_path_to_element, __action, _driver):
    _driver = _driver
    _action = __action
    _action.move_to_element(_driver.find_element(by=By.CSS_SELECTOR, value=_path_to_element)).click().perform()


def scroll_and_click(_path_to_element, _driver):
    element_date = _driver.find_element(by=By.XPATH, value=_path_to_element)
    element_date.location_once_scrolled_into_view
    element_date.click()


def wait_to_load(_path_to_element, _driver, _time_to_wait=60, _wait_type='visual_invisibile'):
    if _wait_type == 'visual_invisibile':
        __element = WebDriverWait(_driver, _time_to_wait).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, _path_to_element))
        )
    elif _wait_type == 'element_occurence':
        __element = WebDriverWait(_driver, _time_to_wait).until(
            EC.element_to_be_clickable((By.XPATH, _path_to_element))
        )
    return __element


def insert_text(_path_to_element, _text, _driver):
    _driver.find_element(by=By.XPATH, value=_path_to_element).clear()
    _driver.find_element(by=By.XPATH, value=_path_to_element).send_keys(_text)


def checkbox_click(_path_to_element, click_action, _driver):
    try:
        _driver.find_element(by=By.XPATH, value=_path_to_element)
    except NoSuchElementException:
        return None
    click_action


def download_wait(_path_to_downloads, _file_name, _time=30):
    seconds = 0
    dl_wait = True
    _result_find = False
    _file_name += '.csv'
    while dl_wait and seconds < _time:
        time.sleep(1)
        files = os.listdir(_path_to_downloads)
        for fname in files:
            if fname == _file_name:
                dl_wait = False
                _result_find = True
        seconds += 1
    return _result_find


def xpath_sub_id(_path_to_main_element, _id_opt, _text_to_replace, _part_new1, _part_new2, _driver):
    old_xpath = _driver.find_element(by=By.XPATH, value=_path_to_main_element).get_attribute('id')
    id_opt = int(_id_opt)
    id_opt += int(old_xpath.replace(_text_to_replace, ""))
    final_xpath = _part_new1 + str(id_opt) + _part_new2
    return final_xpath


# Defining path to the base config file
baseconfig = "/configs/base_config.xlsx"

# *************************************************
# Main Function
# *************************************************


def gds_automation(jobconfig):
    # reading base configuration file
    try:
        baseconfig_path = '/'.join(os.getcwd().split('\\')) + baseconfig
        _baseconfig = pd.read_excel(baseconfig_path)
    except:
        print(f"Error with reading Excel file: {baseconfig}")

    # reading job configuration file
    try:
        config_df_path = '/'.join(os.getcwd().split('\\')) + jobconfig
        _config_df = pd.read_excel(config_df_path)
    except:
        print(f"Error with reading Excel file: {jobconfig}")

    # getting configuration addresses - driver, firefox profile, log file
    global __LOG_FILE_SUCCESS_PATH
    global __LOG_FILE_ERROR_PATH
    _LOG_FILE_SUCCESS_NAME = '/log.txt'
    _LOG_FILE_ERROR_NAME = '/log_error.txt'
    _DRIVER_FILEPATH = _baseconfig.loc[0, 'INPUT_PATH']
    _PROFILE_PATH = _baseconfig.loc[1, 'INPUT_PATH']
    _DOWNLOAD_FOLDER_PATH = _baseconfig.loc[2, 'INPUT_PATH']
    _BINARY_LOCATION_PATH = _baseconfig.loc[3, 'INPUT_PATH']
    __LOG_FILE_PATH_SUCCESS = _DOWNLOAD_FOLDER_PATH + _LOG_FILE_SUCCESS_NAME
    __LOG_FILE_PATH_ERROR = _DOWNLOAD_FOLDER_PATH + _LOG_FILE_ERROR_NAME

    # setting selenium settings
    profile = webdriver.FirefoxProfile(_PROFILE_PATH)
    profile.set_preference("dom.webdriver.enabled", False)
    profile.set_preference('useAutomationExtension', False)
    profile.update_preferences()
    desired = DesiredCapabilities.FIREFOX
    opts = webdriver.FirefoxOptions()
    opts.headless = True
    opts.binary_location = _BINARY_LOCATION_PATH
    log_content("Settings has been set", __LOG_FILE_PATH_SUCCESS)

    for row_nr in range(len(_config_df.index)):

        # Getting dashboard url, list of countries and xpaths to particular visuals
        _url = _config_df.loc[row_nr, 'URL']
        _country_filter = int(_config_df.loc[row_nr, 'FILTER_BY_COUNTRY'])
        _countries = _config_df.loc[row_nr, 'COUNTRIES']
        _page_load_element = _config_df.loc[row_nr, 'PAGE_LOAD_ELEMENT']
        _progress_bar = _config_df.loc[row_nr, 'PROGRESS_BAR']
        _date_filter = int(_config_df.loc[row_nr, 'FILTER_BY_DATE'])
        _date1 = _config_df.loc[row_nr, 'DATE1']
        _date2 = _config_df.loc[row_nr, 'DATE2']
        _date3 = _config_df.loc[row_nr, 'DATE3']
        _date3_1 = _config_df.loc[row_nr, 'DATE31']
        _date3_2 = _config_df.loc[row_nr, 'DATE32']
        _date3_3 = _config_df.loc[row_nr, 'DATE33']
        _date4 = _config_df.loc[row_nr, 'DATE4']
        _cnt_button = _config_df.loc[row_nr, 'CNT1']
        _cnt_button_checkbox = _config_df.loc[row_nr, 'CNT2']
        _input_country = _config_df.loc[row_nr, 'CNT3']
        _cnt_button_checkbox_input = _config_df.loc[row_nr, 'CNT4']
        _visual_to_check = _config_df.loc[row_nr, 'VISUAL_TO_CHECK']
        _export_menu = _config_df.loc[row_nr, 'EXPORT_MENU']
        _export_input = _config_df.loc[row_nr, 'EXPORT_INPUT']
        _file_title = _config_df.loc[row_nr, 'FILE_TITLE']
        _export_button = _config_df.loc[row_nr, 'EXPORT_BUTTON']

        # Setting Firefox driver
        driver_used = set_driver(_driver_path=_DRIVER_FILEPATH, _profile=profile,
                                 _desired_capabilities=desired, _options=opts)

        # Setting ActionChain
        _action = set_window_action(_webpage=_url, _driver=driver_used)

        # Wait for the dashboard to be load
        time.sleep(4)
        try:
            print("trying wait_to_load function")
            wait_to_load(_path_to_element=_page_load_element, _wait_type='element_occurence', _driver=driver_used)
            log_content("Site has been launched", __LOG_FILE_PATH_SUCCESS)
        except:
            log_content("Error with loading page and finding main element. Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        # If date filter is TRUE in the config file, applying country filter
        if _date_filter == 1:
            try:
                click_button(_path_to_element=_date1, __action=_action, _driver=driver_used)
                click_button(_path_to_element=_date2, __action=_action, _driver=driver_used)

                log_content("Date filters has been initialized", __LOG_FILE_PATH_SUCCESS)
            except:
                log_content("Problem with expanding date menu filters. Process stopped.", __LOG_FILE_PATH_ERROR)
                break

                # Setting date filters
            try:
                date_id = xpath_sub_id(_path_to_main_element=_date2, _id_opt=_date3, _text_to_replace=_date3_1,
                                   _part_new1=_date3_2, _part_new2=_date3_3, _driver=driver_used)
                scroll_and_click(_path_to_element=date_id, _driver=driver_used)
                click_button(_path_to_element=_date4, __action=_action, _driver=driver_used)
                log_content("Date filters has been set", __LOG_FILE_PATH_SUCCESS)
            except:
                log_content("Problem with setting date filters. Process stopped.", __LOG_FILE_PATH_ERROR)
                break
        elif _date_filter == 0:
            log_content("Date filter was not applied", __LOG_FILE_PATH_SUCCESS)

        # If country filter is TRUE in the config file, applying country filter
        if _country_filter == 1:

            # Wait for the dashboard to load
            try:
                wait_to_load(_path_to_element=_cnt_button, _wait_type='element_occurence', _driver=driver_used)
                log_content("Dashboard refreshed", __LOG_FILE_PATH_SUCCESS)
            except:
                log_content("Problem with refreshing dashboard after setting date filters. Process stopped.", __LOG_FILE_PATH_ERROR)
                break

                # Turning on country filter
            try:
                _action.move_to_element(driver_used.find_element(by=By.XPATH, value=_cnt_button)).click().perform()
                _action.pause(1)
                _action.move_to_element(driver_used.find_element(by=By.XPATH, value=_cnt_button_checkbox)).click().perform()
                log_content("Filtering by country started", __LOG_FILE_PATH_SUCCESS)
            except:
                log_content("Problem with turning on country filter. Process stopped.", __LOG_FILE_PATH_ERROR)
                break

            # Creating country list
            try:
                _countries = _countries.split(";")
            except:
                log_content("Wrong country list. Check split character. Process stopped.", __LOG_FILE_PATH_ERROR)
                break

            # Filtering by country
            # Finishing filtering by country
            try:
                for country in _countries:
                    insert_text(_path_to_element=_input_country, _text=country, _driver=driver_used)
                    checkbox_click(_path_to_element=_cnt_button_checkbox_input,
                                   click_action=click_button(_path_to_element=_cnt_button_checkbox_input, __action=_action,
                                                             _driver=driver_used), _driver=driver_used)
                log_content("Filtering by country finished", __LOG_FILE_PATH_SUCCESS)
            except:
                log_content("Problem with finishing country filter. Process stopped.", __LOG_FILE_PATH_ERROR)
                break

            # Wait for the dashboard to be load - Part 1
            try:
                click_button(_path_to_element=_visual_to_check, __action=_action, _driver=driver_used)
                log_content("Dashboard refreshed after filtering by country", __LOG_FILE_PATH_SUCCESS)
            except:
                log_content("Error with finding the main element. Process stopped.", __LOG_FILE_PATH_ERROR)
                break

        elif _country_filter == 0:
            log_content("Country filter was not applied", __LOG_FILE_PATH_SUCCESS)

        # Wait for the dashboard to be load - Part 2
        try:
            _action.pause(3)
            wait_to_load(_path_to_element=_progress_bar, _driver=driver_used)
            log_content("Loading is finished", __LOG_FILE_PATH_SUCCESS)
        except:
            log_content("Problem with waiting for progress bar to disappear. Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        # Turning on menu
        try:
            _action.context_click().perform()
            _action.pause(1)
            log_content("Menu has been opened", __LOG_FILE_PATH_SUCCESS)
        except:
            log_content("Problem with opening menu. Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        # Using export option
        try:
            time.sleep(2)
            click_button_css_selector(_path_to_element=_export_menu, __action=_action, _driver=driver_used)
            log_content("Export button has been clicked", __LOG_FILE_PATH_SUCCESS)
        except:
            log_content("Problem with finding export button in menu. Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        # Naming exported csv file
        try:
            id_time = time.strftime('%d-%m-%y_%H.%M.%S', time.localtime())
            _file_title += "_" + id_time
            insert_text(_path_to_element=_export_input, _text=_file_title, _driver=driver_used)
            _action.pause(1)
            log_content("Exported CSV file has been named", __LOG_FILE_PATH_SUCCESS)
        except:
            log_content("Problem with naming CSV file. Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        # Starting download process
        try:
            click_button_css_selector(_path_to_element=_export_button, __action=_action, _driver=driver_used)
            log_content("Download process has been launched", __LOG_FILE_PATH_SUCCESS)
        except:
            log_content("Error with downloading the file. Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        if download_wait(_DOWNLOAD_FOLDER_PATH, _file_title):
            log_content("File: " + _file_title + " has been successfully downloaded", __LOG_FILE_PATH_SUCCESS)
        else:
            log_content("Error with downloading file: " + _file_title + " Process stopped.", __LOG_FILE_PATH_ERROR)
            break

        driver_used.close()

    driver_used.quit()
    print("Script has been executed successfully. You can close this window.")


if __name__ == "__name__":
    gds_automation()
