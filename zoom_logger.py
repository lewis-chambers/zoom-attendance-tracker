from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common import exceptions as EX
from datetime import datetime
import time
import traceback
import os
from os import path
import json
import xlsxwriter

def check_cookie_message():
    # accepting cookies
    try:
        WebDriverWait(driver, 1).until(
            EC.element_to_be_clickable((By.ID, 'onetrust-accept-btn-handler'))
        ).click()
    except EX.TimeoutException:
        pass

def log_in():
    login_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'inputname'))
    )

    check_cookie_message()

    # logging in
    login_input.clear()
    login_input.send_keys(your_name, Keys.ENTER)

def move_to_and_click(element):
    action = ActionChains(driver)
    action.move_to_element(element)
    action.perform()
    element.click()

def accept_terms():
    WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'wc_agree1'))
        ).click()

def enter_passcode():
    passcode_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, 'inputpasscode'))
    )

    password = input("Meeting Password: ")

    passcode_input.clear()
    passcode_input.send_keys(password, Keys.ENTER)

def get_participants():

    if 'participants-ul' not in driver.page_source:
        # Getting buttons
        buttons = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'button'))
        )

        participants_button = None
        more_button = driver.find_element(By.ID, 'moreButton')
        for _, button in enumerate(buttons):
            aria_str = button.get_attribute('aria-label')

            if aria_str is not None and 'participants' in aria_str:
                participants_button = button

        if participants_button is None:
            move_to_and_click(more_button)
            participants_button = driver.find_element(By.PARTIAL_LINK_TEXT, 'Participants')

        move_to_and_click(participants_button)

    collected_names = driver.find_elements(By.CLASS_NAME, 'participants-item__display-name')
 
    return [x.get_attribute('innerText') for x in collected_names]

def update_participants():

    existing_participants = participant_dict.keys()
    participant_names = get_participants()

    unknown_names = [x for x in participant_names if x not in existing_participants]
    missing_names = [x for x in existing_participants if x not in participant_names]
    left_names = [key for key, val in participant_dict.items() if len(val['leave_time']) != 0]
    rejoined_names = [x for x in left_names if x in participant_names]

    current_time = datetime.now().strftime('%Y/%m/%d, %H:%M:%S')

    print('Update: {}\n'.format(current_time))

    # Adding new names and start time to dict
    for name in unknown_names:
        participant_dict[name] = {'join_time': [current_time], 'leave_time': []}

    for name in missing_names:
        if len(participant_dict[name]['join_time']) != len(participant_dict[name]['leave_time']):
            participant_dict[name]['leave_time'].append(current_time)

    for name in rejoined_names:
        if len(participant_dict[name]['join_time']) == len(participant_dict[name]['leave_time']):
            participant_dict[name]['join_time'].append(current_time)

    for p in participant_dict.items():
        print(p)
    print('')

    last_participants = participant_names
    return last_participants

def page_id():
    try:
        source = driver.page_source
        body_text = driver.find_element(By.TAG_NAME, 'body').text

        if 'meeting-client' in source:
            if 'This meeting has been ended by host' in body_text:
                new_page = 603
            else:
                new_page = 4
        elif 'Your Name' in body_text:
            new_page = 1
        elif 'To use Zoom, you need to agree to the' in source:
            new_page = 2
        elif 'inputpasscode' in source:
            new_page = 3
        elif 'Thank you for attending the meeting' in source:
            new_page = 5
        elif 'The meeting has not started' in source:
            new_page = 600
        elif 'Joining Meeting' in source:
            new_page = 601
        elif 'This meeting link is invalid' in source:
            new_page = 602
        else:
            new_page = 0

        return new_page
    except:
        traceback.format_exc()

def log_participants():
    with open(output_json, 'w') as f:
        json.dump(participant_dict, f, indent=4)

    with open(output_json_backup, 'w') as f:
        json.dump(participant_dict, f, indent=4)

def initialise_directory():
    if not path.isdir(output_path):
        os.mkdir(output_path)

    if not path.isdir(run_path):
        os.mkdir(run_path)

def query_meeting_id():
    while True:
        meeting = input('Meeting ID: ').replace(' ', '')

        try:
            if len(meeting) not in [10, 11]:
                raise ValueError
            meeting = int(meeting)
            break
        except ValueError:
            print("Invalid meeting ID")
            pass

    return meeting

def finalise_end_times():
    current_time = datetime.now().strftime('%Y/%m/%d, %H:%M:%S')
    for item, value in participant_dict.items():
        if len(value['join_time']) != len(value['leave_time']):
            participant_dict[item]['leave_time'].append(current_time)
            log_participants()

def export_to_csv():
    with xlsxwriter.Workbook(output_excel) as workbook:
        worksheet = workbook.add_worksheet()

        headers = ['Participant', 'Join Time', 'Leave Time']
        for i, header in enumerate(headers):
            worksheet.write(0, i, header)

        row = 1
        col_width = [len(header[0]), len(header[1]), len(header[2])]

        for key, val in participant_dict.items():
            if len(key) > col_width[0]:
                col_width[0] = len(key)

            worksheet.write(row, 0, key)
            join_times, end_times = val.values()
            for _, (joined, left) in enumerate(zip(join_times, end_times)):
                worksheet.write(row, 1, joined)
                worksheet.write(row, 2, left)

                row += 1
                if len(joined) > col_width[1]:
                    col_width[1] = len(joined)
                if len(left) > col_width[2]:
                    col_width[2] = len(left)

        for i, val in enumerate(col_width):
            worksheet.set_column(i, i, col_width[i] + 2)


if __name__ == '__main__':
    cur_path = path.dirname(path.realpath(__file__))
    driver_path = path.join(cur_path, 'drivers','chromedriver')
    output_path = path.join(cur_path, 'output')
    time_now = datetime.now()
    time_str = time_now.strftime('%Y-%m-%d, %H;%M;%S')

    run_path = path.join(output_path, time_str)
    output_json = path.join(run_path, 'participants.txt')
    output_json_backup = path.join(run_path, 'participants_backup.txt')
    output_excel = path.join(run_path, 'participants.xlsx')

    options = webdriver.ChromeOptions()

    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36')
    #options.add_argument('window-size=1920,1080')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument('--disable-blink-features=AutomationControlled')
    prefs = {'credentials_enable_service': False,
             'profile.password_manager_enabled': False}

    options.add_experimental_option('prefs', prefs)

    meeting_ID = query_meeting_id()
    your_name = input("Your Name: ")

    sign_in_answer = ''
    sign_in = False
    while sign_in_answer.lower() not in ['y','n']:
        sign_in_answer = input('Sign in? (y/n): ')

    if sign_in_answer.lower() == 'y':
        sign_in = True
        #input("Hit enter when signed in: ")
    with webdriver.Chrome(service=Service(executable_path=driver_path), options=options) as driver:
        current_page = 0
        participant_dict = {}
        initialise_directory()
        last_participants = []
        while True:
            try:
                winds = driver.window_handles

                if len(winds) == 0:
                    raise EX.WebDriverException
            except EX.WebDriverException:
                print("Web driver closed")
                if path.isfile(output_json):
                    finalise_end_times()
                    export_to_csv()
                break

            try:
                page_counter = 0
                while page_counter < 5:
                    last_page = current_page
                    current_page = page_id()

                    if last_page == current_page:
                        page_counter += 1
                    else:
                        page_counter = 0
                    time.sleep(0.5)

                if current_page == 0:
                    # Logging In
                    sign_in_url = 'https://zoom.us/signin'
                    login_url = 'https://zoom.us/wc/join/{}'.format(meeting_ID)
                    if sign_in:
                        driver.get(sign_in_url)
                        input("Hit enter when logged in: ")
                        sign_in = False
                    else:
                        driver.get(login_url)
                    
                elif current_page == 1:
                    log_in()
                elif current_page == 2:
                    # Agree to privacy policy
                    accept_terms()
                elif current_page == 3:
                    # Enter meeting passcode
                    enter_passcode()
                elif current_page == 4:
                    last_paricipants = update_participants()
                    log_participants()
                    time.sleep(5)
                elif current_page == 5:
                    print("Meeting left")
                    if path.isfile(output_json):
                        print("Adding final exit times")
                        finalise_end_times()
                        print('Exporting to XLSX')
                        export_to_csv()
                    break
                elif current_page == 600:
                    print("Meeting not started")
                    time.sleep(5)
                elif current_page == 601:
                    print("Starting Meeting")
                    time.sleep(5)
                elif current_page == 602:
                    print("Invalid Meeting")
                    meeting_ID = query_meeting_id()
                    login_url = 'https://zoom.us/wc/join/{}'.format(meeting_ID)
                    driver.get(login_url)
                elif current_page == 603:
                    if path.isfile(output_json):
                        print("Adding final exit times")
                        finalise_end_times()
                        print('Exporting to XLSX')
                        export_to_csv()
                    break
                else:
                    print("Unknown page")
            except EX.StaleElementReferenceException:
                pass
            except Exception as e:
                traceback.print_exc()
