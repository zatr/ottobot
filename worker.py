from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By


def hover_click(driver, element_hover, element_click):
    ActionChains(driver).move_to_element(element_hover).click(element_click).perform()


def element_wait(driver, element, by_attr, wait_seconds=10):
    return WebDriverWait(driver, wait_seconds).until(
        EC.presence_of_element_located((by_attr, element)))


def frame_wait(driver, frame):
    return WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it(frame)
    )


def open_browser():
    try:
        driver = webdriver.Firefox()
        return driver
    except:
        print 'Error: Could not start Selenium WebDriver (Firefox)'


def open_browser_connect_to_site(site, assert_text):
    dvr = webdriver.Firefox()
    dvr.get(site)
    assert assert_text == dvr.title
    return dvr


import settings


def login_analyst(driver):
    driver.switch_to.frame('main')
    login = driver.find_element_by_id('hlsys_button1')
    login.send_keys(Keys.RETURN)
    username = element_wait(driver, 'textsys_field3', By.ID)
    username.send_keys(settings.app_username)
    password = driver.find_element_by_id('textsys_field2')
    password.send_keys(settings.app_password)
    submit = driver.find_element_by_name('ctl02')
    submit.send_keys(Keys.RETURN)


def open_connect_login_analyst():
    driver = open_browser()
    driver.get(settings.app_url)
    login_analyst(driver)
    return driver


def go_to_request(driver, rid, base_url=settings.app_url):
    driver.get('%s/ReqInfo.aspx?sys_request_id=%i' % (base_url, rid))
    try:
        element_wait(driver, 'ctl00_ContentPlaceHolder1_textsys_field1', By.ID, 2)
        print '\nOpened Request', rid
        return True
    except:
        print 'Request not found:', rid


def go_to_item(driver, problem_or_change, item_id):
    driver.get('%s/%sInfo.aspx?sys_%s_id=%i' % (settings.app_url,
                                                problem_or_change,
                                                problem_or_change,
                                                item_id))


def client_login(driver):
    username = driver.find_element_by_id('req1')
    username.send_keys(settings.client_username)
    password = driver.find_element_by_id('req2')
    password.send_keys(settings.client_password)
    login = driver.find_element_by_class_name('blueLoginBtn')
    login.click()


def open_client_ticket(driver, ticket_id):
    tickets = element_wait(driver, 'Tickets', By.LINK_TEXT)
    tickets.click()
    driver.get(settings.ticket_url + ticket_id)


import gspread


def get_item_number_list(cell_range, product, version):
    try:
        print 'Connecting to Google Apps'
        gc = gspread.login(settings.google_user, settings.google_pass)
    except:
        raise Exception('Authentication or Connection issue to Google Apps.')
    document_name = settings.doc_name % (product, version)
    doc = gc.open(document_name)
    print 'Opened document:', document_name
    worksheet = doc.worksheet(settings.worksheet_name)
    print 'Opened worksheet:', settings.worksheet_name
    print 'Getting values from cell range:', cell_range
    cell_list = worksheet.range(cell_range)
    item_list = []
    for cell in cell_list:
        try:
            item = int(cell.value)
            item_list.append(item)
        except (TypeError, ValueError):
            pass
    print 'Found these values:'
    for item in item_list:
        print item
    return item_list


def rename_bug_to_problem(boc):
    if boc == 'bug':
        return 'problem'
    else:
        return boc


import pyodbc


def db_connect():
    db_connect_string = settings.db_connect_string
    try:
        return pyodbc.connect(''.join(db_connect_string))
    except:
        raise Exception('SQL Error: Failed to connect to server: ' +
                        db_connect_string[0] +
                        db_connect_string[1] +
                        db_connect_string[2] +
                        db_connect_string[3] +
                        'PWD=********;' +
                        db_connect_string[5])


import inspect


def whoisparent():
    return inspect.stack()[2][3]


def exec_sql_read(query):
    cursor = db_connect().cursor()
    try:
        cursor.execute(query)
        results = []
        while 1:
            row = cursor.fetchone()
            if not row:
                break
            results.append(row)
        return results
    except:
        print 'SQL Error: Query failed in function: %s' % whoisparent()
    finally:
        cursor.close()


def sql_select_linked_requests(item_number, bug_or_change):
    poc = rename_bug_to_problem(bug_or_change)
    sql = ('select req_problem_change.sys_%s_id, request.sys_request_id, '
           'usr_Customer_Name, usr_Cust_Email, %s.sys_%s_summary '
           'from request join req_problem_change '
           'on request.sys_request_id = req_problem_change.sys_request_id '
           'join %s '
           'on req_problem_change.sys_%s_id = %s.sys_%s_id '
           'where req_problem_change.sys_%s_id = %i') % (poc, poc, poc, poc,
                                                         poc, poc, poc, poc,
                                                         item_number)
    return sql


from collections import defaultdict, namedtuple


def get_psummary_linked_requests(items, bug_or_change):
    linked_requests = defaultdict(list)
    psummary_linked_req = namedtuple('LinkedRequestList', ['problem_summary',
                                                           'requests'])
    all_linked_rids = []
    print '\nLoading requests linked to %ss:' % bug_or_change, items
    cursor = db_connect().cursor()
    for item_number in items:
        try:
            cursor.execute(sql_select_linked_requests(item_number, bug_or_change))
        except:
            print 'ERROR: SQL query failed.'
        request_list = []
        while 1:
            row = cursor.fetchone()
            if not row:
                break
            request_list.append(row[1])
            item_summary = row[4]
        if request_list:
            print '%s %i linked requests:' % (bug_or_change, item_number), request_list
            for rid in request_list:
                all_linked_rids.append(rid)
            linked_requests[item_number] = psummary_linked_req(item_summary, request_list)
        else:
            print 'No requests linked to %s %i!' % (bug_or_change, item_number)
    cursor.close()
    print '\nAll linked requests queued for notification:'
    for rid in sorted(all_linked_rids):
        print rid
    return linked_requests


def collect_cust_fname_email(driver):
    usr_customer_name = element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield2', By.ID)
    first_name = usr_customer_name.get_attribute('value').split()[0]
    usr_cust_email = element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield3', By.ID)
    email_address = usr_cust_email.get_attribute('value')
    return first_name, email_address


def generate_product_update_comment_request_on_notification(driver):
    def get_comment():
        template = open(settings.templates_path + '/on_notification_notification.html')
        text = template.read()
        template.close()
        return text
    first_name, email_address = collect_cust_fname_email(driver)
    comment_text = get_comment() % (first_name,
                                    release_info.product_name,
                                    release_info.update_version,
                                    settings.client_url,
                                    )
    return email_address, comment_text


def generate_product_update_comment_linked_request(driver, item_number, item_summary, ticket_id, update_leapfile):
    boc = release_info.bug_or_change
    first_name, email_address = collect_cust_fname_email(driver)
    variables = (first_name,
                 release_info.product_name,
                 release_info.update_version,
                 item_number,
                 item_summary,
                 settings.client_url,
                 )
    if ticket_id:
        template = open(os.path.join(settings.templates_path, '%s_notification_ticket.html' % boc))
    elif update_leapfile:
        template = open(os.path.join(settings.templates_path, 'leapfile_update_notification.html'))
        variables = variables[:-1]
    else:
        template = open(os.path.join(settings.templates_path, '%s_notification.html' % boc))
    text = template.read()
    template.close()
    comment_text = text % variables
    return email_address, comment_text


def generate_leapfile_email_body(driver, item_number, item_summary):
    boc = release_info.bug_or_change
    template = open(os.path.join(settings.templates_path, '%s_leapfile_email_body.html' % boc))
    text = template.read()
    template.close()
    first_name, email_address = collect_cust_fname_email(driver)
    email_body = text % (first_name,
                         release_info.product_name,
                         release_info.update_version,
                         item_number,
                         item_summary,
                         )
    return email_address, email_body


def comment_create(driver, email_address, comment_text):
    comments = element_wait(driver, 'x:324881036.4:mkr:ti3', By.ID)
    comments.click()
    new_comment = element_wait(driver, 'ctl00_ContentPlaceHolder1_tabMain_btnAddCommnt', By.ID)
    new_comment.click()
    comment_body_html = element_wait(driver, 'HTML', By.LINK_TEXT)
    comment_body_html.click()
    frame_wait(driver, 1)
    comment_body = element_wait(driver, '//textarea[1]', By.XPATH)
    comment_body.send_keys(comment_text)
    driver.switch_to.default_content()
    enter_email_address = element_wait(driver, 'ctl00_ContentPlaceHolder1_tbFreeInputEmail', By.ID)
    enter_email_address.send_keys(email_address)
    add_email_address = element_wait(driver, 'ctl00_ContentPlaceHolder1_imgAddFreeInputEmail', By.ID)
    add_email_address.click()
    save_comment = element_wait(driver, 'ctl00_ContentPlaceHolder1_imgSave', By.ID)
    save_comment.click()


def client_comment_create(ticket_id, comment_text):
    driver = open_browser_connect_to_site(
        settings.client_url, settings.client_title_assert)
    client_login(driver)
    open_client_ticket(driver, ticket_id.get_attribute('value'))
    comment_frame = driver.find_element_by_xpath('//*[@title="Rich Text Editor, comment"]')
    driver.switch_to.frame(comment_frame)
    comment_body = driver.switch_to_active_element()
    comment_body.send_keys(comment_text)
    driver.find_element_by_id('post_comment').click()
    driver.switch_to.default_content()
    posted_comment = driver.find_element_by_class_name('analysis_comment')
    posted_comment.find_element_by_class_name('comment_view').click()
    return posted_comment.text


def comment_save_to_linked_request(driver, bug_or_change_id, problem_summary, ticket_id, update_leapfile):
    email, comment = generate_product_update_comment_linked_request(
        driver, bug_or_change_id, problem_summary, ticket_id, update_leapfile)
    if ticket_id:
        copy_comment = client_comment_create(ticket_id, comment)
        comment_create(driver, '',
                       'Comment in Ticket %s:\n\n%s' % (ticket_id, copy_comment))
        print 'Comment update notification posted to Request copied from Ticket: %s' % ticket_id
    else:
        comment_create(driver, email, comment)
        print 'Comment update notification posted to Request sent to: %s' % email


def comment_save_to_request_on_notification(driver):
    email, comment = generate_product_update_comment_request_on_notification(driver)
    comment_create(driver, email, comment)
    print 'Comment update notification sent to: %s' % email


def set_request_solution(driver, update_version):
    solution_tab = element_wait(driver, 'x:324881036.2:mkr:ti1', By.ID)
    solution_tab.click()
    frame_wait(driver, 'ctl00_ContentPlaceHolder1_tabMain_htmlsys_field31_contentIframe')
    solution_desc = driver.switch_to_active_element()
    solution_desc.send_keys(Keys.CONTROL, 'a')
    solution_desc.send_keys('%s Upgrade' % update_version)
    print 'Solution updated'


import time


def select_dropdown_item(driver, button_id, item_id):
    dropdown_button = element_wait(driver, button_id, By.ID)
    dropdown_button.click()
    time.sleep(1)
    item = element_wait(driver, item_id, By.ID)
    item.click()


def set_request_pending_status(driver, status):
    driver.switch_to.default_content()
    element_wait(driver, 'x:654027361.4:mkr:ButtonImage', By.ID).click()
    element_wait(driver, status, By.LINK_TEXT).click()
    print 'Pending status updated:', status


def set_request_status(driver, status):
    try:
        change_request_status = element_wait(driver, 'ctl00_ContentPlaceHolder1_imgbtnsys_button2', By.ID)
        change_request_status.click()
        dropdown = element_wait(driver, 'x:738609024.4:mkr:ButtonImage', By.ID)
        dropdown.click()
        select_closed = element_wait(driver, status, By.LINK_TEXT, 2)
        select_closed.click()
        save = element_wait(driver, 'ctl00_ContentPlaceHolder1_imgSaveStatus', By.ID)
        save.click()
        try:
            confirm_suspend = element_wait(
                driver, 'ctl00_ContentPlaceHolder1_dialogReqSuspension_tmpl_imgChgStatus', By.ID)
            confirm_suspend.click()
        except:
            pass
        print 'Request status updated:', status
    except:
        print 'Request status already set:', status


def set_product_version_unknown_if_empty(driver):
    dropdown_container = element_wait(driver, 'x:654027363.7:mkr:List:nw:1', By.ID)
    selected_item = dropdown_container.find_element_by_class_name('igdd_%sListItemSelected' % settings.dd_prefix)
    selected_item_text = selected_item.find_element_by_tag_name('a').get_attribute('innerHTML')
    if not selected_item_text:
        select_dropdown_item(driver, 'x:654027363.4:mkr:ButtonImage', 'x:654027363.86:adr:78')


def update_linked_requests(driver, pslr, update_leapfile):
    for bug_or_change_id in pslr.keys():
        for rid in pslr[bug_or_change_id].requests:
            go_to_request(driver, rid)
            set_product_version_unknown_if_empty(driver)
            ticket_id_element = driver.find_element_by_id('ctl00_ContentPlaceHolder1_textfield6')
            ticket_id = ticket_id_element.get_attribute('value')
            problem_summary = pslr[bug_or_change_id].problem_summary
            comment_save_to_linked_request(driver,
                                           bug_or_change_id,
                                           problem_summary,
                                           ticket_id,
                                           update_leapfile,
                                           )
            if update_leapfile:
                email_address, email_body = generate_leapfile_email_body(driver,
                                                                         bug_or_change_id,
                                                                         problem_summary)
                send_leapfile(email_address, email_body)
            set_request_solution(driver, release_info.update_version)
            set_request_pending_status(driver, 'Completed')
            set_request_status(driver, 'Closed')


def update_requests_on_notification(driver, rids_on_notification):
    for rid in rids_on_notification:
        if go_to_request(driver, rid):
            comment_save_to_request_on_notification(driver)
            set_request_pending_status(driver, 'Waiting for Customer')
            set_request_status(driver, 'Hold')


def prompt_for_release_info(release_info):
    print 'Starting Product Release Auto-notification..'
    print 'This application will automatically generate request comments in %s ' \
          'and trigger emails to customers.' % settings.app_title_assert
    on_notification = []
    while True:
        while True:
            print 'Select from the following products:'
            for p in settings.products_acronyms:
                print '(%s)%s' % (p[0].upper(), p[1:])
            entry = raw_input('Select release product: ').lower()
            for p in settings.products_acronyms:
                if entry in (p.lower(), p[0].lower()):
                    product = p
                    break
            if product:
                break
            else:
                print 'Invalid entry.'
        while True:
            update_version = raw_input('Release version number? (e.g. 6.4, 6.4.6.1, 6.5.2): ')
            if update_version:
                break
        while True:
            bug_or_change = raw_input('Notify users for (b)ugs or (c)hanges?  ').lower()
            if bug_or_change in ('b', 'bug', 'bugs'):
                bug_or_change = 'bug'
                break
            elif bug_or_change in ('c', 'change', 'changes'):
                bug_or_change = 'change'
                break
            else:
                bug_or_change = ''
                print 'Invalid entry.'
        while True:
            cell_range = raw_input('Enter cell range of the %s numbers from the Google '
                                   'test sheet titled: "%s %s Test Sheet." '
                                   'Non-integer cells will be omitted. '
                                   '(e.g. B8:B19, B21:B26): ' % (bug_or_change,
                                                                 settings.products_acronyms[product],
                                                                 update_version))
            if cell_range:
                break
        while True:
            rid = raw_input('Enter a Request ID On Notification for the update. Press ENTER when finished. ')
            if not rid:
                break
            else:
                try:
                    on_notification.append(int(rid))
                except:
                    print 'Invalid entry.'
        while True:
            print '\n\nRELEASE SUMMARY:\n'
            print 'Product: %s' % product
            print 'Version: %s' % update_version
            print 'Bugs/Changes: %ss' % bug_or_change
            print 'Cell Range: %s\n' % cell_range
            print 'On notification: %s\n' % on_notification
            confirm = raw_input('Confirm? (Y)es or (N)o: ').lower()
            if confirm in ('y', 'yes'):
                return release_info(product,
                                    update_version,
                                    bug_or_change,
                                    cell_range,
                                    on_notification)
            elif confirm in ('n', 'no'):
                break
            else:
                print 'Invalid entry.'


def product_update_processor(driver, args):
    release_data = namedtuple('Release_Info', ['product_name',
                                               'update_version',
                                               'bug_or_change',
                                               'cell_range',
                                               'on_notification'])
    global release_info
    if args.test_data:
        release_info = release_data(settings.test_release_info['product_name'],
                                    settings.test_release_info['update_version'],
                                    settings.test_release_info['bug_or_change'],
                                    settings.test_release_info['cell_range'],
                                    settings.test_release_info['on_notification'])
    else:
        release_info = prompt_for_release_info(release_data)
    item_list = get_item_number_list(release_info.cell_range,
                                     settings.products_acronyms[release_info.product_name],
                                     release_info.update_version)
    pslr = get_psummary_linked_requests(item_list, release_info.bug_or_change)
    update_linked_requests(driver, pslr, args.update_leapfile)
    update_requests_on_notification(driver, release_info.on_notification)


import urllib2
import os


def save_attachments(fieldset):
    attach_container = fieldset.find_element_by_class_name('main_attachments')
    attachments = attach_container.find_elements_by_tag_name('a')
    file_list = []
    for a in attachments:
        response = urllib2.urlopen(a.get_attribute('href'))
        f = open(os.path.join(settings.attachments_path, a.text), 'w')
        f.write(response.read())
        f.close()
        file_list.append(a.text)
    return file_list


def build_ticket_dict(driver, ticket_id):
    def get_element_by_name(fieldset, name):
        return fieldset.find_element_by_xpath('//input[@name="%s"]' % name).get_attribute('value')
    def get_element_by_class(fieldset, class_name):
        return fieldset.find_element_by_class_name('%s' % class_name).text
    ticket_details = element_wait(driver, 'ticket_details', By.CLASS_NAME)
    ticket_details.click()
    fieldset = element_wait(driver, 'usualValidate', By.ID).find_element_by_tag_name('fieldset')
    return {'ticket_id': ticket_id,
            'customer_id': fieldset.find_element_by_xpath('//div[@class="formRight"]/span').text,
            'contact_name': get_element_by_name(fieldset, 'contact'),
            'email': get_element_by_name(fieldset, 'email'),
            'phone': get_element_by_name(fieldset, 'phone'),
            'region': get_element_by_class(fieldset, 'ticket_region_view'),
            'product': get_element_by_class(fieldset, 'ticket_products'),
            'product_version': get_element_by_name(fieldset, 'product_version'),
            'oper_sys': get_element_by_class(fieldset, 'select_os_ticket'),
            'sql_version': get_element_by_class(fieldset, 'select_sql_ticket'),
            'mail_server': get_element_by_name(fieldset, 'mail_s'),
            'problem_summary': get_element_by_name(fieldset, 'summary'),
            'problem_description': fieldset.find_element_by_xpath(
                '//textarea[@name="description"]').get_attribute('innerHTML').replace(
                    '&lt;p&gt;', '').replace('&lt;/p&gt;', '').replace('&nbsp;', ''),
            'attachments': save_attachments(fieldset)
            }


def get_client_ticket_details(driver, ticket_id):
    open_client_ticket(driver, ticket_id)
    ticket_details = build_ticket_dict(driver, ticket_id)
    return ticket_details


def create_new_request(driver, base_url=settings.app_url):
    driver.get('%s/ReqInfo.aspx?reqclass=(Default)' % base_url)


def click_save(driver):
    save = element_wait(driver, '//input[contains(@title,"Save")]', By.XPATH)
    save.click()


def get_version_ints(product_version):
    version_ints = ''
    for x in product_version:
        try:
            version_ints += str(int(x))
        except:
            pass
    return version_ints


def select_product_version(driver, product_version):
    dropdown_container = driver.find_element_by_id('x:654027363.7:mkr:List:nw:1')
    product_version_field = element_wait(driver, 'x:654027363.2:mkr:Input', By.ID)
    product_version_field.send_keys(Keys.ARROW_DOWN)
    while True:
        product_version_field.send_keys(Keys.ARROW_DOWN)
        selected_item = dropdown_container.find_element_by_class_name('igdd_%sListItemSelected' % settings.dd_prefix)
        selected_item_text = selected_item.find_element_by_tag_name('a').get_attribute('innerHTML')
        version_number = get_version_ints(selected_item_text)
        if version_number == product_version or selected_item_text == 'unknown':
            break


def enter_client_ticket_data_into_request(driver, ticket_data):

    def get_popup_window(driver, open_windows_before_popup):
        for w in driver.window_handles:
            if w not in open_windows_before_popup:
                return w

    def open_popup(driver, selector_id):
        existing_windows = driver.window_handles
        request_type = element_wait(driver, selector_id, By.ID)
        request_type.click()
        time.sleep(1)
        pop_up = get_popup_window(driver, existing_windows)
        driver.switch_to.window(pop_up)

    def select_request_type(driver, product):
        open_popup(driver, 'ctl00_ContentPlaceHolder1_hlsys_field33')
        product_list = []
        for p in settings.products_acronyms:
            product_list.append(p)
        product_list.sort()
        if product_list[0] in product:
            selection = element_wait(driver, 'x:2135213565.352:mkr:dtnContent', By.ID)
            selection.click()
        elif product_list[1] in product:
            selection = element_wait(driver, 'x:2135213565.724:mkr:dtnContent', By.ID)
            selection.click()
        else:
            print 'Product not converted to request type:', product
            clear = element_wait(driver, 'ctl00_ContentPlaceHolder1_hlClear', By.ID)
            clear.click()
        driver.switch_to.window(driver.window_handles[0])

    def select_region(driver, region):
        open_popup(driver, 'ctl00_ContentPlaceHolder1_hlsys_field43selsite')
        app_regions = driver.find_elements_by_tag_name('a')
        uk_regions = ('Europe', 'Middle East')
        for r in app_regions:
            if r.text == region:
                r.click()
                break
            elif r.text == 'UK' and region in uk_regions:
                r.click()
                break
        driver.switch_to.window(driver.window_handles[0])

    def upload_attachments(driver):
        file_list = os.listdir(settings.attachments_path)
        if file_list:
            element_wait(driver, 'x:324881036.5:mkr:ti4', By.ID).click()
            element_wait(
                driver, 'ctl00_ContentPlaceHolder1_tabMain_btnAttach', By.ID).click()
            for file_name in file_list:
                file_path = os.path.join(settings.attachments_path, file_name)
                element_wait(
                    driver, 'ctl00_ContentPlaceHolder1_h2Add', By.ID).click()
                element_wait(driver,
                             'ctl00_ContentPlaceHolder1_dialogInfo_tmpl_fuMain',
                             By.ID).send_keys(file_path)
                driver.find_element_by_id(
                    'ctl00_ContentPlaceHolder1_dialogInfo_tmpl_btnUpload').click()
            element_wait(driver, '//img[@title="Back"]', By.XPATH).click()

    element_wait(driver, 'ctl00_ContentPlaceHolder1_textsys_field6', By.ID).send_keys(ticket_data['customer_id'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield2', By.ID).send_keys(ticket_data['contact_name'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield3', By.ID).send_keys(ticket_data['email'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield4', By.ID).send_keys(ticket_data['phone'])
    select_region(driver, ticket_data['region'])
    request_source_dropdown_id, select_web_submission_id = 'x:654027360.3:mkr:Button', 'x:654027360.13:adr:5'
    select_dropdown_item(driver, request_source_dropdown_id, select_web_submission_id)
    caller_status_dropdown_id, select_customer_id = 'x:654027362.3:mkr:Button', 'x:654027362.10:adr:2'
    select_dropdown_item(driver, caller_status_dropdown_id, select_customer_id)
    select_product_version(driver, get_version_ints(ticket_data['product_version']))
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield6', By.ID).send_keys(ticket_data['ticket_id'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textsys_field35', By.ID).send_keys(ticket_data['problem_summary'])
    select_request_type(driver, ticket_data['product'])
    environment_details = "OS: %s\nSQL Server: %s\nMail Server: %s" % (ticket_data['oper_sys'],
                                                                       ticket_data['sql_version'],
                                                                       ticket_data['mail_server'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield5', By.ID).send_keys(environment_details)
    upload_attachments(driver)
    problem_desc_tab = element_wait(driver, 'x:324881036.1:mkr:ti0', By.ID)
    problem_desc_tab.click()
    frame_wait(driver, 'ctl00_ContentPlaceHolder1_tabMain_htmlsys_field36_contentIframe')
    html_field = driver.switch_to_active_element()
    html_field.send_keys(ticket_data['problem_description'])
    driver.switch_to.default_content()
    click_save(driver)


def copy_ticket_from_client_to_app(driver, ticket_id):

    def confirm_request_saved(problem_summary):
        sql = ('select sys_request_id, usr_ticket_id, sys_problemsummary '
               'from request where usr_ticket_id = %s and '
               'sys_problemsummary = "%s"') % (ticket_id, problem_summary)
        cursor = db_connect().cursor()
        cursor.execute(sql)
        row = cursor.fetchone()
        cursor.close()
        if not row:
            raise Exception('Request not found. Copy failed.')
        else:
            print 'Ticket successfully copied to Request:', row[0]

    client_login(driver)
    ticket_details = get_client_ticket_details(driver, ticket_id)
    driver.get(settings.app_url)
    login_analyst(driver)
    create_new_request(driver)
    enter_client_ticket_data_into_request(driver, ticket_details)
    confirm_request_saved(ticket_details['problem_summary'])


def send_leapfile(email_address, email_body):

    def download_file_from_fileserver():
        response = urllib2.urlopen(os.path.join(settings.fileserver_url,
                                                product_acronym.lower(),
                                                release_info.update_version,
                                                file_name))
        local_file_path = os.path.join(settings.attachments_path, file_name)
        f = open(local_file_path, 'w')
        f.write(response.read())
        print 'File downloaded successfully:', file_name
        f.close()

    product_acronym = settings.products_acronyms[release_info.product_name]
    version = release_info.update_version.replace('.', '')
    file_name = ''.join((product_acronym.upper(), version, 'p.zip'))
    if file_name not in os.listdir(settings.attachments_path):
        download_file_from_fileserver()
    driver = open_browser_connect_to_site(settings.leapfile_url, settings.leapfile_title_assert)
    driver.find_element_by_id('employeeLoginButton').click()
    username = element_wait(driver, 'userID', By.ID)
    username.send_keys(settings.leapfile_username)
    driver.find_element_by_id('password').send_keys(settings.leapfile_password)
    driver.find_element_by_xpath('//input[@name="logon"]').click()
    driver.find_element_by_class_name('transfer-button').click()
    driver.find_element_by_id('contactList').send_keys(email_address)
    message_subject = '%s %s' % (release_info.product_name, release_info.update_version)
    driver.find_element_by_id('s').send_keys(message_subject)
    driver.switch_to.frame('m___Frame')
    driver.find_element_by_id('xEditingArea').send_keys(email_body)
    driver.switch_to.default_content()
    driver.find_element_by_xpath('//input[@name="basic"]').click()
    driver.find_element_by_xpath('//input[@name="file0"]').send_keys(
        os.path.join(settings.attachments_path, file_name))
    driver.find_element_by_id('UploadAndSend').click()
    driver.quit()


import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-u', '--update_notifier', help='Run update notifier.', action='store_true')
    parser.add_argument('-l', '--update_leapfile', help='Send update with Leapfile', action='store_true')
    parser.add_argument('-t', '--test_data', help='Run update notifier with test data.', action='store_true')
    parser.add_argument('-c', '--copy_ticket', help='Copy ticket to request')
    args = parser.parse_args()
    if args.update_notifier:
        driver = open_browser_connect_to_site(settings.app_url, settings.app_title_assert)
        login_analyst(driver)
        product_update_processor(driver, args.test_data)
    elif args.update_leapfile:
        driver = open_browser_connect_to_site(settings.app_url, settings.app_title_assert)
        login_analyst(driver)
        product_update_processor(driver, args)
    elif args.copy_ticket:
        driver = open_browser_connect_to_site(settings.client_url, settings.client_title_assert)
        copy_ticket_from_client_to_app(driver, args.copy_ticket)
    for f in os.listdir(settings.attachments_path):
        os.remove(os.path.join(settings.attachments_path, f))
    driver.quit()


if __name__ == '__main__':
    main()