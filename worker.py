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
    print '\nOpened Request', rid


def go_to_item(driver, problem_or_change, item_id):
    driver.get('%s/%sInfo.aspx?sys_%s_id=%i' % (settings.app_url,
                                                problem_or_change,
                                                problem_or_change,
                                                item_id))


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


import pyodbc


def connect_to_db():
    try:
        return pyodbc.connect(''.join(settings.db_connect_string))
    except:
        print 'Database connection failed. Check settings.'


def sql_select_linked_requests(item_number, bug_or_change):
    if bug_or_change == 'bug':
        return ('select req_problem_change.sys_problem_id, request.sys_request_id, '
                'usr_Customer_Name, usr_Cust_Email, problem.sys_problem_summary '
                'from request join req_problem_change '
                'on request.sys_request_id = req_problem_change.sys_request_id '
                'join problem '
                'on req_problem_change.sys_problem_id = problem.sys_problem_id '
                'where req_problem_change.sys_problem_id = %i') % item_number
    if bug_or_change == 'change':
        return ('select req_problem_change.sys_change_id, request.sys_request_id, '
                'usr_Customer_Name, usr_Cust_Email, change.sys_change_summary '
                'from request join req_problem_change '
                'on request.sys_request_id = req_problem_change.sys_request_id '
                'join change '
                'on req_problem_change.sys_change_id = change.sys_change_id '
                'where req_problem_change.sys_change_id = %i') % item_number


from collections import defaultdict, namedtuple


def get_psummary_linked_requests(items, bug_or_change):
    linked_requests = defaultdict(list)
    psummary_linked_req = namedtuple('LinkedRequestList', ['problem_summary',
                                                           'requests'])
    all_linked_rids = []
    print '\nLoading requests linked to %ss:' % bug_or_change, items
    cursor = connect_to_db().cursor()
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
    connect_to_db().close()
    print '\nAll requests queued for notification:'
    for rid in sorted(all_linked_rids):
        print rid
    return linked_requests


def generate_comment(driver, item_number, item_summary):
    def collect_cust_fname_email():
        usr_customer_name = element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield2', By.ID)
        first_name = usr_customer_name.get_attribute('value').split()[0]
        usr_cust_email = element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield3', By.ID)
        email_address = usr_cust_email.get_attribute('value')
        return first_name, email_address
    def get_comment(boc):
        template = open(settings.templates_path + '/%s_notification.html' % boc)
        text = template.read()
        template.close()
        return text
    first_name, email_address = collect_cust_fname_email()
    comment_text = get_comment(release_info.bug_or_change) % (first_name,
                                                              release_info.product_name,
                                                              release_info.update_version,
                                                              item_number,
                                                              item_summary,
                                                              settings.client_url)
    return email_address, comment_text


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


def comment_save_to_linked_request(driver, bug_or_change_id, problem_summary):
    email, comment = generate_comment(driver, bug_or_change_id, problem_summary)
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


def select_dropdown_item(driver, button_id, item_id):
    dropdown_button = element_wait(driver, button_id, By.ID)
    dropdown_button.click()
    time.sleep(1)
    item = element_wait(driver, item_id, By.ID)
    item.click()


def set_request_pending_status_completed(driver):
    driver.switch_to.default_content()
    element_wait(driver, 'x:654027361.4:mkr:ButtonImage', By.ID).click()
    element_wait(driver, 'Completed', By.LINK_TEXT).click()
    print "Pending status set to 'Completed'"


def set_request_status_closed(driver):
    try:
        change_request_status = element_wait(driver, 'ctl00_ContentPlaceHolder1_imgbtnsys_button2', By.ID)
        change_request_status.click()
        dropdown = element_wait(driver, 'x:738609024.4:mkr:ButtonImage', By.ID)
        dropdown.click()
        select_closed = element_wait(driver, 'Closed', By.LINK_TEXT, 2)
        select_closed.click()
        save = element_wait(driver, 'ctl00_ContentPlaceHolder1_imgSaveStatus', By.ID)
        save.click()
        print 'Request closed'
    except:
        print 'Request already closed'


def set_product_version_unknown_if_empty(driver):
    dropdown_container = element_wait(driver, 'x:654027363.7:mkr:List:nw:1', By.ID)
    selected_item = dropdown_container.find_element_by_class_name('igdd_%sListItemSelected' % settings.dd_prefix)
    selected_item_text = selected_item.find_element_by_tag_name('a').get_attribute('innerHTML')
    if not selected_item_text:
        select_dropdown_item(driver, 'x:654027363.4:mkr:ButtonImage', 'x:654027363.86:adr:78')


def apply_changes_to_linked_requests(driver, pslr):
    for bug_or_change_id in pslr.keys():
        for rid in pslr[bug_or_change_id].requests:
            go_to_request(driver, rid)
            set_product_version_unknown_if_empty(driver)
            comment_save_to_linked_request(driver, bug_or_change_id,
                                           pslr[bug_or_change_id].problem_summary)
            set_request_solution(driver, release_info.update_version)
            set_request_pending_status_completed(driver)
            set_request_status_closed(driver)


def set_item_list_status_closed(driver, items, problem_or_change):
    for item in items:
        go_to_item(driver, problem_or_change, item)
        set_request_status_closed(driver)


def rename_bug_to_problem():
    if release_info.bug_or_change == 'bug':
        return 'problem'
    if release_info.bug_or_change == 'change':
        return 'change'
    else:
        print 'How did you even do this?'


def prompt_for_release_info(release_info):
    print 'Starting Product Release Auto-notification..'
    print 'This application will automatically generate request comments in %s ' \
          'and trigger emails to customers. Choose product, specify version, ' \
          'and select cells from the Google worksheet to build problem/change list. ' \
          % settings.app_title_assert
    product = update_version = bug_or_change = cell_range = ''
    while True:
        while not product:
            print 'Select from the following products:'
            for p in settings.products_acronyms:
                print '(%s)%s' % (p[0].upper, p[1:])
            entry = raw_input('Select release product:' % ()).lower()
            for p in settings.products_acronyms:
                if entry in (p.lower, p[0].lower):
                    product = p
            if not product:
                print 'Invalid entry.'
        while not update_version:
            update_version = raw_input('Release version number? (e.g. 6.4, 6.4.6.1, 6.5.2): ')
        while not bug_or_change:
            bug_or_change = raw_input('Notify users for (b)ugs or (c)hanges?  ').lower()
            if bug_or_change in ('b', 'bug', 'bugs'):
                bug_or_change = 'bug'
            elif bug_or_change in ('c', 'change', 'changes'):
                bug_or_change = 'change'
            else:
                bug_or_change = ''
                print 'Invalid entry.'
        while not cell_range:
            cell_range = raw_input('Enter cell range of the %s numbers from the Google test sheet titled: %s %s Test Sheet '
                                   'Non-integer cells will be omitted. (e.g. B8:B19, B21:B26): ' % (bug_or_change,
                                                                                                    product,
                                                                                                    update_version))
        print '\n\nRELEASE SUMMARY:\n'
        print 'Product: %s' % product
        print 'Version: %s' % update_version
        print 'Bugs/Changes: %ss' % bug_or_change
        print 'Cell Range: %s\n' % cell_range
        confirm = raw_input('Confirm? (Y)es or (N)o: ').lower()
        if confirm in ('y', 'yes'):
            return release_info(product, update_version, bug_or_change, cell_range)
        elif confirm in ('n', 'no'):
            return False
        else:
            print 'Invalid entry.'


def product_update_processor(driver, use_test_data=False):
    release_data = namedtuple('Release_Info', ['product_name',
                                               'update_version',
                                               'bug_or_change',
                                               'cell_range'])
    global release_info
    if use_test_data:
        release_info = release_data(settings.test_release_info['product_name'],
                                    settings.test_release_info['update_version'],
                                    settings.test_release_info['bug_or_change'],
                                    settings.test_release_info['cell_range'])
    else:
        release_info = prompt_for_release_info(release_data)
    item_list = get_item_number_list(release_info.cell_range,
                                     settings.products_acronyms[release_info.product_name],
                                     release_info.update_version)
    psummary_linked_requests = get_psummary_linked_requests(item_list, release_info.bug_or_change)
    apply_changes_to_linked_requests(driver, psummary_linked_requests)


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


import urllib2


def save_attachments(fieldset):
    attach_container = fieldset.find_element_by_class_name('main_attachments')
    attachments = attach_container.find_elements_by_tag_name('a')
    file_list = []
    for a in attachments:
        response = urllib2.urlopen(a.get_attribute('href'))
        f = open(settings.attachments_path + '/' + a.text, 'w')
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


import time


def enter_client_ticket_data_into_request(driver, ticket_data):
    def get_popup_window(driver, open_windows_before_popup):
        for w in driver.window_handles:
            if w not in open_windows_before_popup:
                return w
    def select_request_type(driver, product):
        existing_windows = driver.window_handles
        request_type = element_wait(driver, 'ctl00_ContentPlaceHolder1_hlsys_field33', By.ID)
        request_type.click()
        time.sleep(1)
        pop_up = get_popup_window(driver, existing_windows)
        driver.switch_to.window(pop_up)
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
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textsys_field6', By.ID).send_keys(ticket_data['customer_id'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield2', By.ID).send_keys(ticket_data['contact_name'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield3', By.ID).send_keys(ticket_data['email'])
    element_wait(driver, 'ctl00_ContentPlaceHolder1_textfield4', By.ID).send_keys(ticket_data['phone'])
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
    problem_desc_tab = element_wait(driver, 'x:324881036.1:mkr:ti0', By.ID)
    problem_desc_tab.click()
    frame_wait(driver, 'ctl00_ContentPlaceHolder1_tabMain_htmlsys_field36_contentIframe')
    html_field = driver.switch_to_active_element()
    html_field.send_keys(ticket_data['problem_description'])
    driver.switch_to.default_content()
    click_save(driver)


def copy_ticket_from_client_to_app(driver, ticket_id):
    client_login(driver)
    ticket_details = get_client_ticket_details(driver, ticket_id)
    driver.get(settings.app_url)
    login_analyst(driver)
    create_new_request(driver)
    enter_client_ticket_data_into_request(driver, ticket_details)