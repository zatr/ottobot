import os.path


basepath = os.path.dirname(__file__)
templates_path = os.path.abspath(os.path.join(basepath, 'templates'))
attachments_path = os.path.abspath(os.path.join(basepath, 'attachments'))

app_url = 'http://www.example.com'
app_title_assert = 'page title'
app_username = 'user'
app_password = 'pass'
db_connect_string = ['DRIVER={FreeTDS};',
                     'SERVER=SERVER_NAME\SQL_INSTANCE;',
                     'PORT=1433;',
                     'UID=username;',
                     'PWD=password;',
                     'DATABASE=database_name;'
                     ]

dd_prefix = 'prefix'

products_acronyms = {'Product1': 'P1',
                     'Product2': 'P2'}

test_release_info = {'product_name': 'Product1',
                     'update_version': '1.2.3',
                     'bug_or_change': 'bug',
                     'cell_range': 'A1:A10',
                     'on_notification': [12345, 67890]}

google_user = 'google_user@gmail.com'
google_pass = 'pass'
doc_name = '%s %s document'
worksheet_name = 'name'

client_url = 'http://www.example.com'
ticket_url = client_url + '/tickets?ticket='
client_title_assert = 'page title'
client_username = 'user'
client_password = 'pass'
ticket_id = '12345'

leapfile_url = 'http://site.leapfile.com'
leapfile_title_assert = 'Page title'
leapfile_username = 'username'
leapfile_password = 'password'

fileserver_url = 'http://fileserver'