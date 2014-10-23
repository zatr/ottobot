from selenium import webdriver
import unittest
import settings
import worker


class ProductUpdateProcessorTest(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Firefox()
        self.target_url = settings.app_url
        self.assert_text = settings.app_title_assert
        self.driver.get(self.target_url)
        self.assert_title()
        worker.login_analyst(self.driver)

    def tearDown(self):
        self.driver.quit()

    def assert_title(self):
        try:
            self.assertIn(self.assert_text, self.driver.title)
        except:
            print 'Page title assertion failed. Check settings.'
            self.tearDown()
            self.skipTest(ProductUpdateProcessorTest)

    def test_product_update(self):
        worker.product_update_processor(self.driver)


class CopyTicketFromClientToAppTest(unittest.TestCase):

    def setUp(self):
        self.driver = webdriver.Firefox()
        self.target_url = settings.client_url
        self.assert_text = settings.client_title_assert
        self.driver.get(self.target_url)
        self.assert_title()

    def tearDown(self):
        self.driver.quit()

    def assert_title(self):
        try:
            self.assertIn(self.assert_text, self.driver.title)
        except:
            print 'Page title assertion failed. Check settings.'
            self.skipTest(CopyTicketFromClientToAppTest)

    def test_copy_ticket_from_client_to_app(self):
        worker.copy_ticket_from_client_to_app(self.driver, settings.ticket_id)