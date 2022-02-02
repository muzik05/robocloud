import os
from re import search
from time import sleep
from glob import glob

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF


class Worker:
    def __init__(self, dir='output', excel_filename='rpa_test.xlsx', selected_agencie=False):
        self.browser = Selenium()
        self.pdf = PDF()
        self.output_dir = os.path.abspath(dir)
        self.url = 'https://itdashboard.gov/'
        self.excel = Files().create_workbook(f'{dir}/{excel_filename}')
        self.selected_agencie = selected_agencie
        if not os.path.isdir(self.output_dir):
            os.mkdir(self.output_dir)

    def run(self):
        # open browser and go to url
        self.open_browser_url(self.url)
        # click 'DIVE IN'
        self.click_element('css:a[aria-controls=home-dive-in]')
        # wait until page loaded
        self.wait_until_page_contains_element(
            '//*[@id="agency-tiles-widget"]//a[./span]')
        # get agencies data
        agencies_data = self.get_elements(
            '//*[@id="agency-tiles-widget"]//a[./span]')
        # parse agencies data
        agencies = self.parse_agencies_data(agencies_data)
        # rename worksheet
        self.excel.rename_worksheet('Agencies')
        # write data to excel
        self.write_to_excel(agencies, keys=['name', 'amount'])
        # go to selected agencie
        self.go_to_url_and_wait(agencies[self.selected_agencie]['link'],
                                 "return !document.querySelector('#investments-table-container').classList.contains('loading')")
        # select from list - All
        self.select_from_list_and_wait('name:investments-table-object_length', 'All',
                                       "return !document.querySelector('#investments-table-container').classList.contains('loading')")
        # get selected agencie investments into dict
        investments_table = self.get_dict_from_investments_current_agencie()
        # create worksheet - selected_agencie
        self.excel.create_worksheet(selected_agencie)
        # write data to excel
        self.write_to_excel(investments_table, keys=['uii', 'bureau', 'investment_title',
                                                     'total_spending', 'type', 'cio_rating', 'projects_count'])
        self.excel.save()
        # download all pdf files
        self.download_all_pdfs(investments_table)
        # compare pdf files with table
        self.compare_pdf_with_table(investments_table)

    def compare_pdf_with_table(self, investments_table):
        res_file = open(f'{self.output_dir}/compare_result.csv', 'w+')

        for line in investments_table.values():
            if 'pdf_link' not in line:
                continue

            pdf_filename = line['pdf_link'].split('/')[-1]+'.pdf'
            text = self.pdf.get_text_from_pdf(
                f'{self.output_dir}/{pdf_filename}', trim=False)

            investment_title = search(
                'Name of this Investment:\s+([^\n]+)', text[1])
            investment_title = self.get_found_element_or_false(investment_title)

            uii = search(
                'Unique Investment Identifier \(UII\):\s+([^\n]+)', text[1])
            uii = self.get_found_element_or_false(uii)

            if investment_title == line['investment_title'] and uii == line['uii']:
                equal = True
            else:
                equal = False

            res_file.write(
                f'{pdf_filename},{investment_title},{uii},{line["investment_title"]},{line["uii"]},{equal}\n\n')
            
            self.pdf.close_all_pdfs()
  
    def get_found_element_or_false(self, res):
        return False if len(res.groups()) != 1 else res[1]

    def download_all_pdfs(self, investments_table):
        for line in investments_table.values():
            if 'pdf_link' not in line:
                continue

            self.go_to_url_and_wait(
                line['pdf_link'], "return !document.querySelector('#investment-quick-stats-container').classList.contains('loading')")
            
            count = len(glob(f'{self.output_dir}/*'))
            self.click_element('//*[@id="business-case-pdf"]/a')
            while len(glob(f'{self.output_dir}/*')) == count:
                sleep(1)
        
        while glob(f'{self.output_dir}/*.crdownload'):
            sleep(1)

    def get_dict_from_investments_current_agencie(self):
        names = ['uii', 'bureau', 'investment_title',
                 'total_spending', 'type', 'cio_rating', 'projects_count']
        names_count = len(names)
        table = {}

        tds = self.get_elements('//table[@id="investments-table-object"]/tbody//td')

        for i, td in enumerate(tds):
            if int(i/names_count) in table:
                table[int(i/names_count)].update({names[i % names_count]: td.text})
            else:
                table[int(i/names_count)] = {names[i % names_count]: td.text}

            if not i % names_count:
                try:
                    a = td.find_element_by_xpath('./a')
                    table[int(i/names_count)]['pdf_link'] = a.get_attribute('href')
                except:
                    pass
        
        return table

    def select_from_list_and_wait(self, list_patern, selected, wait_patern=False):
        self.browser.select_from_list_by_label(list_patern, selected)
        self.wait_for_condition(wait_patern)

    def go_to_url_and_wait(self, url, wait_patern=False):
        self.browser.go_to(url)
        self.wait_for_condition(wait_patern)

    def wait_for_condition(self, wait_patern=False, timeout=20):
        if wait_patern:
            self.browser.wait_for_condition(wait_patern, timeout)

    def parse_agencies_data(self, agencies_data):
        agencies = {}

        for a in agencies_data:
            text = a.text.split('\n')
            agencies[text[0]] = dict(
                name=text[0], amount=text[2], link=a.get_attribute('href'))
        
        return agencies

    def wait_until_page_contains_element(self, element, timeout=20):
        self.browser.wait_until_page_contains_element(element, timeout)

    def click_element(self, element):
        self.browser.click_element(element)

    def get_elements(self, patern):
        return self.browser.get_webelements(patern)

    def write_to_excel(self, data, keys=[]):
        for i, line in enumerate(data.values()):
            for x, (key, value) in enumerate(line.items()):
                if keys and key in keys:
                    self.excel.set_cell_value(i+1, keys.index(key)+1, value)
                elif not keys:
                    self.excel.set_cell_value(i+1, x+1, value)

    def open_browser_url(self, url):
        self.browser.open_available_browser(url,
                          headless=True,
                          preferences={'download.default_directory': self.output_dir})


if __name__ == '__main__':
    selected_agencie = open('select.txt').read()

    Worker(selected_agencie=selected_agencie).run()
