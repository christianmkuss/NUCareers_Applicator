from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common import exceptions
import xlsxwriter
from textblob import TextBlob
import matplotlib.pyplot as plt
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
import yaml
from io import StringIO

URL = "https://nucareers.northeastern.edu/students/student-login.htm"


# Returns a dictionary of the words that have been applied to with the value being the number of occurrences
def get_my_jobs():
    # Login to NUCareers
    browser = webdriver.Chrome("C:\\Program Files (x86)\\chromedriver_win32\\chromedriver.exe")
    browser.maximize_window()
    browser.get(URL)
    WebDriverWait(browser, 5).until(
        ec.presence_of_element_located((By.CLASS_NAME, "btn-primary"))
    )
    browser.find_element_by_css_selector('.btn.btn-primary.btn-large').click()

    # Read and use username and password from login.yaml
    with open("login.yaml", 'r') as login:
        try:
            login_info = yaml.safe_load(login)
            username = login_info['username']
            password = login_info['password']
            browser.find_element_by_id('username').send_keys(username)
            browser.find_element_by_id('password').send_keys(password)
            browser.find_element_by_css_selector('.form-button').click()
        except yaml.YAMLError as exc:
            print(exec)
            exit(1)

    browser.find_element_by_xpath("//a[@href='/myAccount/co-op/jobs.htm']").click()
    # The wait is required so the page can run the javascript
    WebDriverWait(browser, 5).until(
        ec.presence_of_element_located((By.LINK_TEXT, "Applied To"))
    )
    browser.find_element_by_xpath("//a[contains(text(),'Applied To')]").click()
    WebDriverWait(browser, 20).until(
        ec.presence_of_element_located((By.LINK_TEXT, "view"))
    )

    html = browser.page_source
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    companies = []
    skills = ['python', 'c++', 'experience', 'java', 'c#', 'unity']

    # TODO: Have this be read from the yaml
    term = "2020 - Spring"

    # Gathers the company names
    for row in soup.findAll('table')[0].tbody.findAll('tr'):
        if row.findAll('td')[1].text == term:
            job = [row.findAll('td')[4].text, row.findAll('td')[3].text, row.findAll('td')[10].text]
            companies.append([row.findAll('td')[4].text.split()])
            jobs.append(job)

    descriptions = description(browser, len(companies))
    browser.close()
    my_profile = create_common_words(descriptions, companies, 3)

    for item in skills:
        if item in my_profile:
            my_profile[item] += 1
        else:
            my_profile[item] = 1

    return my_profile


def apply(profile_list, location_list):
    # Opens a browser and navigates to the right page
    browser = webdriver.Chrome("C:\\Program Files (x86)\\chromedriver_win32\\chromedriver.exe")
    browser.maximize_window()
    browser.get(URL)
    WebDriverWait(browser, 5).until(
        ec.presence_of_element_located((By.CLASS_NAME, "btn-primary"))
    )
    browser.find_element_by_css_selector('.btn.btn-primary.btn-large').click()

    browser.find_element_by_id('username').send_keys('*****')
    browser.find_element_by_id('password').send_keys('******')
    browser.find_element_by_css_selector('.btn-submit').click()

    browser.find_element_by_xpath("//a[@href='/myAccount/co-op/jobs.htm']").click()
    WebDriverWait(browser, 5).until(
        ec.presence_of_element_located((By.LINK_TEXT, "For My Program"))
    )
    browser.find_element_by_xpath("//a[contains(text(),'For My Program')]").click()
    WebDriverWait(browser, 20).until(
        ec.presence_of_element_located((By.LINK_TEXT, "View"))
    )

    jobs_to_apply = []
    company_names = []
    all_description = []
    for page_num in range(9):
        browser.refresh()
        WebDriverWait(browser, 20).until(
            ec.presence_of_element_located((By.LINK_TEXT, "View"))
        )
        for job in browser.find_elements_by_link_text("View"):
            counter = 0.0
            this_description = []
            location = ''
            job.click()
            browser.switch_to.window(browser.window_handles[1])
            html = browser.page_source
            soup = BeautifulSoup(html, "html.parser")
            try:
                location = soup.findAll('table')[2].tbody.findAll('tr')[8].find('td', {"width": "75%"}).text.strip()
                for row in soup.findAll('table')[2].tbody.findAll('tr')[12].findAll('td', {"width": "75%"}):
                    descript = row.text
                    descript = descript.replace("\n", "")
                    descript = descript.replace("\t", "")
                    descript = descript.replace("\xa0", " ")
                    descript = TextBlob(descript)
                    for element in descript.noun_phrases:
                        this_description.append(element)
                        all_description.append(element)
                    company_names.append(browser.find_element_by_class_name("span6").text)
            except IndexError:
                print("Could not gather data from ", browser.find_element_by_class_name("span6").text)
                pass

            for noun in this_description:
                if noun in profile_list:
                    counter += 1.0

            if location in location_list:
                job_name = browser.find_element_by_class_name("span6").text
                jobs_to_apply.append(job_name)
                try:
                    browser.find_element_by_class_name('applyButton').click()
                    browser.find_element_by_xpath("//input[@value='existingPkg']").click()
                    package = browser.find_element_by_xpath("//input[@name='pac']")
                    browser.execute_script("return arguments[0].scrollIntoView(true);", package)
                    package.click()
                    submit = browser.find_element_by_xpath("//input[@value='Submit Application']")
                    submit.click()
                except exceptions.NoSuchElementException:
                    print('Could not apply to ', job_name)
            browser.close()
            browser.switch_to.window(browser.window_handles[0])

        browser.switch_to.window(browser.window_handles[0])
        position = browser.find_element_by_xpath("//button[contains(text(), 'Apply Filters')]")
        browser.execute_script("return arguments[0].scrollIntoView(true);", position)
        browser.find_element_by_xpath("//a[contains(text(), '»')]").click()

    all_words = create_common_words(all_description, company_names, 50)

    for job in jobs_to_apply:
        print(job)

    return all_words


def write_all_job(applications):
    book = xlsxwriter.Workbook('MyApplications.xlsx')
    sh = book.add_worksheet()

    row = 0
    col = 0

    for word in applications:
        sh.write(row, col, word)
        sh.write(row, col + 1, applications[word])
        row += 1
    book.close()


def write_file(applications):
    book = xlsxwriter.Workbook('MyApplications.xlsx')
    sh = book.add_worksheet()

    row = 0
    col = 0

    for company, job, date in applications:
        sh.write(row, col, company)
        sh.write(row, col + 1, job)
        sh.write(row, col + 2, date)
        row += 1
    book.close()


def description(browser, num_jobs):
    job_descriptions = []
    for job in browser.find_elements_by_link_text("view"):
        if num_jobs == 0:
            break
        job.click()
        browser.find_element_by_link_text("new tab").click()
        browser.switch_to.window(browser.window_handles[1])
        html = browser.page_source
        soup = BeautifulSoup(html, "html.parser")
        try:
            job_description = soup.findAll('table')[2].tbody.findAll('tr')[11]
            if job_description.td.text.strip() == 'Job Description:':
                this_description = job_description.find('td', {'width': '75%'}).text
                this_description = this_description.replace("\n", "")
                this_description = this_description.replace("\t", "")
                this_description = this_description.replace("\xa0", " ")
                this_description = TextBlob(this_description)
                job_descriptions.append(this_description.noun_phrases)
        except IndexError:
            pass
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
        num_jobs -= 1
    return job_descriptions


def create_common_words(nouns, names, cutoff):
    # Words takes all the words
    words = []
    # Final takes the words and counts the usages
    final = {}

    bad_words = ['sexual orientation', 'gpa', 'veteran status', 'co-op', 'product', 'will', 'who', 'january',
                 'wide range']

    for job in nouns:
        if cutoff > 1:
            for word in job:
                words.append(word)
        else:
            words.append(job)

    for this_word in words:
        if words.count(this_word) > cutoff and this_word not in final and this_word not in names and this_word not in bad_words:
            final[this_word] = words.count(this_word)

    print(final)
    return final


def resume_scraper(resume):
    profile = []
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(resume, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    data = TextBlob(text).noun_phrases
    for noun in data:
        profile.append(noun)
    return create_common_words(profile, [], 0)


# Creates a pie chart of a dictionary
def create_plot(data):
    labels = []
    size = []
    colors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue']
    y = 0
    actual_colors = []
    for x in data:
        labels.append(x)
        size.append(data.get(x))
        if y == len(colors):
            y = 0
        actual_colors.append(colors[y])
        y += 1
    size, labels = (list(t) for t in zip(*sorted(zip(size, labels))))
    plt.pie(size, labels=labels, colors=colors, shadow=True)
    plt.axis('equal')
    plt.show()


if __name__ == '__main__':
    profile = {**resume_scraper('Kuss_Resume.pdf'), **get_my_jobs()}
    # jobs = apply(profile, locations)
    # create_plot(jobs)
