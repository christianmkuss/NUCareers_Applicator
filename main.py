from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common import exceptions
import xlsxwriter
from textblob import TextBlob
import matplotlib.pyplot as plt
import PyPDF2
from TexSoup import TexSoup
import re

# Returns a dictionary of the words that have been applied to with the value the number of occurrences
def get_my_jobs():
    # Login to NUCareers
    url = "https://nucareers.northeastern.edu/students/student-login.htm"
    browser = webdriver.Chrome("C:\\Program Files (x86)\\chromedriver_win32\\chromedriver.exe")
    browser.maximize_window()
    browser.get(url)
    WebDriverWait(browser, 5).until(
        ec.presence_of_element_located((By.CLASS_NAME, "btn-primary"))
    )
    browser.find_element_by_css_selector('.btn.btn-primary.btn-large').click()

    browser.find_element_by_id('username').send_keys('****')
    browser.find_element_by_id('password').send_keys('******')
    browser.find_element_by_css_selector('.btn-submit').click()

    browser.find_element_by_xpath("//a[@href='/myAccount/co-op/jobs.htm']").click()
    WebDriverWait(browser, 5).until(
        ec.presence_of_element_located((By.LINK_TEXT, "Applied To"))
    )
    browser.find_element_by_xpath("//a[contains(text(),'Applied To')]").click()
    WebDriverWait(browser, 20).until(
        ec.presence_of_element_located((By.LINK_TEXT, "Cancel Application"))
    )

    html = browser.page_source
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    companies = []
    skills = ['python', 'c++', 'experience', 'java', 'c#', 'unity']

    # Gathers the company names
    for row in soup.findAll('table')[0].tbody.findAll('tr'):
        job = [row.findAll('td')[4].text, row.findAll('td')[3].text, row.findAll('td')[10].text]
        companies.append([row.findAll('td')[4].text.split()])
        jobs.append(job)

    descriptions = description(browser)
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
    url = "https://nucareers.northeastern.edu/students/student-login.htm"
    browser = webdriver.Chrome("C:\\Program Files (x86)\\chromedriver_win32\\chromedriver.exe")
    browser.maximize_window()
    browser.get(url)
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


def description(browser):
    this_description = []
    for job in browser.find_elements_by_link_text("view"):
        job.click()
        browser.find_element_by_link_text("new tab").click()
        browser.switch_to.window(browser.window_handles[1])
        html = browser.page_source
        soup = BeautifulSoup(html, "html.parser")
        try:
            for row in soup.findAll('table')[2].tbody.findAll('tr')[12].findAll('td', {"width": "75%"}):
                descript = row.text
                descript = descript.replace("\n", "")
                descript = descript.replace("\t", "")
                descript = descript.replace("\xa0", " ")
                descript = TextBlob(descript)
                this_description.append(descript.noun_phrases)
        except IndexError:
            pass
        browser.close()
        browser.switch_to.window(browser.window_handles[0])
    return this_description


def create_common_words(nouns, names, number):
    # Words takes all the words
    words = []
    # Final takes the words and counts the usages
    final = {}

    bad_words = ['sexual orientation', 'gpa', 'veteran status', 'co-op', 'product', 'will', 'who', 'january',
                 'wide range']

    for job in nouns:
        for word in job:
            words.append(word)

    for this_word in words:
        if words.count(this_word) > number and this_word not in final and this_word not in names and this_word not in bad_words:
            final[this_word] = words.count(this_word)

    return final


def latex_scraper():
    resume = open("Kuss Global Resume\\Kuss_GlobalResume.tex")
    soup = TexSoup(resume)
    soup = str(soup)
    text = re.search('{(.+?)]', soup)
    print(text)


def resume_scraper(resume):
    my_resume = []
    pdfFileObj = open('Kuss_Global_Resume.pdf', 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    pageObj = pdfReader.getPage(0)
    data = pageObj.extractText()
    pdfFileObj.close()
    data = TextBlob(data).noun_phrases
    print(data)
    for noun in data:
        my_resume.append(noun)
    return data


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
    # latex_scraper()
    # pdf_file = open('Resume Official Online.pdf', 'rb')
    # read_pdf = PyPDF2.PdfFileReader(pdf_file)
    # page = read_pdf.getPage(0)
    # page_content = page.extractText()
    # print(page_content.encode('utf-8').strip())
    # profile = get_my_jobs()
    # create_plot(profile)
    # print(profile)
    profile = []
    locations = ['Boston', 'Burlington', 'Seattle', 'Tokyo', 'Berlin', 'Munich', 'San Francisco', 'Redwood', 'Palo Alto'
                 'Waltham', 'Cambridge', 'Los Angeles', 'Chicago', 'Singapore', 'Hong Kong']
    jobs = apply(profile, locations)
    # create_plot(jobs)
