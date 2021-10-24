from selenium import webdriver
from selenium.webdriver import ActionChains
import auth
import time
import win32api
import os
import re
import winsound


def month_to_str(month):
    month_str = ""
    if month < 10:
        month_str = "0" + str(month)
    else:
        month_str = str(month)
    return month_str


def generate_new_path_name(download_dir, journal_name, year_start, month_start, year_end, month_end):
    path_article_name = ""
    article_words = journal_name.split(" ")
    for word in article_words:
        c_word = word.capitalize()
        path_article_name = path_article_name + c_word + " "
    path_article_name = path_article_name.strip()
    path_article_name = path_article_name.replace(" ", "-")
    if month_start == month_end and year_start == year_end:
        new_path = path_article_name + "-" + year_start + month_to_str(month_start)
    else:
        new_path = path_article_name + "-" + year_start + month_to_str(month_start) + "-" + year_end + month_to_str(month_end)

    return new_path


if __name__ == "__main__":
    url = "https://search.ebscohost.com/login.aspx?authtype=sso&custId=s8983984"
    my_username = auth.username    # 用户名
    my_password = auth.password    # 密码
    download_dir = "E:\\Test3"    # 文件下载路径

    article_name = "harvard business review"    # 期刊名称

    month_start = 1        # 起始月份
    year_start = "2020"    # 起始年份
    month_end = 1          # 结束月份
    year_end = "2020"      # 结束年份

    new_path_name = generate_new_path_name(download_dir, article_name, year_start, month_start, year_end, month_end)
    full_new_path = download_dir + "\\" + new_path_name
    if not os.path.exists(full_new_path):
        if not os.path.exists(download_dir):
            os.mkdir(download_dir)
        os.mkdir(full_new_path)
    download_dir = full_new_path

    # browser_path = 'C:/EBSCO/chromedriver.exe'    # 爬虫exe文件所在路径
    browser_path = 'D:/selenium_chromedriver_88/chromedriver.exe'

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,
        "download.extensions_to_open": "applications/pdf"})

    driver = webdriver.Chrome(executable_path=browser_path,
                              chrome_options=chrome_options)

    driver.get(url)  # 请求网页
    print(driver)  # 打印浏览器对象
    driver.maximize_window()
    html = driver.page_source

    username = driver.find_element_by_xpath('//*[@id="username"]')
    username.send_keys(my_username)
    """ 输入密码 """
    password = driver.find_element_by_xpath('//*[@id="password"]')
    password.send_keys(my_password)
    time.sleep(2)
    """ 点击登录 """
    driver.find_element_by_name("_eventId_proceed").click()

    time.sleep(15)
    # driver.find_element_by_xpath("//*[@id='ctl00_ctl00_MainContentArea_MainContentArea_selectAll']").click()
    """ Limit the database to Business Source Complete """
    driver.find_element_by_xpath('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_SelectDbControl_dbList_ctl30_itemCheck"]').click()
    time.sleep(2)
    # Click Continue button
    driver.find_element_by_xpath("//*[@id='ctl00_ctl00_MainContentArea_MainContentArea_continue1']").click()
    time.sleep(5)

    month_option_start = str(month_start + 1)
    month_option_end = str(month_end + 1)
    month_start_xpath = '//*[@id="common_DT1"]/option[' + month_option_start + ']'
    month_end_xpath = '//*[@id="common_DT1_ToMonth"]/option[' + month_option_end + ']'

    """ Enter the article name to be searched """

    search_input = driver.find_element_by_xpath("//*[@id='SearchTerm1']")
    search_input.send_keys(article_name)
    time.sleep(2)

    """ Limit the start and end of month & year """
    driver.find_element_by_xpath('//*[@id="common_DT1"]').click()
    time.sleep(0.5)
    driver.find_element_by_xpath(month_start_xpath).click()
    time.sleep(0.5)
    driver.find_element_by_xpath('//*[@id="common_DT1_ToMonth"]').click()
    time.sleep(0.5)
    driver.find_element_by_xpath(month_end_xpath).click()
    time.sleep(0.5)
    year_start_input = driver.find_element_by_xpath('//*[@id="common_DT1_FromYear"]')
    year_start_input.send_keys(year_start)
    time.sleep(1)
    year_end_input = driver.find_element_by_xpath('//*[@id="common_DT1_ToYear"]')
    year_end_input.send_keys(year_end)
    time.sleep(1)
    driver.find_element_by_xpath("//*[@id='SearchButton']").click()
    time.sleep(5)

    # Limit the publication to Harvard Business Review
    driver.find_element_by_xpath('//*[@id="multiSelectCluster_JournalTrigger"]').click()
    time.sleep(2)
    """ 选择Publication """
    driver.find_element_by_xpath('//*[@id="_cluster_Journal%24harvard+business+review"]').click()    # Harvard Business Review
    # driver.find_element_by_xpath('//*[@id="_cluster_Journal%24journal+of+cleaner+production"]').click()    # Academy of Management Journal
    # driver.find_element_by_xpath('//*[@id="_cluster_Journal%24administrative+science+quarterly"]').click()    # Administrative Science Quarterly
    # driver.find_element_by_xpath('/html/body/form/div[2]/div[1]/nav/div/div/div[2]/div/div[3]/div[3]/div/ul/li[1]/input').click()    # Publiction 列表中的第一个
    time.sleep(8)

    """选择每页显示50个结果"""
    # driver.find_element_by_xpath('//*[@id="lnkPageOptions"]').click()
    # time.sleep(1)
    # driver.find_element_by_xpath('//*[@id="pageOptions"]/li[3]/ul/li[6]/a').click()
    # time.sleep(5)

    success_downloads_count = 0
    download_fail_list = []
    article_title_list = []
    while True:
        all_articles_link = driver.find_elements_by_link_text("PDF Full Text")
        for element in all_articles_link:
            actionChains = ActionChains(driver)

            actionChains.context_click(element).perform()
            time.sleep(0.5)
            win32api.keybd_event(84, 0, 0, 0)

            handles = driver.window_handles
            driver.switch_to.window(handles[-1])

            time.sleep(5)    # 打开新标签页等待时间
            article_title = driver.find_element_by_xpath('//*[@id="aspnetForm"]/h1').get_attribute('title')
            article_title = re.sub(r'[^A-Za-z0-9 ]+', '', article_title)
            article_title = article_title + ".pdf"

            pdf_source_url = driver.find_element_by_xpath('//*[@id="pdfUrl"]').get_attribute('value')
            # print(pdf_source_url)

            try:
                print("_________________________________________________________________________________________________")
                print(" ***** " + "《" + article_title + "》" + " is being downloaded..." + " ***** ")
                driver.get(pdf_source_url)
                time.sleep(5)
                download_status = False
                download_duration = 5
                while not download_status:
                    download_status = True
                    temp = 0
                    for i in os.listdir(download_dir):
                        if i.endswith(".crdownload"):
                            download_status = False
                            time.sleep(2)
                            download_duration += 2

                time.sleep(1)
                print("《" + article_title + "》" + " 下载成功！ 用时 " + str(download_duration) + "秒")
                success_downloads_count += 1
                article_title_list.append(article_title)
            except:
                print("《" + article_title + "》" + " 下载失败")
                download_fail_list.append(article_title)

            print("下载成功： " + str(success_downloads_count) + " 篇； 下载失败： " + str(len(download_fail_list)) + " 篇。")

            driver.close()
            current_handles = driver.window_handles
            driver.switch_to.window(current_handles[0])
            time.sleep(2)

        try:
            # driver.find_element_by_xpath('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_bottomMultiPage_rptPageLinks_ctl01_lnkPageLink"]')    # page 2
            driver.find_element_by_xpath('//*[@id="ctl00_ctl00_MainContentArea_MainContentArea_bottomMultiPage_lnkNext"]').click()   # next
            time.sleep(10)
        except:
            break

    time.sleep(2)
    print("**************************************************************************************************")
    print("下载任务结束。")

    final_dir_lists = os.listdir(download_dir)
    final_dir_lists.sort(key=lambda fn: os.path.getctime(download_dir + '\\' + fn))
    target_index = -1 * success_downloads_count
    target_list = final_dir_lists[target_index:]
    os.chdir(download_dir)
    for i in range(len(target_list)):
        try:
            os.rename(target_list[i], article_title_list[i])
        except:
            print("_________________________________________________________________________________________________")
            print("《" + article_title_list[i] + "》" + " 文件未重命名，现文件名为： " + target_list[i])

    time.sleep(5)
    print("文件重命名结束。")

    winsound.Beep(440, 3000)

