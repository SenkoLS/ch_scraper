import requests
import wget
import xlsxwriter
from bs4 import BeautifulSoup
from datetime import datetime


class ParsCourseHunter:
    def __init__(self):
        self.session = requests.Session()
        self.set_session()
        self.dictionary_courses = {}

    def set_session(self):
        data_auth = {'e_mail': 'example@gmail.com', 'password': 'Pas$w0rd'}
        url_auth = 'https://coursehunter.net/sign-in'
        self.session.post(url_auth, data=data_auth)

    def get_start_links(self):
        url = 'https://coursehunter.net/categories'
        source_code = requests.get(url, cookies=self.session.cookies)
        soup = BeautifulSoup(source_code.text, 'html.parser')
        main_title = soup.find_all(name='a')
        start_links = []
        for title in main_title:
            href = title.get('href')
            index = str(href).find('https://coursehunter.net/source')
            excluded_link = ['https://t.me/coursehunters', 'https://coursehunters.online/',
                             'https://coursehunter.net/pricing', 'https://coursehunter.net/contacts',
                             'https://coursehunter.net/changelog', 'https://coursehunter.net/faq',
                             'https://coursehunter.net/categories', 'https://coursehunter.net/logout',
                             'https://coursehunter.net/history']
            if len(href) > 25 and href not in excluded_link and index == -1:
                start_links.append(href)

        return set(start_links) - self.get_repeated_links(start_links)

    def get_repeated_links(self, start_links):
        need_excluded = []
        for link in start_links:
            flag = False
            if str(link).count('/') == 3:
                for link2 in start_links:
                    if str(link2).count('/') > 3 and str(link2).find(link) != -1:
                        flag = True
            if flag:
                need_excluded.append(link)
        return set(need_excluded)

    def download_all_video_off_course(self):
        # The course URL is the web address of the course page
        url = 'https://coursehunter.net/course/ios-programmirovanie-na-swift-v-xcode-max-level-50-chasov'
        source_code = requests.get(url, cookies=self.session.cookies)
        soup = BeautifulSoup(source_code.text, 'html.parser')
        main_title = soup.find_all(name='link', attrs={'itemprop': 'url'})
        for title in main_title:
            filename = str(title.get('href')).replace('https://vs2.coursehunter.net/', '').replace('/', '_').replace(
                '-',
                '_')
            print('Загрузка файла: ', filename)
            wget.download(title.get('href'), filename)
            print('Загрузка файла ', filename, ' завершена')

    def get_content_course_links(self, link):
        video_links = []
        zip_links = []
        dic = {}
        source_code = requests.get(link, cookies=self.session.cookies)
        soup = BeautifulSoup(source_code.text, 'html.parser')

        # Get course name
        main_title = soup.find_all(name='p', attrs={'class': 'hero-description'})[0].text
        dic.update({'name': main_title})

        # Get course language
        main_title = soup.find_all(name='div', attrs={'class': 'course-box-value'})
        language = '-'
        for tag in main_title:
            if str(tag).find('Русский') != -1:
                language = 'Русский'
            if str(tag).find('English') != -1:
                language = 'Английский'
        dic.update({'language': language})

        main_title = soup.find_all(name='link', attrs={'itemprop': 'url'})
        for title in main_title:
            video_links.append(title.get('href'))
        dic.update({'video_links': video_links})

        main_title = soup.find_all(name='a', attrs={'title': 'Download course materials'})
        for title in main_title:
            zip_links.append(title.get('href'))
        dic.update({'zip': zip_links})

        return dic

    def get_course_links_from_the_start_link(self, start_link):
        current_page = 1
        all_course_links = []
        while True:
            link = start_link + '?page=' + str(current_page)
            course_links = self.get_course_links_from_the_page(link)
            if len(course_links) < 2:
                break
            else:
                all_course_links += course_links
            current_page += 1
        return set(all_course_links)

    def get_course_links_from_the_page(self, link):
        print('Старт обработки страницы: ', link)
        course_links = []
        source_code = requests.get(link, cookies=self.session.cookies)
        soup = BeautifulSoup(source_code.text, 'html.parser')
        main_title = soup.find_all(name='picture')
        for title in main_title:
            course_links.append(title.get('data-link'))
        print('Завершение обработки страницы: ', link)
        return [link for link in course_links if link]

    def get_all_courses(self):
        start_links = self.get_start_links()
        all_courses = set()
        for link in start_links:
            all_courses = all_courses.union(self.get_course_links_from_the_start_link(link))

        return all_courses

    def write_file_all_courses(self):
        with open('all_courses.txt', 'w') as file:
            for line in self.get_all_courses():
                file.write(line + '\r')

    def read_file_all_courses(self):
        with open('all_courses.txt', 'r') as file:
            return file.read().splitlines()


if __name__ == "__main__":
    start_time = datetime.now()
    parser = ParsCourseHunter()
    parser.write_file_all_courses()

    # Создание новой книги и листа
    workbook = xlsxwriter.Workbook('courses.xlsx')
    worksheet = workbook.add_worksheet()

    # Формат для ререноса текста
    cellWrapText = workbook.add_format()
    cellWrapText.set_text_wrap()
    cellWrapText.set_align("left")
    cellWrapText.set_align("top")
    cellWrapText.set_border()

    worksheet.name = "Курсы"

    worksheet.set_column('A:A', 55)
    worksheet.set_column('B:B', 55)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 100)
    worksheet.set_column('E:E', 100)

    count_rows = 1
    for link in parser.read_file_all_courses():
        print(str(count_rows) + ') Старт обработки курса: ', link)
        d = parser.get_content_course_links(link)
        print('Завершение получения данных о курсе: ', link)

        worksheet.write_string('B' + str(count_rows), link, cellWrapText)

        for k, v in d.items():
            if k == 'name':
                worksheet.write_string('A' + str(count_rows), v, cellWrapText)
            if k == 'language':
                worksheet.write_string('C' + str(count_rows), v, cellWrapText)
            if k == 'video_links':
                worksheet.write_string('D' + str(count_rows), '\r\n'.join([str(link) for link in v]), cellWrapText)
            if k == 'zip':
                worksheet.write_string('E' + str(count_rows), '\r\n'.join([str(link) for link in v]), cellWrapText)
        count_rows += 1
    workbook.close()
    print(datetime.now() - start_time)
