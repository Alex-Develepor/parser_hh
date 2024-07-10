import requests
import openpyxl

key_words = ['Трейдер', 'Алгоритмический трейдер', 'финансовый советник/консультант', 'инвестиционный аналитик',
             'финансовый аналитик', 'риск менеджер', 'менеджер по продажам инвестиционных продуктов',
             'персональный брокер', 'риск аналитик', 'портфельный управляющий']

vacancies = []


def create_excel_file(
):
    wb = openpyxl.load_workbook('vacancies.xlsx')
    sheet = wb.active
    name = sheet['B2']
    name.value = 'Название вакансии'
    salary_from = sheet['C2']
    salary_from.value = 'Зарплата от'
    salary_to = sheet['D2']
    salary_to.value = 'Зарплата до'
    city = sheet['E2']
    city.value = 'Город'
    employer = sheet['F2']
    employer.value = 'Компания'
    description = sheet['G2']
    description.value = 'Ключевые навыки'
    url = sheet['H2']
    url.value = 'Ссылка'
    wb.save('vacancies.xlsx')
    wb.close()


def fill_xlsx_file(dict_info: dict):
    print('сохраняем данные в exl')
    wb = openpyxl.load_workbook('vacancies.xlsx')
    sheet = wb.active
    i = 3
    for url, data in dict_info.items():
        name = sheet[f'B{i}']
        name.value = data[2]
        salary_from = sheet[f'C{i}']
        salary_from.value = data[3]
        salary_to = sheet[f'D{i}']
        salary_to.value = data[4]
        city = sheet[f'E{i}']
        city.value = data[5]
        employer = sheet[f'F{i}']
        employer.value = data[0]
        description = sheet[f'G{i}']
        description.value = data[1]
        vacancy_url = sheet[f'H{i}']
        vacancy_url.value = url
        i += 1
    wb.save('vacancies.xlsx')
    wb.close()


def get_vacancies(page=0, per_page=10):
    result_dict = {}
    for key_word in key_words:
        for i in range(0, page + 1):
            params = {
                'text': key_word,
                'page': i,
                'per_page': per_page
            }
            request = requests.get('https://api.hh.ru/vacancies', params=params).json()
            try:
                for vacancy in request['items']:
                    if key_word.lower() in vacancy['name'].lower():
                        vacancy_name = vacancy['name']
                        print(f'Проверяем вакансию {vacancy_name}')
                        try:
                            vacancy_salary_from = vacancy['salary']['from']
                        except TypeError:
                            vacancy_salary_from = 'Не указана'
                        try:
                            vacancy_salary_to = vacancy['salary']['to']
                        except TypeError:
                            vacancy_salary_to = 'Не указана'
                        try:
                            vacancy_address_city = vacancy['address']['city']
                        except TypeError:
                            vacancy_address_city = 'Не указан'
                        vacancy_url = vacancy['alternate_url']
                        vacancy_skill = vacancy['snippet']['requirement']
                        employer_name = vacancy['employer']['name']
                        data = [employer_name, vacancy_skill, vacancy_name,
                                vacancy_salary_from, vacancy_salary_to, vacancy_address_city]
                        result_dict[vacancy_url] = data
            except KeyError:
                print('Ошибка, HH ограничил поиск для этого адреса! Попробуйте через 30 мин')
    print('Поиск окончен')
    return result_dict


if __name__ == '__main__':
    print('старт')
    create_excel_file()
    page = 5
    per_page = 200
    data = get_vacancies(page, per_page)
    fill_xlsx_file(data)
    print('Работа выполнена')

