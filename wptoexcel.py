from openpyxl import load_workbook
import os
import argparse
import re
from openpyxl import Workbook
from frozendict import frozendict


config = {
    "REPORT_SIZE": 1000,
    "REPORT_DIR": "reports",
    "WP_DIR": "img",
    "WPTXT_FILE": "WhatsApp_1811.txt",
    "BASE_DIR": os.path.dirname(os.path.abspath(__file__)),
    # "TARGET_FILE": "check_hardware_1411.xlsx"
    # "TARGET_FILE": "check_hardware_15111.xlsx"
    # "TARGET_FILE": "check_hardware_1511.xlsx"
    "TARGET_FILE": "check_hardware_1811.xlsx"
}


def get_img_name(msg):
    if 'IMG-' in msg:
        img_search = re.search(r'IMG-\d{8}-\w+\.jpg', msg)
        img_name = img_search.group(0)
        return img_name


def make_row_from_msg():
    previous_img = ''

    def get_line_from_msg(msg):
        unit_name = ''
        hardware_name = ''
        nonlocal previous_img
        imgname = get_img_name(msg)

        if not imgname:
            unit_name = get_unit_name(msg)
            hardware_name = get_hardware_name(msg)
            if previous_img:

                row = (unit_name or '', hardware_name or '', msg, previous_img)
                previous_img = ''
                return row
        else:
            if previous_img:

                row = (unit_name, hardware_name, '', previous_img)
                previous_img = imgname
                return row
            previous_img = imgname
    return get_line_from_msg


def get_unit_name(msg):
    unit_dict = frozendict({
        'псч': 'ПЧ',
        'спч': 'ПЧ',
        'пч': 'ПЧ',
        'ронд': 'РОНД',
        'оп': 'ОП',
        'узао': 'УЗАО',
        'осенняя': 'УЗАО Осенняя',
    })
    msg = msg.lower()
    for unit, target in unit_dict.items():
        if unit in msg:
            # print(unit, msg)
            res = re.search(r'\d{0,3}[ ]?%s(\d{0,3}[ ]?\d{0,3})' % unit, msg)
            unit_str = res.group(0).strip()
            if len(res.group(1)) > 1:
                # print(len(res.group(1)))
                unit_id = res.group(1).strip()
            else:
                unit_id = unit_str.replace(unit, '').strip()
            return '{} {}'.format(target, unit_id)


def get_hardware_name(msg):
    hardware_dict = {
        'монитор': 'монитор',
        'сист блок': 'системный блок',
        'системный блок': 'системный блок',
        'МФУ': 'МФУ',
        'Корал': 'АТС',
        'Coral': 'АТС',
        'АТС': 'АТС',
        'Моноблок': 'моноблок',
        'Ноутбук': 'ноутбук',
        'Awaya': 'телефон',
        'Радиостанц': 'радиостанция',
        'принтер': 'принтер',
        'Стрелец Мониторинг': 'Стрелец Мониторинг',
        'Стрелец-Мониторинг': 'Стрелец Мониторинг',
        'телефон': 'телефон',
        'МФЦ': 'МФУ',
        'Пульт ГГС': 'пульт',
        'Пульт': 'пульт',
        'бесперебойник': 'ИБП',
        'UPS': 'ИБП',
        'ИБП': 'ИБП',
        'шредер': 'шредер',
        'карт': 'АТС',
        'антен': 'антена',
        'Коммутатор': 'коммутатор',
        'Сканер': 'сканер',
        'Телевизор': 'телевизор',
        'TV': 'телевизор',
        'Факс': 'факс',
        'Проектор': 'проектор',
    }
    msg = msg.lower()
    for kunit, target in hardware_dict.items():
        if kunit.lower() in msg:
            return target


def load_data_from_file(
        file_name=config["WPTXT_FILE"],
        file_path=os.path.join(config["BASE_DIR"], config["WP_DIR"])):
    # try to read json object
    get_row = make_row_from_msg()
    book = Workbook()
    sheet = book.active
    try:
        with open(os.path.join(file_path, file_name), "r") as fh:
            for line in fh:
                row = get_row(line)
                if row:
                    sheet.append(row)
                    # img = Image(os.path.join(config['WPTXT_DIR'], row[-1]))
                    # comment = Comment('file:///wp/{}'.format(row[-1]), "Author")
                    # path1 = sheet['H1']
                    # sheet['D'][-1].add_image()
                    # sheet.add_image(img.thumbnail((50, 50)))
                    sheet['D'][-1].hyperlink = '{}\{}'.format(config['WP_DIR'], row[-1])
            book.save(config['TARGET_FILE'])
    except Exception as ex:
        raise argparse.ArgumentTypeError(
            "file:{0} is not a valid file. Read error: {1}".format(file_name, ex))


if __name__ == '__main__':
    # get_row = make_row_from_msg()
    # print(get_row(
    #     '15.11.2019, 12:57 - Семлянских Александр Михайлович: ‎IMG-20191115-WA0192.jpg (файл добавлен)'))
    # print(get_row('СПТ 27'))
    # print(get_row('15.11.2019, 12:58 - Влад М: ‎IMG-20191115-WA0195.jpg (файл добавлен)'))
    # print(get_row('Моноблок КИС УСС ПСЧ28 пункт связи'))
    # print(get_row('15.11.2019, 14:27 - Влад М: ‎IMG-20191115-WA0111.jpg (файл добавлен)'))
    # print(get_row(
    #     '15.11.2019, 14:27 - Влад М: ‎IMG-20191115-WA0222.jpg (файл добавлен)'))
    # print(get_row(
    #     '15.11.2019, 14:27 - Влад М: ‎IMG-20191115-WA0333.jpg (файл добавлен)'))
    # print(get_row(
    #         '15.11.2019, 15:27 - Семлянских Александр Михайлович: ‎IMG-20191115-WA0417.jpg (файл добавлен)'))
    # print(get_row(
    #     '4 ПСЧ комната хранения  14а - бесперебойник'))
    load_data_from_file()

    # msg = '28 СПЧ МФУ отдел ГПН..'
    # print(msg)
    # print(get_unit_name(msg))