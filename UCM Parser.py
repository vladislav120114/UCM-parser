from lxml import etree
from openpyxl import Workbook
from deep_translator import GoogleTranslator
import os
import json

def get_positions(positions):
    arr = [['Подразделение', 'Номер части', 'Краткие сведения', 'Предыдущие появления', 'Широта', 'Долгота']]
    translation = {}
    if os.path.exists('translation.json'):
        with open('translation.json', encoding="utf-8") as f:
            translation = json.load(f)
    for pos in positions:
        try: name = pos.xpath('kml:name/text()', namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: continue
        try: unit = pos.xpath('kml:ExtendedData/kml:Data[@name="Military Unit Number"]/kml:value/text()',
                              namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: unit = ""
        try: description = pos.xpath('kml:ExtendedData/kml:Data[@name="description"]/kml:value/text()',
                                     namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: description = ""
        try: geo = pos.xpath('kml:ExtendedData/kml:Data[@name="Older Geolocations"]/kml:value/text()',
                                     namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: geo = ""
        description = description.replace('\n', ' ').replace('\xa0', '')
        try: cords = pos.xpath('kml:Point/kml:coordinates/text()',
                               namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: cords = ""
        cords = cords.replace(' ', '').replace('\n', '').split(",")
        try: lat = cords[1]
        except: lat = ''
        try: lon = cords[0]
        except: lon = ''
        if name in translation:
            t_name = translation[name]["t_name"]
        else:
            try:
                translation[name] = {"t_name": translator.translate(name)}
                t_name = translation[name]["t_name"]
            except: continue
        if "unit" in translation[name]:
            unit = translation[name]["unit"]
        else:
            try:
                translation[name]["unit"] = translator.translate(unit)
                unit = translation[name]["unit"]
            except: pass
        if "description" in translation[name]:
            description = translation[name]["desc"]
        else:
            try:
                translation[name]["desc"] = translator.translate(description)
                description = translation[name]["desc"]
            except: pass
        lat = lat.replace(".", ",")
        lon = lon.replace(".", ",")
        translation[name]["lat"] = lat
        translation[name]["lon"] = lon

        print(f"{t_name} | {unit} | {description} | {lat} | {lon}")
        arr.append([t_name, unit, description, geo, lat, lon])
        if t_name == "Зона ожидания для российских войск":
            break
    arr.pop(1)
    arr.pop()
    with open("translation.json", "w", encoding="utf-8") as f:
        json.dump(translation, f, ensure_ascii=False, indent=4)
    return arr


def make_file(name, arr):
    wb = Workbook()
    ws = wb.active
    for i in arr:
        try:
            ws.append(i)
        except:
            ws.append(["Ошибка"])
    wb.save(f"{name}.xlsx")


if __name__ == ('__main__'):

    print("\nПарсер Google карт СВО.\n"
          "Для работы парсера.\n"
          "1. Перейдите по ссылке https://www.google.com/maps/d/u/0/viewer?mid=1xPxgT8LtUjuspSOGHJc2VzA5O5jWMTE\n"
          "2. Зайдите в опции и выберите 'Скачать KML' (Download KML).\n"
          "3. В настройках выберите опцию 'Экспортировать KML вместо KMZ' (Export as KML instead of KMZ)\n"
          "4. Скачаный файл расположите рядом с файлом парсер и переименнуйте его в 'Ukraine Control Map.kml'\n")

    translator = GoogleTranslator(source='auto', target='ru')

    tree = etree.parse('Ukraine Control Map.kml')
    u_positions = tree.xpath('//kml:Folder', namespaces={"kml":"http://www.opengis.net/kml/2.2"})[2]
    r_positions = tree.xpath('//kml:Folder', namespaces={"kml":"http://www.opengis.net/kml/2.2"})[3]

    u_positions = get_positions(u_positions)
    r_positions = get_positions(r_positions)

    make_file("Позиции Украина", u_positions)
    make_file("Позиции Россия", r_positions)