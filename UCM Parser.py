from lxml import etree
from openpyxl import Workbook
from deep_translator import GoogleTranslator


def get_positions(positions):
    arr = [['Подразделение', 'Номер части', 'Краткие сведения', 'Широта', 'Долгота']]
    for pos in positions:
        try: name = pos.xpath('kml:name/text()', namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: name = ""
        try: unit = pos.xpath('kml:ExtendedData/kml:Data[@name="Military Unit Number"]/kml:value/text()',
                              namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: unit = ""
        try: description = pos.xpath('kml:ExtendedData/kml:Data[@name="description"]/kml:value/text()',
                                     namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: description = ""
        description = description.replace('\n', ' ').replace('\xa0', '')
        try: cords = pos.xpath('kml:Point/kml:coordinates/text()',
                               namespaces={"kml":"http://www.opengis.net/kml/2.2"})[0]
        except: cords = ""
        cords = cords.replace(' ', '').replace('\n', '').split(",")
        try: lat = cords[1]
        except: lat = ''
        try: lon = cords[0]
        except: lon = ''
        name = translator.translate(name)
        unit = translator.translate(unit)
        description = translator.translate(description)
        lat = lat.replace(".", ",")
        lon = lon.replace(".", ",")
        print(f"{name} | {unit} | {description} | {lat} | {lon}")
        arr.append([name, unit, description, lat, lon])
    arr.pop(1)
    arr.pop()
    return arr


def make_file(name, arr):
    print(name)
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
          "4. Скачаный файл расположите рядом с файлом парсер и переименнуйте его в 'Ukraine Control Map.kml'\n"
          "Когда все будет готово Enter\n")

    input()

    translator = GoogleTranslator(source='auto', target='ru')

    tree = etree.parse('Ukraine Control Map.kml')
    u_positions = tree.xpath('//kml:Folder', namespaces={"kml":"http://www.opengis.net/kml/2.2"})[2]
    r_positions = tree.xpath('//kml:Folder', namespaces={"kml":"http://www.opengis.net/kml/2.2"})[3]

    make_file("Позиции Украина", get_positions(u_positions))
    make_file("Позиции Россия", get_positions(r_positions))