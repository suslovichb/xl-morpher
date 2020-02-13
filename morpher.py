# import openpyxl
from openpyxl import load_workbook, Workbook
import pymorphy2

class Phrase:
    def __init__(self, raw_phrase):
        self.origin_phrase = raw_phrase
        

wb = load_workbook("excel1.xlsx")
ws = wb['Лист1']
column = ws['A']
items_list = [column[x].value for x in range(len(column))]
#print(items_list)

#new_wb = Workbook()
#new_ws = wb.active
#for item in items_list:
#    new_ws.append(item)
#new_wb.save("excel1_morphed.xlsx")

analyzer = pymorphy2.MorphAnalyzer()
morphed_items = {}
for item in items_list:
    parsed_item = analyzer.parse(item)[0]
    morphed_items[item] = {'nomn_sing': parsed_item.inflect({'sing', 'nomn'}).word,
                           'gent_sing': parsed_item.inflect({'sing', 'gent'}).word,
                           'datv_sing': parsed_item.inflect({'sing', 'datv'}).word,
                           'accs_sing': parsed_item.inflect({'sing', 'accs'}).word,
                           'ablt_sing': parsed_item.inflect({'sing', 'ablt'}).word,
                           'loct_sing': parsed_item.inflect({'sing', 'loct'}).word,
                           'nomn_plur': parsed_item.inflect({'plur', 'nomn'}).word,
                           'gent_plur': parsed_item.inflect({'plur', 'gent'}).word,
                           'datv_plur': parsed_item.inflect({'plur', 'datv'}).word,
                           'accs_plur': parsed_item.inflect({'plur', 'accs'}).word,
                           'ablt_plur': parsed_item.inflect({'plur', 'ablt'}).word,
                           'loct_plur': parsed_item.inflect({'plur', 'loct'}).word}
                           
#print(morphed_items)
for item in morphed_items:
    print(item, ": ", ", ".join(morphed_items[item].values()))
