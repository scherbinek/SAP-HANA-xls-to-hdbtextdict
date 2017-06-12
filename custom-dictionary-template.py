# encoding=utf8
import sys
reload(sys)
sys.setdefaultencoding('utf8')
from xlrd import open_workbook
from operator import attrgetter

class Entity(object):
    def __init__(self, entity_category, entity_name_standard, standard_variant_name):
        self.entity_category = entity_category
        self.entity_name_standard = entity_name_standard
        self.standard_variant_name = standard_variant_name

    #def __str__(self):
    def __repr__(self):
        return repr((self.entity_category, self.entity_name_standard, self.standard_variant_name))
        #return("Element:\n"
        #       "  entity_category = {0}\n"
        #       "  entity_name_standard = {1}\n"
        #       "  standard_variant_name = {2}\n"
        #       .format(self.entity_category, self.entity_name_standard, self.standard_variant_name))

# Start reading Excel
wb = open_workbook('custom-dictionary-template.xlsx')
for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    items = []

    rows = []
    for row in range(1, number_of_rows):
        values = []
        for col in range(number_of_columns):
            value  = (sheet.cell(row,col).value)
            try:
                value = str(int(value))
            except ValueError:
                pass
            finally:
                values.append(value)
        item = Entity(*values)
        items.append(item)


# No Sorting, sorting by Excel
#items.sort()
#sorted(items, key=attrgetter('entity_category', 'entity_name_standard', 'standard_variant_name'))

# Start writing file
f = open('custom-dictionary-template.hdbtextdict','w')
# HEADER
f.write('<?xml version="1.0" encoding="UTF-8"?>' + '\n')
f.write('<dictionary xmlns="http://www.sap.com/ta/4.0">' + '\n\n')
l = len(items)
previous_ = next_ = None

# ENTITIES
for index, item in enumerate(items):
    if index == 0:
        f.write('\t<entity_category name="' + format(item.entity_category) + '">\n')
        f.write(format('\t\t<entity_name standard_form="' + item.entity_name_standard) + '">\n')
        if len(item.standard_variant_name) > 0:
            f.write(format('\t\t\t<variant name="' + item.standard_variant_name) + '" />\n')
    if index > 0:
        if item.entity_category == items[index - 1].entity_category:
            if item.entity_name_standard == items[index - 1].entity_name_standard:
                if len(item.standard_variant_name) > 0:
                    f.write(format('\t\t\t<variant name="' + item.standard_variant_name) + '" />\n')
            else:
                f.write(format('\t\t</entity_name>\n\n'))
                f.write(format('\t\t<entity_name standard_form="' + item.entity_name_standard) + '">\n')
                if len(item.standard_variant_name) > 0:
                    f.write(format('\t\t\t<variant name="' + item.standard_variant_name) + '" />\n')
        else:
            f.write(format('\t\t</entity_name>\n'))
            f.write(format('\t</entity_category>\n\n'))
            f.write('\t<entity_category name="' + format(item.entity_category) + '">\n')
            f.write(format('\t\t<entity_name standard_form="' + item.entity_name_standard) + '">\n')
            if len(item.standard_variant_name) > 0:
                f.write(format('\t\t\t<variant name="' + item.standard_variant_name) + '" />\n')
    #if index < (l - 1):
        #next_ = items[index + 1].entity_name_standard

# FOOTER
f.write(format('\t\t</entity_name>\n'))
f.write(format('\t</entity_category>\n\n'))
f.write('</dictionary>')
f.close()