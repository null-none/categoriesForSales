from django.core.management.base import BaseCommand, CommandError

import xlrd
import os

from category.models import Category

class Command(BaseCommand):

    def handle(self, *args, **options):
        rb = xlrd.open_workbook('{0}/category.xls'.format(os.path.dirname(__file__)), formatting_info=True)
        sheet = rb.sheet_by_index(0)
        for rownum in range(sheet.nrows):
            row = sheet.row_values(rownum)
            try:
                if row[1] == '2':
                    Category.objects.create(title_en=row[3], title_ru=row[4], name_en=row[2], name_ru=row[2])
                if row[1] == '3':
                    parent = Category.objects.get(name_en=row[2].split('.')[0])
                    BrandCategory.objects.create(parent=parent, title_en=row[3], title_ru=row[4], name_en=row[2], name_ru=row[2])
            except Exception, e:
                pass
        self.stdout.write("Finished", ending='')