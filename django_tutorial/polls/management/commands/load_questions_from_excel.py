from django.utils import timezone

from django.core.management.base import BaseCommand
from polls.models import Question
from openpyxl import load_workbook


class Command(BaseCommand):
    help = 'Creates questions and choices from spreadsheet.'

    def add_arguments(self, parser):
        parser.add_argument('path', nargs='+', type=str)

    def handle(self, *args, **options):
        print(options['path'])

        wb = load_workbook(filename=options['path'][0])

        ws = wb.active

        row = 2

        while ws.cell(column=1, row=row).value:
            question = Question(question_text=ws.cell(column=1, row=row).value, pub_date=timezone.now())
            question.save()
            col = 2
            while ws.cell(column=col, row=row).value:
                question.choice_set.create(choice_text=ws.cell(column=col, row=row).value)
                col += 1
            row += 1

        print('Done creating questions.')
