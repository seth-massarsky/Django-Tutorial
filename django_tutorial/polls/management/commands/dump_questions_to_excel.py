import datetime

from django.core.management.base import BaseCommand
from openpyxl import Workbook
from polls.models import Question

LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


class Command(BaseCommand):
    help = 'Dumps all questions and choices to an Excel file.'

    def handle(self, *args, **kwargs):
        wb = Workbook()
        ws = wb.active
        ws.append(['Question'])
        max_row_length = 0

        for question in Question.objects.all():
            row = [question.question_text] + [choice.choice_text for choice in question.choice_set.all()]
            row_length = len(row)
            max_row_length = row_length if row_length > max_row_length else max_row_length
            ws.append(row)

        for x in range(1, max_row_length):
            ws[f'{LETTERS[x]}1'] = f'Choice {x}'

        wb.save(f'poll_dump_{str(datetime.datetime.now())}.xlsx')
