from django.core.management.base import BaseCommand
from polls.models import Question


class Command(BaseCommand):
    help = 'Prints all questions and related choices.'

    def handle(self, *args, **kwargs):
        for question in Question.objects.all():
            self.stdout.write(question.question_text)
            for choice in question.choice_set.all():
                self.stdout.write(f'    -{choice.choice_text}')
