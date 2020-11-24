from django import forms
from django.core import validators

WEEK_KEUZE = [
    ('WEEK1', 'Week 1'),
    ('WEEK2', 'Week 2'),
    ('WEEK3', 'Week 3'),
    ('WEEK4', 'Week 4'),
    ('all', 'Alle Weken'),
]


class UserForm(forms.Form):
    excel_file = forms.FileField()
    weken = forms.CharField(label='Welke week wil je analyseren?', widget=forms.Select(choices=WEEK_KEUZE))
    botcatcher = forms.CharField(required=False,
                                 widget=forms.HiddenInput,
                                 validators=[validators.MaxLengthValidator(0)])
