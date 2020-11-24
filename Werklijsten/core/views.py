from django.shortcuts import render
import openpyxl
import core.controlestaf as controlestaf
import core.controleassist as controleassist
from . import forms


# Create your views here.


def index(request):
    return render(request, 'core/welkom.html', {})


def staf_view(request):
    form = forms.UserForm()

    if request.method == 'POST':
        form = forms.UserForm(request.POST, request.FILES)

        if form.is_valid():
            excel_file = request.FILES["excel_file"]
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            weekkeuze = form.cleaned_data['weken']

            control_data = controlestaf.main(wb, weekkeuze)

            return render(request, 'core/resultpage.html', {'control_data': control_data})

    else:

        form = forms.UserForm()
        return render(request, 'core/form_staf.html', {"form": form})


def assistenten_view(request):
    form = forms.UserForm()

    if request.method == 'POST':
        form = forms.UserForm(request.POST, request.FILES)

        if form.is_valid():
            excel_file = request.FILES["excel_file"]
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            weekkeuze = form.cleaned_data['weken']

            control_data = controleassist.main(wb, weekkeuze)

            return render(request, 'core/resultpage.html', {'control_data': control_data})

    else:

        form = forms.UserForm()
        return render(request, 'core/form_assist.html', {"form": form})
