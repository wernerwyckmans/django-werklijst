from django.shortcuts import render
import openpyxl
import core.controlestaf as controlestaf
import core.controleassist as controleassist
from . import forms

from django.contrib.auth import authenticate, login, logout
from django.http import HttpResponseRedirect, HttpResponse
from django.urls import reverse
from django.contrib.auth.decorators import login_required


# Create your views here.

@login_required
def index(request):
    return render(request, 'core/welkom.html', {})


@login_required
def user_logout(request):
    logout(request)
    return HttpResponseRedirect(reverse('index'))


@login_required
def staf_view(request):
    form = forms.UserForm()

    if request.method == 'POST':
        form = forms.UserForm(request.POST, request.FILES)

        if form.is_valid():

            excel_file = request.FILES["excel_file"]
            file_name = str(excel_file).split(".")[0]
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            weekkeuze = form.cleaned_data['weken']

            control_data = controlestaf.main(wb, weekkeuze)

            context = {'control_data': control_data, 'weekkeuze': weekkeuze, 'file_name': file_name}

            return render(request, 'core/resultpage.html', context)

    else:
        form = forms.UserForm()
    return render(request, 'core/form_staf.html', {"form": form})


@login_required
def assistenten_view(request):
    form = forms.UserForm()

    if request.method == 'POST':
        form = forms.UserForm(request.POST, request.FILES)

        if form.is_valid():
            excel_file = request.FILES["excel_file"]
            file_name = str(excel_file).split(".")[0]
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            weekkeuze = form.cleaned_data['weken']

            control_data = controleassist.main(wb, weekkeuze)

            context = {'control_data': control_data, 'weekkeuze': weekkeuze, 'file_name': file_name}

            return render(request, 'core/resultpage.html', context)

    else:

        form = forms.UserForm()
    return render(request, 'core/form_assist.html', {"form": form})


def user_login(request):

    if request.method == 'POST':
        username = request.POST.get('username')
        password = request.POST.get('password')

        user = authenticate(username=username, password=password)

        if user:
            if user.is_active:
                login(request, user)
                return HttpResponseRedirect(reverse('index'))
            else:
                return HttpResponse("ACCOUNT NOT ACTIVE!")

        else:
            print("Someone tried to login and failed")
            print("Username: {} and password: {}".format(username, password))
            return HttpResponse("Invalid login details")

    else:
        return render(request, 'core/login.html', {})


