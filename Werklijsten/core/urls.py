from django.conf.urls import url, include
from core import views

app_name = "core"

urlpatterns = [
    url(r'^index/$', views.index, name="index"),
    url(r'^staf/$', views.staf_view, name='stafform'),
    url(r'^assistenten/$', views.assistenten_view, name='assistentenform'),
    url(r'^user_login/$', views.user_login, name='user_login'),
]