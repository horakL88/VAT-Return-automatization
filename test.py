import requests
from bs4 import BeautifulSoup

url = requests.get("https://www.raiffeisen.hu/hasznos/arfolyamok/vallalati/valutaarfolyamok?p_p_id=hu_raiffeisen_website_currency_rates_web_portlet_CurrencyRatesPortlet_INSTANCE_VyVoydIU31cU&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view").text

soup = BeautifulSoup(url)
