#Recursos
import time
import warnings
import traceback
import pandas as pd
import math

from VAST import calc
from GOOGLE import gapi
from pytz import timezone
from datetime import datetime
from thefuzz import fuzz, process
from pickle import FALSE, TRUE

#Ignorar DeprecationWarning
warnings.filterwarnings("ignore", category=DeprecationWarning)

#Jalada de correos a sheets
def scrape():
    #Llamadas al API
    error = None
    try:
        gt = gapi(2).mail()
        gapi(3).sheets(gt,'A','last')
    except Exception as e:
        error = e
    if error is None:
        gapi(2).read()
        
while TRUE:
    time.sleep(10)
    try:
        scrape()
    except:
        print('Volviendo a intentar')
        continue
