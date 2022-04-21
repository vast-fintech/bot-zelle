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
    try:
        gt = gapi(2).mail()
    except:
        scrape()
    #Vaciado en SHEETS
    try:
        vast = gapi(3).sheets(gt,'A','last')
    except:
        scrape()
        
while TRUE:
    time.sleep(30)
    scrape()
