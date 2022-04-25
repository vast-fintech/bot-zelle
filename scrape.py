
   
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
    # try:
    gt = gapi(2).mail()
    gapi(3).sheets(gt,'A','last')
    #gapi(2).read()
    # except:
    #     time.sleep(2)
    #     scrape()
        
while TRUE:
    time.sleep(1)
    scrape()