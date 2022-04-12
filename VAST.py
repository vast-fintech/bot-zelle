import pandas as pd
import numpy as np
import re
from datetime import datetime
from thefuzz import fuzz

class calc:
    def __init__(self, amount):
        self.amount = amount

    def tolerance(self):
        if self.amount <=3:
            return 0.5
        elif self.amount >=3 and self.amount <10 :
            return 0.1
        elif self.amount >=10 and self.amount <30:
            return 0.12
        elif self.amount >=30 and self.amount <60:
            return 0.15
        elif self.amount >=60:
            return 0.2
