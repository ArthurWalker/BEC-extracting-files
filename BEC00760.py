import pandas as pd
#import numpy as np
import os

class BEC00760(object):
    path = os.path.join('C:/Users/pphuc/Desktop/Docs/Current Using Docs/')

    def __init__(self,file):
        self.bec00760_file = pd.ExcelFile(self.path+file)