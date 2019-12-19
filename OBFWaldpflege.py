################################################################################
#
#   fuc_tbl_waldpflegeplan (fr)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

class OBFWaldpflege(object):

    def __init__(self, data):

        self.data = data

        data.rename(columns={'Unnamed: 2':'Massnahme', 'Unnamed: 3':'Einheit'}, inplace=True)

        # fill nan with 0
        self.data = self.data.fillna(0)


    def printit(self):
        print(self.data)
        return(self.data)

    ### 7.1.6 Auszug aus dem Waldpflegeplan

    def fuc_tbl_waldpflegeplan(self, fr):

        # make pivot table
        table = pd.pivot_table(self.data, index=['Massnahme', 'Einheit', 'Soll/Ist/Stand|Forstrevier'],columns = ['Geschäftsjahr'], \
                               values=[fr], \
                              aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table.columns = table.columns.droplevel()

        # round
        table = round(table,1)

        # reset index
        table.reset_index(level=2, inplace=True)
        table.reset_index(level=1, inplace=True)
        table.reset_index(level=0, inplace=True)

        table = table[table['All'].notnull()]

        return table

        # rename index
        #table_wuech.rename(index={1:'sehr gering wüchsig', 2:'gering wüchsig', 3:'mittelwüchsig', 4:'wüchsig' , 5:'sehr wüchsig', 'All': 'Ges.'}, inplace=True)
        #table_wuech.rename(columns={'All':'Ges.'}, inplace=True)
