################################################################################
#
#
################################################################################

import numpy as np
import pandas as pd
from OBFDictionary import OBFDictionary

class OBFEt(object):

    def __init__(self, data_et, data_op):

        self.data_op = data_op
        self.data_et = data_et
        self.dic = OBFDictionary()
        self.bag = self.data_op['Baumart_group'].unique()
        self.bag = np.delete(self.bag, np.where((self.bag == 0) | (self.bag == 'SN') | (self.bag == 'SL')))

    def fuc_tbl_et_set(self):
        # make table - Ertragstafel pro BKL

        # get unique bkl
        bkl = np.sort(self.data_op['Betriebsklasse'].unique())[1:]
        # create emty list
        et_sets = []
        # loop through bkl
        for i in bkl:
            temp = np.sort(self.data_op[self.data_op['Betriebsklasse']==i]['Ertragstafelnummer'].unique()[1:])
            if 1 in temp:
                et_sets.append([int(i), 1, 'FI-Hochgebirge/TA-Württemberg'])
            elif 2 in temp:
                et_sets.append([int(i), 2, 'FI-Bayern/TA-Württemberg'])
            elif 3 in temp:
                et_sets.append([int(i), 3, 'FI-Bruck/TANWD'])
            elif 4 in temp:
                et_sets.append([int(i), 4, 'Weitra/TA-NWD'])
            elif (51 in temp) & (7 in temp):
                et_sets.append([int(i), 5, 'Fichte Tirol, Kalk niedrig'])
            elif (52 in temp) & (7 in temp):
                et_sets.append([int(i), 6, 'Fichte Tirol, Kalk mittel'])
            elif (53 in temp) & (7 in temp):
                et_sets.append([int(i), 7, 'Fichte Tirol, Silikat mittel'])
            elif (54 in temp) & (7 in temp):
                et_sets.append([int(i), 8, 'Fichte Tirol, Silikat hoch'])
            elif 55 in temp:
                et_sets.append([int(i), 9, 'Fichte Flysch'])
            elif (51 in temp) & (29 in temp):
                et_sets.append([int(i), 10, 'Fichte Tirol, Kalk niedrig'])
            elif (52 in temp) & (29 in temp):
                et_sets.append([int(i), 11, 'Fichte Tirol, Kalk mittel'])
            elif (53 in temp) & (29 in temp):
                et_sets.append([int(i), 12, 'Fichte Tirol, Silikat mittel'])
            elif (54 in temp) & (29 in temp):
                et_sets.append([int(i), 13, 'Fichte Tirol, Silikat hoch'])

        et_sets = pd.DataFrame.from_records(et_sets, columns=['BKL', 'Ertragstafelset', 'Bezeichnung'])

        return(et_sets)


    def fuc_tbl_et_ba(self, ets_nr, allba=True):

        # filter data
        et_ba = self.data_et[self.data_et['Ertragstafelsetnr']==ets_nr][['Ertragstafelsetnr', 'Baumart', 'Baumart Name', 'Ertragstafelname', 'Substitutionsprozent']]

        # filter just most important ones
        if allba == False:
            et_ba = et_ba[et_ba['Baumart'].isin(self.bag)]

        # sort table
        et_ba['sort'] = et_ba['Baumart'].map(self.dic.dic_ba_order)
        et_ba = et_ba.sort_values(by='sort')
        et_ba = et_ba.iloc[:,:-1]

        return(et_ba)
