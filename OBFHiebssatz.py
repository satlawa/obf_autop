################################################################################
#
#   fuc_tbl_hs_fest (cat)
#   fuc_tbl_hs_fest_bkl(fr)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

class OBFHiebssatz(object):

    def __init__(self, data):

        self.data = data

        # fill nan with 0
        self.data = self.data.fillna(0)

        # filter relavant data for Hiebsatz
        self.data = self.data[self.data['Hiebssatz Kennzahlty'] == 'JHS']
        self.data = self.data.fillna(0)

    def printit(self):
        print(self.data)
        return(self.data)

    ###=========================================================================
    ###   6.2.3 Festgestezter HS
    ###=========================================================================

    def fuc_tbl_hs_fest(self, cat):

        #x = self.data[[cat, 'Bewirtschaftungsform', 'EN Laubholz', 'EN Nadelholz', 'VN Laubholz', 'VN Nadelholz']]
        x = self.data[[cat, 'Bewirtschaftungsform', 'ENLH', 'ENNH', 'VNLH', 'VNNH']]
        y = x.groupby([cat,'Bewirtschaftungsform']).sum()

        # mutiply all numbers with 10 -> dezennaler HS
        y = y * 10

        y.reset_index(level=1, inplace=True)
        y.reset_index(level=0, inplace=True)
        u = y[cat].unique()

        ### function

        # make tamplet
        a = y[(y[cat] == u[0]) & (y['Bewirtschaftungsform'] == y['Bewirtschaftungsform'].unique()[0])].T
        a[:] = 0

        # create dataframe
        fin = a.copy()

        # loop in cat
        for i in u:
            # filter cat
            y2 = y[y[cat] == i]
            b = y2['Bewirtschaftungsform'].unique()

            # loop in Bewirtschaftungsform
            for j in ['W', 'S']:
                # find if W or S is in Bewirtschaftungsform
                if j in b:
                    y3 = y2[y2['Bewirtschaftungsform'] == j].T
                    y3.iloc[1,:] = j + 'W'
                    fin = pd.concat([fin, y3], axis=1)
                else:
                    y3 = a
                    y3.iloc[0,:] = i
                    y3.iloc[1,:] = j + 'W'
                    fin = pd.concat([fin, a], axis=1)
            # calc sum
            y3 = fin.iloc[:,-2] + fin.iloc[:,-1]
            y3.iloc[0] = int(i)
            y3.iloc[1] = 'Ges.'
            fin = pd.concat([fin, y3], axis=1)

        # delete first row and transponse
        fin = fin.iloc[:,1:].T

        # calc sum of TO
        sum_fin = fin.groupby(['Bewirtschaftungsform']).sum()
        # sort
        sequence = ['WW', 'SW', 'Ges.']
        sum_fin = sum_fin.reindex(index=sequence)
        sum_fin.reset_index(level=0, inplace=True)
        sequence = [cat, 'Bewirtschaftungsform', 'ENLH', 'ENNH', 'VNLH', 'VNNH']
        sum_fin = sum_fin.reindex(columns=sequence)
        sum_fin[cat] = 'TO'

        fin = pd.concat([fin, sum_fin], axis=0)
        fin = fin.reset_index(drop=True)

        y = fin

        if cat == 'Forstrevier':
            name = 'FR'
        elif cat == 'Umtriebszeit':
            name = 'UZ'

        #z = pd.concat([y.iloc[:,0:2], y.iloc[:,2:4], y.iloc[:,2] + y.iloc[:,3], y.iloc[:,4:], y.iloc[:,4] + y.iloc[:,5], y.iloc[:,2] + y.iloc[:,4], y.iloc[:,3] + y.iloc[:,5], y.iloc[:,2] + y.iloc[:,3] + y.iloc[:,4] + y.iloc[:,5]], axis=1)
        z = pd.concat([y.iloc[:,0:2], y.iloc[:,2:4].astype(int), y.iloc[:,2].astype(int) + y.iloc[:,3].astype(int), y.iloc[:,4:].astype(int), y.iloc[:,4].astype(int) + y.iloc[:,5].astype(int), y.iloc[:,2].astype(int) + y.iloc[:,4].astype(int), y.iloc[:,3].astype(int) + y.iloc[:,5].astype(int), y.iloc[:,2].astype(int) + y.iloc[:,3].astype(int) + y.iloc[:,4].astype(int) + y.iloc[:,5].astype(int)], axis=1)

        z.columns = [name, 'WW SW', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.']

        # mutiply all numbers with 10 -> dezennaler HS
        #z.ix[:,~np.in1d(z.dtypes,['object'])] *= 10

        return z


    ###=========================================================================
    ###   6.3.3 Festgestezter HS nach BKL
    ###=========================================================================

    def fuc_tbl_hs_fest_bkl(self, fr):

        x = self.data[self.data['Forstrevier']==fr]
        y = x[['Betriebsklasse', 'Umtriebszeit', 'Bewirtschaftungsform', 'Ertragssituation', 'ENLH', 'ENNH', 'VNLH', 'VNNH', 'Summe']]

        y.loc[:,('ENLH', 'ENNH', 'VNLH', 'VNNH', 'Summe')] = y.loc[:,('ENLH', 'ENNH', 'VNLH', 'VNNH', 'Summe')] * 10

        y.loc[:,'Bewirtschaftungsform'] = y.loc[:,'Bewirtschaftungsform'].replace(['S', 'SS'], 'SW')
        y.loc[:,'Bewirtschaftungsform'] = y.loc[:,'Bewirtschaftungsform'].replace(['W'], 'WW')

        #z = pd.concat([y.iloc[:,0:4], y.iloc[:,4:6], y.iloc[:,4] + y.iloc[:,5], y.iloc[:,6:8], y.iloc[:,6] + y.iloc[:,7], y.iloc[:,4] + y.iloc[:,6], y.iloc[:,5] + y.iloc[:,7], y.iloc[:,4] + y.iloc[:,5] + y.iloc[:,6] + y.iloc[:,7]], axis=1)
        z = pd.concat([y.iloc[:,0:4], y.iloc[:,4:6].astype(int), y.iloc[:,4].astype(int) + y.iloc[:,5].astype(int), y.iloc[:,6:8].astype(int), y.iloc[:,6].astype(int) + y.iloc[:,7].astype(int), y.iloc[:,4].astype(int) + y.iloc[:,6].astype(int), y.iloc[:,5].astype(int) + y.iloc[:,7].astype(int), y.iloc[:,4].astype(int) + y.iloc[:,5].astype(int) + y.iloc[:,6].astype(int) + y.iloc[:,7].astype(int)], axis=1)

        z.columns = ['BKL', 'UZ', 'WW SW', 'Ert.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.']
        z = z.sort_values(by=['BKL', 'UZ', 'WW SW', 'Ert.'], ascending=[True, True, False, False])

        z[['BKL', 'UZ']] = z[['BKL', 'UZ']].astype(int)

        z.loc['total'] = z.sum()
        z.iloc[-1,0:4] = ''
        z.iloc[-1,0] = 'Ges.'

        # mutiply all numbers with 10 -> dezennaler HS
        #z.ix[:,~np.in1d(z.dtypes,['object'])] *= 10

        return z
