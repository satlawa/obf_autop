################################################################################
#
#   fuc_tbl_hiebsatzbilanz (to, filterx)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

class OBFEinschlag(object):

    def __init__(self, data):

        self.data = data

        # fill nan with 0
        self.data = self.data.fillna(0)


    def printit(self):
        print(self.data)
        return(self.data)


    ### 6.1.1 Einschlagsübersicht

    def fuc_tbl_hiebsatzbilanz(self, to, fr, filterx):
        '''
            self.data = input self.data (dataframe)
            to = Teiloperat (int)
            filterx = Vornutzung / Endnutzung ('EN'|'VN')
        '''
        data_filter = self.data[self.data['Teiloperats-ID'].isin(to)]

        data_filter = data_filter[data_filter['Forstrevier'].isin(fr)]

        es = data_filter[data_filter['Hiebssatz Kennzahlty'] == 'JES']
        hs = data_filter[data_filter['Hiebssatz Kennzahlty'] == 'JHS']

        es_jahre = es.groupby('Jahr')
        hs_jahre = hs.groupby('Jahr')

        es = es_jahre.sum()*(-1)
        hs = hs_jahre.sum()
        bilz = (es-hs).cumsum()*(-1)


        ## make table

        # Hiebsatz
        table = hs.loc[:,[filterx[0] + 'LH']]
        table.columns = ['HS_LH_EN']
        table['HS_NH_EN'] = hs.loc[:,[filterx[0] + 'NH']]
        table['HS_EN'] = table['HS_LH_EN'] + table['HS_NH_EN']

        # Einschlag
        table['ES_LH_EN'] = es.loc[:,[filterx[0] + 'LH']]
        table['ES_NH_EN'] = es.loc[:,[filterx[0] + 'NH']]
        table['ES_EN'] = table['ES_LH_EN'] + table['ES_NH_EN']

        # bilanzierter Hiebsatz
        table['bil_LH_EN'] = bilz.loc[:,[filterx[0] + 'LH']]
        table['bil_NH_EN'] = bilz.loc[:,[filterx[0] + 'NH']]
        table['bil_EN'] = table['bil_LH_EN'] + table['bil_NH_EN']


        ## plot figure

        # prepare for plot ALL
        tableX = table.loc[:,['ES_EN']]
        tableX['hs'] = table.loc[:,['HS_EN']]
        tableX['bilz'] = table.loc[:,['bil_EN']]
        tableX.columns = ['Einschlag', 'Hiebsatz', 'Bilanz']

        # decide the location of the legend
        up_down = table.max().max() + table.min().min()

        # plot

        fig = plt.figure()
        ax = fig.add_subplot(111)
        ax.set_title('Hiebsatzbilanz - ' + filterx[1])
        ax.plot(tableX.index, tableX['Einschlag'], label="Einschlag")
        ax.plot(tableX.index, tableX['Hiebsatz'], label="Hiebsatz")
        ax.plot(tableX.index, tableX['Bilanz'], label="Bilanz")
        ax.fill_between(tableX.index, 0, tableX['Bilanz'], facecolor="none", edgecolor='green', hatch='/', alpha=0.5)
        #ax.fill_between(nutzungX.index, 0, nutzungX['bilz'], color='green', edgecolor='green', hatch='/', alpha=0.25)
        #ax1.bar(range(1, 5), range(1, 5), color='none', edgecolor='red', hatch="/", lw=1., zorder = 0)
        if up_down < 0:
            ax.legend(loc='lower left', frameon=False)
        else:
            ax.legend(loc='upper left', frameon=False)
        #ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

        #ax.spines['left'].set_position('center')
        ax.spines['right'].set_color('none')
        ax.spines['bottom'].set_position(('data',0))
        ax.spines['top'].set_color('none')
        ax.spines['left'].set_smart_bounds(True)
        ax.spines['bottom'].set_smart_bounds(True)
        ax.xaxis.set_ticks_position('bottom')
        ax.yaxis.set_ticks_position('left')
        ax.set(xlabel='Jahr', ylabel='Einschlag [efm]')

        for label in ax.get_xticklabels() + ax.get_yticklabels():
            label.set_fontsize(8)
            label.set_bbox(dict(facecolor='white', edgecolor='None', alpha=0.1 ))

        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        # table

        table.reset_index(level=0, inplace=True)
        table = round(table,0)
        table = table.astype(int)
        table.columns = ['Jahr', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.']

        return (table)
