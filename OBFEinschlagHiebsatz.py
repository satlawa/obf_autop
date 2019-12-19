################################################################################
#
#   plot_es_hs(filtery, filterx)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from OBFDictionary import OBFDictionary

class OBFEinschlagHiebsatz(object):

    def __init__(self, data_alter, data_neig, data_seeh, data_uz):

        self.data_alter = data_alter
        self.data_neig = data_neig
        self.data_seeh = data_seeh
        self.data_uz = data_uz
        self.dic = OBFDictionary()

    def set_fr(self, fr):

        self.data_alter = self.data_alter[self.data_alter['Forstrevier'].isin(fr)]
        self.data_neig = self.data_neig[self.data_neig['Forstrevier'].isin(fr)]
        self.data_seeh = self.data_seeh[self.data_seeh['Forstrevier'].isin(fr)]
        self.data_uz = self.data_uz[self.data_uz['Forstrevier'].isin(fr)]
        # array with


    def printit(self):
        print(self.data)
        return(self.data)

    ###=========================================================================
    ###   6.1.2 Vergleich Einschlag zu Hiebsatz
    ###=========================================================================

    def plot_es_hs(self, filtery, filterx):
        '''
            filterz = 'alter', 'neigung', 'seehoehe', 'umtriebszeit'
            filtery = ['Altersgruppe 1. Schicht', 'Neigungsgruppe (%)', 'Seehöhe', 'Umtriebszeit']
            filterx = ['Endnutzung', 'Vornutzung']
        '''

        if filtery == 'Altersgruppe 1. Schicht':
            data = self.data_alter
        elif filtery == 'Neigungsgruppe (%)':
            data = self.data_neig
        elif filtery == 'Seehöhe':
            data = self.data_seeh
        elif filtery == 'Umtriebszeit':
            data = self.data_uz

        data = data.fillna(0)

        # copy data
        en_filter = data[[filtery, 'Maßnahmengruppe', 'Tax-Menge', 'Abmassmenge']].copy()
        #print(en_filter)
        #print('xxxxxx')

        # filter
        en_filter = en_filter[(en_filter['Maßnahmengruppe'] == filterx) & (en_filter[filtery] != 'Ergebnis') & (en_filter[filtery] != '#')]
        #en_filter.reset_index(level=1, inplace=True)
        #en_filter.reset_index(level=0, inplace=True)
        en_filter[filtery] = en_filter[filtery].astype(int)

        en_filter = en_filter.groupby([filtery]).sum()
        #2en_filter = en_filter.groupby(['Maßnahmengruppe', filtery]).sum()
        #print(en_filter)
        #print('xxxxxx')
        #en_filter = en_filter.groupby(['Maßnahmengruppe', filtery]).sum()
        #2en_filter.reset_index(level=1, inplace=True)
        #4en_filter.reset_index(level=0, inplace=True)

        #2en_filter[filtery] = en_filter[filtery].astype(int)
        #4en_filter = en_filter.sort_values(by=[filtery])

        # set first column as index
        #4en = en_filter.set_index(filtery)
        # filter VN or EN
        #3en = en_filter[en_filter['Maßnahmengruppe'] == filterx]
        # slice array into needed form
        #en = en.iloc[:-1,:]
        en = en_filter[['Tax-Menge', 'Abmassmenge']]
        #en.Index.astype(str)
        #print('xxxxxx')

        if filtery == 'Umtriebszeit':
            plot_kind = 'bar'
        else:
            plot_kind = 'line'

        fig = plt.figure(figsize=(6,3.2))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        en.plot(kind=plot_kind, ax=ax, linewidth = 2, alpha=0.9, rot=45, title=self.dic.dic_es_hs_head[filtery] + ' - ' + filterx)

        ax.set_ylabel("Masse [efm]")
        ax.set_xlabel(self.dic.dic_es_hs_unit[filtery])
        ax.set_ylim(ymin=0)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=2, borderaxespad=0., frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()
