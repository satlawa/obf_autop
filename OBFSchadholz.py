################################################################################
#
#   fuc_tbl_schadholz (kindx, columnx, filterx)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

class OBFSchadholz(object):

    def __init__(self, data):

        self.data = data

        # fill nan with 0
        self.data = self.data.fillna(0)

        # filter data

        self.data = self.data[(self.data['Nutzungsart planmäßig'] == 'zufällig') & (self.data['Nutzungsart'] != 'Ergebnis') & (self.data['Nutzungsart'] != 'ZE') &(self.data['Nutzungsart'] != 'ZV')].copy()

    def set_fr(self, fr):

        self.data = self.data[self.data['Forstrevier'].isin(fr)]

    def printit(self):
        print(self.data)
        return(self.data)

    ###=========================================================================
    ###   6.1.3 Schadholz
    ###=========================================================================

    def fuc_tbl_schadholz(self, columnx, filterx, to):
        '''
            columnx = ['Nutzungsart', 'Forstrevier']
            filterx = ['Endnutzung', 'Vornutzung']
            to = '1341'
        '''

        ## prepare data

        # filter data
        data_mag = self.data[self.data['Maßnahmengruppe'] == filterx]

        # make pivot table
        table_zufaellige = pd.pivot_table(data_mag, index=['Abmaßjahr'], columns=[columnx],
                                                 values=['Abmassmenge'], aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table_zufaellige.columns = table_zufaellige.columns.droplevel()

        # delete 'All' sum in index
        table_zufaellige_plot = table_zufaellige[table_zufaellige.index != 'All']

        # delete 'All' sum in columns
        table_zufaellige_plot = table_zufaellige_plot.drop('All', axis=1)

        # change index type to int
        table_zufaellige_plot.index = table_zufaellige_plot.index.astype(int)

        if columnx == 'Nutzungsart':
            plot_kind = 'area'
        else:
            plot_kind = 'line'

        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_zufaellige_plot.plot(kind=plot_kind, stacked=False, ax=ax, alpha=0.5, title='Schadholzmasse in der ' + filterx + ', ' + columnx)
        #color=['#FC0D1B', '#FDB32B', '#FFFD38', '#80DA29', '#108112']
        ax.set_ylabel("Masse [efm]")
        ax.set_xlabel("Jahr")
        #ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)
        ax.set_ylim(ymin=0)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=6, borderaxespad=0., frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        ## table

        # rename columns
        table_zufaellige.rename(columns={'Abmaßjahr':'Jahr', 'All':'Ges.'}, inplace=True)
        table_zufaellige.rename(index={'All':'Ges.'}, inplace=True)

        # get index inside data (for printing into docx)
        table_zufaellige = round(table_zufaellige,0)
        table_zufaellige = table_zufaellige.astype(int)
        table_zufaellige.reset_index(level=0, inplace=True)

        #table_wuech['Wuechsigkeit'] = ['sehr gering wüchsig', 'gering wüchsig', 'mittelwüchsig', 'wüchsig', 'sehr wüchsig', 'All']

        return table_zufaellige
