################################################################################
#
#   fuc_tbl_hiebsatzbilanz (to, filterx)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

class OBFKlima(object):

    def __init__(self, data_temperature, data_precipitation, data_snow):

        self.data_temperature = data_temperature
        self.data_precipitation = data_precipitation
        self.data_snow = data_snow

    def printit(self):
        print(self.data)
        return(self.data)


    ### 6.1.1 Einschlagsübersicht

    def fuc_tbl_climate(self, stations):

        data = self.data_temperature.iloc[0:3]
        data = data.append(self.data_temperature.iloc[4:5])
        data = data.append(self.data_precipitation.iloc[4:5])
        data = data.append(self.data_precipitation.iloc[17:18])
        data = data.append(self.data_snow.iloc[31:33])

        table=[]
        for i in stations:
            table.append(data[i])
        table = pd.concat(table, axis=1)

        table = table.T
        table = table.reset_index()
        table.columns = ['Station', 'Längengrad [°]', 'Breitengrad [°]', 'Seehöhe [m]', 'Jahresmitteltemperatur [°C]', 'Mittlerer Jahresniederschlag [mm]', 'max. Tagesniederschlag [mm]', 'Mittlere Anzahl der Tage mit min. 1 cm Schneedeckenhöhe [cm]', 'Mittlere Anzahl der Tage mit min. 20 cm Schneedeckenhöhe [cm]']

        return(table)


    def fuc_plt_climate(self, stations, cat):

        if cat == 'temp':
            data = self.data_temperature.iloc[5:17]
            filterx = 'Monatsmitteltemperatur [°C]'
            titlex = 'Monatsmitteltemperatur'
        elif cat == 'prec_month':
            data = self.data_precipitation.iloc[5:17]
            filterx = 'Mittlerer Monatsniederschlag [mm]'
            titlex = 'Mittlerer Monatsniederschlag'
        elif cat == 'prec_day':
            data = self.data_precipitation.iloc[18:30]
            filterx = 'max. monatl. Tagesniederschlag [mm]'
            titlex = 'max. monatl. Tagesniederschlag'

        table=[]
        for i in stations:
            table.append(data[i])
        table = pd.concat(table, axis=1)

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table.plot(kind='line', ax=ax, stacked=False, alpha=0.9, title=titlex)
        ax.set_ylabel(filterx)
        ax.set_xlabel("Monat")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        #ax.legend(loc='center left', bbox_to_anchor=(0.97, 0.47), frameon=False)
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=3, borderaxespad=0., frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()

    def fuc_plt_climate_snow(self, stations):

        data = self.data_snow
        data = data.iloc[31:33]

        table = []
        for i in stations:
            table.append(data[i])
        table = pd.concat(table, axis=1)

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table.T.plot(kind='bar', ax=ax, stacked=False, alpha=0.9, title='Mittlere Anzahl der Tage mit min. 1 cm / 10 cm Schneedeckenhöhe')
        ax.set_ylabel("Tage")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=0)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='center left', bbox_to_anchor=(0.97, 0.47), frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()
