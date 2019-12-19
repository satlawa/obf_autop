################################################################################
#
#   add_missing_data()
#   fuc_tbl_spi(fb, fr)
#   fuc_tbl_spi_v_ha(fb, fr)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl

from OBFDictionary import OBFDictionary

class OBFSpi(object):

    def __init__(self, data):

        self.data = data
        self.add_missing_data()
        self.dic = OBFDictionary()

    def printit(self):
        print(self.data)
        return(self.data)

    #---------------------
    #   missing data
    #---------------------

    def add_missing_data(self):
        ## add missing data

        # fill nan with 0
        self.data = self.data.fillna(0)

        # count the trees per point
        data_filter = self.data[((self.data['Soz.Stell.']!=5) | ((self.data['Soz.Stell.']==5) & (self.data['Alterskl.']=='THP'))) & (self.data['BI-Status']<7) & (self.data['BI-Status']!=3)]
        anz_pro_punkt = data_filter.groupby('Raster-Pkt').count()
        dic_anz = anz_pro_punkt['Jahr'].to_dict()

        # add count of trees
        self.data['anz_baum'] = self.data['Raster-Pkt']
        self.data['anz_baum'].replace(dic_anz, inplace=True)
        # add rep flaeche
        self.data['rep_fl'] = self.data['Flaefakt'] / self.data['anz_baum']
        # delete probebly
        #self.data['rep_v'] = self.data['   Vfm pro hA'] * self.data['rep_fl']


    def fuc_tbl_spi(self, fb, fr):

        # filter the right FB and FR
        data_filter = self.data[(self.data['FB'] == fb) & (self.data['FR'].isin(fr))].copy()
        # filter gone trees
        data_filter = data_filter[(data_filter['BI-Status']<7) & (data_filter['BI-Status']!=3)]
        # filter dead trees
        data_filter = data_filter[(data_filter['Soz.Stell.']!=5) | ((data_filter['Soz.Stell.']==5) & (data_filter['Alterskl.']=='THP'))].copy()

        # add BL to ALK I
        data_filter['Alterskl.'] = data_filter['Alterskl.'].replace('uK','I')
        data_filter['Alterskl.'] = data_filter['Alterskl.'].replace('THP','I')

        # create pivot table - Vorrat
        spi_v = pd.pivot_table(data_filter, index=['FR'], columns=['Alterskl.'], values=['Vfm vor Ort'], aggfunc=np.sum, fill_value=0, margins=True)
        spi_v.columns = spi_v.columns.droplevel()

        # create pivot table - Flaeche
        spi_fl = pd.pivot_table(data_filter, index=['FR'], columns=['Alterskl.'], values=['rep_fl'], aggfunc=np.sum, fill_value=0, margins=True)
        spi_fl.columns = spi_fl.columns.droplevel()

        # create pivot table - Vorrat/Flaeche
        spi_v_ha = spi_v/spi_fl

        # round numbers
        spi_v = spi_v.round(0)
        spi_fl = spi_fl.round(0)
        spi_v_ha = spi_v_ha.round(0)

        temp_fr = spi_v.index.values
        fr = temp_fr[:-1].tolist()
        fr = [temp_fr[-1]] + fr

        spi_v.reset_index(level=0, inplace=True)
        spi_fl.reset_index(level=0, inplace=True)
        spi_v_ha.reset_index(level=0, inplace=True)

        temp_spi = []

        # loop over fr
        for i in fr:
            x = spi_v[spi_v['FR']==i]
            y = spi_fl[spi_fl['FR']==i]
            z = spi_v_ha[spi_v_ha['FR']==i]

            xyz = pd.concat([x,y,z])
            xyz.iloc[:,1:] = xyz.iloc[:,1:].astype(int)
            xyz['Einheit'] = ['Vorrat [Vfm]', 'Fläche [ha]' , 'Vorrat/ Fläche [Vfm/ha]']

            xyz = xyz[['Einheit', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'All']]
            xyz.rename(columns={'All':'Ges.', 'VIII':'VIII+'}, inplace=True)
            temp_spi.append(xyz)

        return temp_spi


    def fuc_tbl_spi_v_ha(self, fb, fr):
        # filter dead trees
        data_filter = self.data[(self.data['Soz.Stell.']!=5) & (self.data['BI-Status']<7) & (self.data['BI-Status']!=3)].copy()
        data_filter = data_filter[(data_filter['FB'] == fb) & (data_filter['FR'].isin(fr))]

        spi_v = pd.pivot_table(data_filter, index=['FR'], columns=['Alterskl.'], values=['Vfm vor Ort'], aggfunc=np.sum, fill_value=0, margins=True)
        spi_v.columns = spi_v.columns.droplevel()

        spi_fl = pd.pivot_table(data_filter, index=['FR'], columns=['Alterskl.'], values=['rep_fl'], aggfunc=np.sum, fill_value=0, margins=True)
        spi_fl.columns = spi_fl.columns.droplevel()

        spi_v_ha = spi_v/spi_fl

        return spi_v_ha


    def fuc_tbl_spi_ss(self, to):

        # define categories
        kat_all = np.array(['10-14', '15-19', '20-24', '25-29', '30-34', '35-39', '40-44',
                            '45-49', '50-54', '55-59', '60-64', '65-69', '>70'])

        # filter data
        data_filter = self.data[self.data['TOID']==to]
        data_filter = data_filter.loc[data_filter['Soz.Stell.']!=5]
        data_filter = data_filter.loc[data_filter['Soz.Stell.']!=0]
        data_filter = data_filter.loc[(data_filter['BI-Status']<7) & (data_filter['BI-Status']!=3)]

        # filter data with SS
        data_ss = data_filter[data_filter['Schäl']!=1]

        # group by categories
        data_gr_all = data_filter.groupby(by=['Kl BHD'], as_index=True).sum()
        data_gr_ss = data_ss.groupby(by=['Kl BHD'], as_index=True).sum()

        # filter important features (attributes)
        table_all = data_gr_all[['Stz vor Ort', 'Vfm vor Ort']].copy()
        table_ss = data_gr_ss[['Stz vor Ort', 'Vfm vor Ort']].copy()

        kategorie = table_ss.index.values
        # find missing categories
        kat_miss = list(set(kat_all).difference(kategorie.tolist()))
        # add missing categories
        if kat_miss:
            for kat_m in kat_miss:
                table_ss.loc[kat_m] = [0,0]

        # add row with sum
        table_all.loc['Ges.'] = table_all.sum()
        table_ss.loc['Ges.'] = table_ss.sum()

        # calculate percentage
        table_per = table_ss / table_all * 100

        # concatonate tables
        table = pd.concat([table_all['Stz vor Ort'], table_ss['Stz vor Ort'], table_per['Stz vor Ort'], table_all['Vfm vor Ort'], table_ss['Vfm vor Ort'], table_per['Vfm vor Ort']], axis=1, sort=True)

        # round, set to int
        table = round(table,0).astype(int).reset_index(level=0)
        table.columns = ['BHD [cm]', '[Stz vor Ort]', '[Stz vor Ort]', '[%]', '[Vfm vor Ort]', '[Vfm vor Ort]', '[%]']
        table.rename(columns={'Kl BHD': 'BHD [cm]'}, inplace=True)

        return(table)


    def fuc_tbl_plt_spi_schaeden(self, to, typ, title, values):
        '''
        data
        typ     ('Schad.Urs.' 'Kronenverl', 'SchaftQual')
        title   ('Schaftschäden')
        values  ('Vfm vor Ort')
        '''

        # get dictionary
        dictonary = {'Schad.Urs.': self.dic.dic_schael, 'Kronenverl' : self.dic.dic_kronv, 'SchaftQual' : self.dic.dic_qual}
        dictonary_plot = {'Schad.Urs.': self.dic.dic_schael_plot, 'Kronenverl' : self.dic.dic_kronv, 'SchaftQual' : self.dic.dic_qual}

        # filter data
        data_filter = self.data[self.data['TOID']==to]
        data_filter = data_filter.loc[data_filter['Soz.Stell.']!=5]
        data_filter = data_filter.loc[data_filter['Soz.Stell.']!=0]
        data_filter = data_filter.loc[(data_filter['BI-Status']<7) & (data_filter['BI-Status']!=3)]

        # create table
        table = pd.pivot_table(data_filter, values=values, index=[typ], columns=['Alterskl.'], aggfunc=np.sum, margins=True)

        # rename column
        table.rename(columns={'All': 'Ges.', 'VIII':'VIII+'}, inplace=True)

        # sort table & fill missing
        table = pd.DataFrame(table, columns=sorted(self.dic.dic_akl_spi_order, key=self.dic.dic_akl_spi_order.get))

        '''
        # define categories
        kat_akl = np.array(['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'All'])
        # get categories
        kategorie = table.columns.values
        # find missing categories
        kat_miss = list(set(kat_akl).difference(kategorie.tolist()))
        # add missing categories
        if kat_miss:
            for kat_m in kat_miss:
                table[kat_m] = np.zeros(table.shape[0])
        '''

        # fill nan
        table.fillna(0, inplace=True)

        # round & convert to int
        table = round(table).astype(int).reset_index(level=0)


        ## plot

        # make copy for plot
        table_per = table.copy()

        # replace names
        table_per[typ].replace(dictonary_plot[typ], inplace=True)
        table_per.set_index(typ, inplace=True)
        table_per = table_per.iloc[:-1,:] / table_per.iloc[-1,:] * 100

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)

        # plot
        table_per.T.plot(kind='bar', stacked=True, ax=ax, edgecolor='black', width = 1, alpha=1, title=title)
        #edgecolor='white', (possible)

        # Remove plot frame
        ax.set_frame_on(False)
        ax.set_ylabel("Prozent")
        ax.set_xlabel("Altersklassen")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)
        # add legend
        #ax.legend(loc='center left', bbox_to_anchor=(0.97, 0.47), frameon=False)

        # add legend
        if (typ == 'SchaftQual'):
            ncol=4
        else:
            ncol=3

        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=ncol, borderaxespad=0., frameon=False)

        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()

        #table
        # replace names
        table[typ].replace(dictonary[typ], inplace=True)
        if title == 'Kronenverlichtung':
            title == 'Kronen-verlichtung'
        # rename column
        table.rename(columns={typ: title}, inplace=True)

        return(table)
