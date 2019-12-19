################################################################################
#
# x  fuc_tbl_flaechen (self)
# x  fuc_tbl_flaechen_sw (self)
# x  fuc_tbl_plt_flaeche_vorrat (self, code)
# x  fuc_tbl_plt_ba_anteile (self, code)
# x  fuc_tbl_mittel_ (self, cat)                                                 cat: 'Schichtalter', 'Ertragsklasse', 'BaumartenBestockgrad'
# x  fuc_tbl_natuerliche_grundlagen (self, cat)                                  cat: 'Seehöhe', 'Exposition', 'NeigGruppe'
# x  fuc_plt_natuerliche_grundlagen (table, cat, cat2, name)
# x  fuc_tbl_plt_standortseinheiten (self)
# x  fuc_tbl_plt_geologie_wuechsigkeit (self)
# x  fuc_tbl_plt_bonitaetsverlauf (self)
# x  fuc_tbl_mittel (self, cat)                                                  cat: 'Schichtalter', 'Ertragsklasse', 'BaumartenBestockgrad'
# x  fuc_tbl_plt_umtriebsgruppen_wuechsigkeit (self)
# x  fuc_tbl_plt_umtriebsgruppen_stoe (self)
# x  fuc_tbl_hs_waldbau (data,kind)                                              kind: 'Forstrevier', 'Umtriebszeit'
# x  fuc_tbl_hs_waldbau_bkl (data,fr)                                            fr: 6, 7
# x  fuc_tbl_nutzungszeitpunkt(indexx, filterx, valuesx)
#
################################################################################

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from OBFDictionary import OBFDictionary

class OBFFunc(object):

    def __init__(self, data):

        self.data = data

        self.data = self.data[self.data['GUID']!=0]

        # fill nan with 0
        #self.data = self.data.fillna(0)

        self.dic = OBFDictionary()

        self.dictionary()
        self.add_missing_data()

        #---------------------
        #   filter data
        #---------------------

        # Flaeche der WO (2.1)
        self.data_wo_fl = self.data[self.data['Schichtalter'] == 0]

        # Flaeche & Vorrat nach AKL (2.5 | 5.1)
        self.data_fl = self.data[self.data['Repr. Fläche Schicht'] != 0]
        self.data_V = self.data[self.data['Baumartenanteil'] != 0]
        self.data_V_UE = self.data[(self.data['Überh. Laubholz']!=0) | (self.data['Überh. Nadelhz.']!=0)][['Forstrevier', 'Abteilung', 'Unterabteil.', 'Umtriebszeit', 'Bewirtschaftungsform', 'Überh. Laubholz', 'Überh. Nadelhz.']]

        # Baumartenanteile (2.6 | 5.3)
        # filter relavant data for Baumartenanteile
        self.data_ba = self.data[self.data['Ertragstafelnummer'] != 0]
        # Blöße aus Baumarten (BA) & aus Altersklasse (AKL) filtern
        self.data_ba = self.data_ba[(self.data_ba['Baumart'] != 'BL') & (self.data_ba['TAX: Altersklasse'] != 'BL')]

        # Mittel (2.7 | 5.5)
        self.data_ekl = self.data[self.data['Ertragsklasse'] != 0]

        # Standort , Wuechsigkeit , Geologie
        self.data_stoe = self.data[(self.data['Best.-Schicht'] == 0) & (data['WE-Typ'] == 'WO')]

        # Natürliche Grundlagen (3.x)
        # set Maßnahmen to 0
        #data_stoe = data[(data['Vegetationstyp'] != 0) & (data['Schichtanteil'] == 0)].copy()
        #data_stoe = data_stoe[['Forstrevier', 'Abteilung', 'Unterabteil.', 'Teilfl.', 'Fläche in HA', 'Seehöhe', 'Exposition', 'Neigung', 'Standorteinheit', 'Vegetationstyp', 'Waldtyp', 'Wuechsigkeit', 'Geologie']]

        # Maßnahmen = filter 0
        self.data_massnahmen = self.data[self.data['Massnahmengruppe'] != 0]
        # Nutzungsmaßnahmen = filter not WP (Waldpflege)
        self.data_vn_en = self.data_massnahmen[self.data_massnahmen['Massnahmengruppe'] != 'WP']
        # Waldpflegemaßnahmen = filter WP (Waldpflege)
        #data_wp = data_massnahmen[data_massnahmen['Massnahmengruppe'] == 'WP']
        # Vornutzungsmaßnahmen = VN (Vornutzung)
        #data_vn = data_massnahmen[data_massnahmen['Massnahmengruppe'] == 'VN']


    def printit(self):
        print(self.data)
        return(self.data)

    def dictionary(self):

        # get Teiloperat
        self.dic.set_to(self.data['Teiloperats-ID'].unique()[0])
        # get Forstbetrieb
        self.dic.set_fb(self.data['Forstbetrieb'].unique()[0])
        # get Wuchsgebiet
        # XXXXXXXXXXXXXXX
        wg = self.data['Wuchsgebiet'].unique()
        wg = wg[~pd.isnull(wg)]
        wg = np.sort(wg)
        self.dic.set_wg(wg)

        # get Laufzeit
        #x = self.data['Beg. Laufzeit'].iloc[0].to_pydatetime()
        x = self.data['Beg. Laufzeit'].iloc[0][-4:]
        y = self.data['Ende Laufzeit'].iloc[0][-4:]
        self.dic.set_laufzeit_begin(x)
        self.dic.set_laufzeit_end(y)
        self.dic.set_laufzeit(x + '-' + y)

        x = np.sort(self.data['Forstrevier'].unique())
        self.dic.set_fr(x[x.nonzero()])
        x = np.sort(self.data['Umtriebszeit'].unique())
        self.dic.set_uz(x[x.nonzero()])
        x = np.sort(self.data['Nutzdringlichkeit'].unique())
        self.dic.set_dring(x[x.nonzero()])
        x[x.nonzero()]

        # set fr dict
        self.dic.set_dic_fr()

        # define colors
        self.colors = ['#3976AF', '#EF8536', '#529D3E', '#C53932', '#8D6AB8', '#DEAD3B', '#85584E', '#D57DBE', '#7F7F7F', '#57BBCC', '#4291F5', '#2B6419', '#E96B77', '#741E18', '#FEEC61']
        #colors = ['#3976AF', '#EF8536', '#529D3E', '#C53932', '#8D6AB8', '#DEAD3B', '#85584E', '#D57DBE', '#7F7F7F', '#A1C499', '#F3B770', '#B1D668', '#EC8677', '#FEEC61']
        #colors = ['#3976AF', '#EF8536', '#529D3E', '#C53932', '#741E18', '#4291F5', '#F7C652', '#62C554',  '#CD4D19', '#2B6419', '#C38D36', '#741E18',  '#DEAD3B', '#85584E', '#D57DBE', '#7F7F7F', '#A1C499', '#F3B770', '#B1D668', '#EC8677', '#FEEC61']
        # dr #741E18 dy #C38D36 dg #2B6419 r #ED6C63 y #F7C652 g #62C554 b #4291F5

    #---------------------
    #   missing data
    #---------------------

    def add_missing_data(self):
        ## add missing data

        # fill nan with 0
        self.data = self.data.fillna(0)

        # add column Wuechsigkeit
        self.data['Wuechsigkeit'] = self.data['Standorteinheit']
        self.data['Wuechsigkeit'].replace(self.dic.dic_stoe_wuech, inplace=True)

        # add column Geologie
        self.data['Geologie'] = self.data['Standorteinheit']
        self.data['Geologie'].replace(self.dic.dic_stoe_geo, inplace=True)

        # add column LHNH
        self.data['LHNH'] = self.data['Baumart']
        self.data['LHNH'].replace(self.dic.dic_ba_lh_nh, inplace=True)


        # the Baumart_group

        # add Baumart_group - 13 Baumarten everything else SL or SN
        x = self.data[self.data['Ertragstafelnummer'] != 0]

        # Blöße aus Baumarten (BA) & aus Altersklasse (AKL) filtern
        x = x[(x['Baumart'] != 'BL') & (x['TAX: Altersklasse'] != 'BL')]

        x = x[['Baumart','Repr. Fläche Baumart']]
        y = x.groupby(['Baumart'], as_index=False).sum()
        #y.reset_index(level=0, inplace=True)
        y = y[~y['Baumart'].isin(['SL','SN'])]
        y = y.sort_values(by=['Repr. Fläche Baumart'], ascending=False)

        # get 13 most used
        z_start = y.iloc[:13,0].unique()
        # get the rest
        z_end = y.iloc[13:,0].unique()

        # add column Baumart_group
        self.data['Baumart_group'] = self.data['Baumart'].copy()
        name_sl = ' SL: '
        name_sn = ' SN: '
        for i in z_end:
            if self.dic.dic_ba_lh_nh.get(i) == 'LH':
                self.data['Baumart_group'].replace({i:'SL'},inplace=True)
                name_sl = name_sl + i + ', '
            elif self.dic.dic_ba_lh_nh.get(i) == 'NH':
                self.data['Baumart_group'].replace({i:'SN'},inplace=True)
                name_sn = name_sn + i + ', '

        # set SL & SN in dic (class)
        name_sl = name_sl[:-2]
        name_sn = name_sn[:-2]
        name_sl = name_sl + '; '
        name_sn = name_sn + '; '
        self.dic.set_sl_sn(name_sl, name_sn)

        # categorize 'Neigung' into groups

        # create list with categories
        labels = list(range(0,200,10))
        # categorize data
        self.data['NeigGruppe'] = pd.cut(self.data['Neigung'], range(-5, 205, 10), right=False, labels=labels)
        # convert dtype from category into int
        self.data['NeigGruppe'] = self.data['NeigGruppe'].astype(int)

        # replace 'Wüchsigkeit' numbers with text
        #data_stoe['Wuechsigkeit'].replace(dic_num_wuech, inplace=True)


    #***************************************************************************


    ###=========================================================================
    ###   2.1 Allgemein - Flächentabelle (nach Revieren und TO)
    ###=========================================================================

    def fuc_tbl_flaechen(self):

        ## prepare data

        # get Abteilungsnummern
        abt = self.abteilungen_numbers('Forstrevier')

        # pivot table WO [ha]
        tableW = pd.pivot_table(self.data_wo_fl[self.data_wo_fl['Nebengrund Art']==0], index=['Forstrevier'],columns = ['Bewirtschaftungsform'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)
        # drop level
        tableW.columns = tableW.columns.droplevel()
        tableW = pd.DataFrame(tableW, columns=sorted({'W':0, 'S':1, 'All':2}, key={'W':0, 'S':1, 'All':2}.get))

        # pivot table NG [ha]
        tableNG = pd.pivot_table(self.data_wo_fl, index=['Forstrevier'],columns = ['Nebengrund Art'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)
        # drop level
        tableNG.columns = tableNG.columns.droplevel()
        tableNG = pd.DataFrame(tableNG, columns=sorted({0:0, 1:1, 2:2, 3:3, 4:4, 5:5, 6:6, 7:7, 8:8, 9:9, 'All':10}, key={0:0, 1:1, 2:2, 3:3, 4:4, 5:5, 6:6, 7:7, 8:8, 9:9, 'All':10}.get))

        # get difference between arrays
        diff = np.setdiff1d(np.array([0,1,2,3,4,5,6,7,8,9]), self.data_wo_fl['Nebengrund Art'].unique())

        # check if all NGs are
        if diff:
            # add missing data
            for i in diff:
                tableNG[i] = 0
            # sort columns
            sequence = [0,1,2,3,4,5,6,7,8,9,'All']
            tableNG = tableNG.reindex(columns=sequence)

        table = tableW
        table = pd.concat([abt, table], axis=1)
        # Nichtholzboden
        temp = tableNG.iloc[:,3] + tableNG.iloc[:,4]
        table = pd.concat([table, temp], axis=1)
        # Summe Waldfläche
        temp = table.iloc[:,3] + table.iloc[:,4]
        table = pd.concat([table, temp], axis=1)
        # Nebengründe produktiv
        temp = tableNG.iloc[:,5] + tableNG.iloc[:,6]
        # Nebengründe unproduktiv
        temp2 = tableNG.iloc[:,1] + tableNG.iloc[:,2] + tableNG.iloc[:,7] + tableNG.iloc[:,8] + tableNG.iloc[:,9]
        temp3 = temp + temp2
        table = pd.concat([table, temp, temp2, temp3, tableNG.iloc[:,10]], axis=1)

        # rename columns
        table.columns = ['Abteilungen', 'Wirtschafts-wald', 'Schutz-wald', 'Ges. Holzboden', 'Nichtholz-boden' , 'Ges.', 'produktiv', 'unproduktiv', 'Ges.', '[ha]']
        tableNG = tableNG.reindex(columns=['Abteilungen', 'Wirtschafts-wald', 'Schutz-wald', 'Ges. Holzboden', 'Nichtholz-boden' , 'Ges.', 'produktiv', 'unproduktiv', 'Ges.', '[ha]'])
        table.rename(index={'All': 'TO'}, inplace=True)
        # round values
        table = round(table,1)
        #table = table.astype(int)
        # get index inside data (for printing into docx)
        # sort
        table = pd.DataFrame(table, index=sorted(self.dic.dic_fr_order, key=self.dic.dic_fr_order.get))
        table.reset_index(level=0, inplace=True)
        table.rename(columns={'index': 'Forstrevier'}, inplace=True)

        return(table)


    ###=========================================================================
    ###   2.1 Schutzwald - Flächentabelle (nach Revieren und TO)
    ###=========================================================================

    def fuc_tbl_flaechen_sw(self):

        ## prepare data

        # kategorien
        kategorie = self.data_wo_fl['Schutzwaldkategorie'].unique()
        # filter
        kategorie = kategorie[np.where( kategorie != 0 )]
        # sort
        np.sort(kategorie)[::-1]
        # find missing kategorien
        kat_all = ['S', 'O', 'B']
        kat_miss = list(set(kat_all).difference(kategorie.tolist()))
        # make dic for sorting
        dic_kat_order = {"I":0, "A":1, "All":2}

        y=[]
        for i in kat_all:

            if i in kategorie.tolist():
                #print(i)
                # make table
                table = pd.pivot_table(self.data_wo_fl[(self.data_wo_fl['Bewirtschaftungsform']=='S') & (self.data_wo_fl['Schutzwaldkategorie']==i)], index=['Forstrevier'],columns = ['Ertragssituation'], \
                                       values=['Fläche in HA'], \
                                      aggfunc=np.sum, fill_value=0, margins=True)
                # drop level
                table.columns = table.columns.droplevel()
                # sort by dic
                table = pd.DataFrame(table, columns=sorted(dic_kat_order, key=dic_kat_order.get))
                y.append(table)
                #print(table)

            else:
                print('else')
                idx = pd.Index(self.data_wo_fl['Forstrevier'].unique().tolist() + ['All'])
                table = pd.DataFrame(data={'A': [0], 'I': [0], 'All': [0]})
                table = table.reindex(idx, fill_value=0)
                table = pd.DataFrame(table, columns=sorted(dic_kat_order, key=dic_kat_order.get))
                y.append(table)
                #print(table)

        # make table 'all'
        table = pd.pivot_table(self.data_wo_fl[self.data_wo_fl['Bewirtschaftungsform']=='S'], index=['Forstrevier'],columns = ['Ertragssituation'], \
                               values=['Fläche in HA'], \
                              aggfunc=np.sum, fill_value=0, margins=True)
        # drop level
        table.columns = table.columns.droplevel()

        # sort by dic
        table = pd.DataFrame(table, columns=sorted(dic_kat_order, key=dic_kat_order.get))
        y.append(table)

        #print(y)

        table = pd.concat([y[0], y[1], y[2], y[3]], axis=1)
        table = round(table,1)

# precent
        #table_pre = (table.div(table.iloc[:,-1], axis=0))*100
        #table_pre = round(table_pre,0).astype(int)

        #table = self.fuc_concat_tables(table, table_pre)

        # rename columns
        #table.columns = ['SS iE', 'SS iE %', 'SS aE', 'SS aE %' , 'SS', 'SS %', 'OS iE', 'OS iE %', \
        #'OS aE', 'OS aE %' , 'OS', 'OS %', 'SW iE', 'SW iE %', 'SW aE', 'SW aE %' , 'SW', 'SW %']

        table.columns = ['in Ertrag', 'außer Ertrag', 'Ges.', 'in Ertrag', \
        'außer Ertrag', 'Ges.', 'in Ertrag', 'außer Ertrag', 'Ges.', 'in Ertrag', 'außer Ertrag', 'Ges.']
        table.rename(index={'All': 'TO'}, inplace=True)

        # sort
        table = pd.DataFrame(table, index=sorted(self.dic.dic_fr_order, key=self.dic.dic_fr_order.get))

        table.index.name = 'Forstrevier'

        table = table.fillna(0)

        # get index inside data (for printing into docx)
        table.reset_index(level=0, inplace=True)

        return(table)


    ###=========================================================================
    ###   2.4 | 5.1 Flaeche & Vorrat nach AKL
    ###=========================================================================

    def fuc_tbl_plt_flaeche_vorrat(self, code):
        '''
        creats a table and a plot 'Flaeche & Vorrat nach AKL'
        input:  code [FR,UZ,Bew]
        output: table | plot
        '''

        ## prepare data
        data_fl_filter = self.filter_code(self.data_fl,code)
        data_V_filter = self.filter_code(self.data_V,code)
        data_V_UE_filter = self.filter_code(self.data_V_UE,code)
        name = self.name_code(code)

        # make table of Fläche pro Altersklasse
        table_fl = data_fl_filter.groupby('TAX: Altersklasse').sum()
        table_fl = table_fl['Repr. Fläche Schicht']

        # make table of Vorrat pro Altersklasse
        table_V = data_V_filter.groupby('TAX: Altersklasse').sum()
        table_V = table_V['Vorrat am Ort']

        # add volume of Überhälter
        ue = data_V_UE_filter[['Überh. Laubholz', 'Überh. Nadelhz.']].sum().sum()
        if ('VIII' in table_V.index) & (not(np.isnan(ue))):
            table_V['VIII'] = table_V['VIII'] + ue
        elif (not('VIII' in table_V.index)) & (not(np.isnan(ue))):
            table_V['VIII'] = ue

        table_fl.rename(index={'VIII': 'VIII+'}, inplace=True)
        table_V.rename(index={'VIII': 'VIII+'}, inplace=True)

        # make table with percent
        table_fl_p = table_fl.div(table_fl.sum())*100
        table_V_p = table_V.div(table_V.sum())*100
        table_ha = table_V.div(table_fl)

        table_all = pd.concat([table_fl, table_fl_p, table_V, table_V_p, table_ha], axis=1)

        ## get missing values
        '''kategorie = table_all.index.values
        kat_all = ['BL', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII+']
        kat_miss = list(set(kat_all).difference(kategorie.tolist()))

        if kat_miss:
            for kat_m in kat_miss:
                table_all.loc[kat_m] = [0,0,0,0,0]'''

        # rename columns
        table_all.columns = ['Fläche (Schicht) [ha]', 'Fläche (Schicht) [%]', 'Vorrat [Vfm]', 'Vorrat [%]', 'Vorrat/Fläche [Vfm/ha]']

        # sort table & fill in missing indices
        table_all = pd.DataFrame(table_all, index=sorted(self.dic.dic_akl_plot_order, key=self.dic.dic_akl_plot_order.get))

        # fill nan
        table_all.fillna(0,inplace=True)


        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot axis 1
        ax1 = fig.add_subplot(111, label="1")
        # Add a subplot axis 2
        ax2 = ax1.twinx()
        # plot
        table_vol = table_all.iloc[:,2] + 0.01
        plot0 = table_vol.plot(kind='line', ax=ax2, legend=None)
        plot1 = table_all.iloc[:,0].plot(kind='bar', ax=ax1, legend=None, title="Fläche und Vorrat nach AKL - " + name[0] + ' ' + name[2], alpha=0.9, edgecolor='gray', linewidth=0.5)
        #plot1 = table_all.iloc[:,0].plot(kind='bar', ax=ax1, legend=None, title="Fläche und Vorrat nach AKL", alpha=0.9, edgecolor='gray', linewidth=0.5)

        # define colors
        colors = ['#FFFFFF', '#FEEC61', '#FC5154', '#7BB45F', '#39A1D9', '#E7BD62', '#A3A184', '#9DA89F', '#84ABBE']
        # set colors
        for i in range(0,9):
            ax1.get_children()[i].set_color(colors[i])
            ax1.get_children()[i].set_edgecolor('gray')

        # set axes 1
        #ax1.set_frame_on(False)
        # Remove plot frame on the right and upper spines
        ax1.spines['top'].set_visible(False)
        # Only show ticks on the left, right and bottom spines
        ax1.yaxis.set_ticks_position('right')
        ax1.yaxis.set_ticks_position('left')
        ax1.xaxis.set_ticks_position('bottom')

        ax1.set_ylabel("Fläche [ha]")
        ax1.set_xlabel("Altersklassen")
        ax1.tick_params('y')
        ax1.set_xticklabels(ax1.xaxis.get_majorticklabels(), rotation=45)
        # set axes 2
        ax2.set_frame_on(False)
        #ax2.yaxis.tick_right()
        ax2.set_ylabel("Volumen [Vfm]")
        ax2.yaxis.set_label_position('right')
        ax2.tick_params('y')

        #fig.tight_layout()
        # align axes
        self.fuc_align_yaxis(ax1, 0, ax2, 0)

        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        ## table

        # transpose table
        table_all = table_all.T
        # add column 'Summe'
        table_all['Ges.'] = table_all.T.sum()
        # change the sum to average in the last row
        table_all.iloc[4,9] = table_all.iloc[2,9] / table_all.iloc[0,9]
        # round all values
        table_all = table_all.round(0)
        table_all.fillna(0,inplace=True)
        table_all = table_all.astype(int)
        # get index inside data (for printing into docx)
        table_all.reset_index(level=0, inplace=True)
        # rename columns
        table_all.rename(columns={'index': 'Einheit'}, inplace=True)

        return(table_all)


    ###=========================================================================
    ###   2.5 | 5.3 Baumartenanteile nach AKL
    ###=========================================================================

    def fuc_tbl_plt_ba_anteile(self,code):
        '''
        creats a table and a plot 'Baumartenanteile nach AKL'
        input:  code [FR,UZ,Bew]
        output: table | plot
        '''

        ## prepare data
        table_ba_filter = self.filter_code(self.data_ba,code)
        name = self.name_code(code)
        unique = self.data_ba['Baumart_group'].unique()

        # make pivot table
        table_ba_filter = pd.pivot_table(table_ba_filter, index=['Baumart_group'],columns = ['TAX: Altersklasse'], \
                           values=['Repr. Fläche Baumart'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        table_ba_filter.rename(columns={'VIII': 'VIII+'}, inplace=True)

        table_ba_pre = table_ba_filter.div(table_ba_filter.iloc[-1,:], axis=1 )*100

        # drop column level
        table_ba_pre.columns = table_ba_pre.columns.droplevel()

        # delete 'All' sum in index
        table_ba_pre = table_ba_pre[table_ba_pre.index != 'All']

        # sort table index (Baumarten)
        table_ba_pre = pd.DataFrame(table_ba_pre, index=sorted(self.dic.dic_ba_order, key=self.dic.dic_ba_order.get))
        # filter sorted table
        table_ba_pre = table_ba_pre.loc[table_ba_pre.index.isin(unique)]

        table_ba_pre.rename(columns={'Index': 'Altersklasse', 'All': 'Ges.'}, inplace=True)
        # sort table columns (AKL)
        table_ba_pre = pd.DataFrame(table_ba_pre, columns=sorted(self.dic.dic_akl_spi_order, key=self.dic.dic_akl_spi_order.get))

        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)

        # plot
        table_ba_pre.T.plot(kind='bar', stacked=True, ax=ax, edgecolor='black', width = 1, alpha=1, color=self.colors, title='Baumartenanteile im  - ' + name[0] + ' ' + name[2])
        #edgecolor='white', (possible)

        # Remove plot frame
        ax.set_frame_on(False)
        ax.set_ylabel("Prozent")
        ax.set_xlabel("Altersklassen")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)
        # add legend
        ax.legend(loc='center left', bbox_to_anchor=(0.97, 0.47), frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        ## table

        table_ba_pre = round(table_ba_pre, 1)

        # get index inside data (for printing into docx)
        table_ba_pre.reset_index(level=0, inplace=True)

        # rename columns
        table_ba_pre.rename(columns={'index': 'Baumart', 'All': 'Ges.'}, inplace=True)

        table_ba_pre = table_ba_pre.round(1)

        table_ba_pre = table_ba_pre.fillna(0)

        #table_ba_pre = table_ba_pre.astype(int)

        return(table_ba_pre)


    ###=========================================================================
    ###   2.6 Mittlere Ertragsklasse der Hauptbaumarten
    ###=========================================================================

    def fuc_tbl_mittel_(self,cat):
        '''
        cat: 'Schichtalter', 'Ertragsklasse', 'BaumartenBestockgrad'
        '''

        ## prepare data
        unique = self.data_ba['Baumart_group'].unique()

        # create dataframe mittlere EKL
        table_mittel = self.data_ekl.groupby('Baumart_group').apply(self.wavg,cat,'Repr. Fläche Baumart')
        col_name = ['All']

        # add every EKL to dataframe mittlere EKL
        for i in np.sort(self.data_ekl['Umtriebszeit'].unique()):
            data_i = self.data_ekl[self.data_ekl['Umtriebszeit']==i].groupby('Baumart_group').apply(self.wavg,cat,'Repr. Fläche Baumart')
            table_mittel = pd.concat([table_mittel, data_i], axis=1)
            col_name = col_name + [i]

        for i in np.sort(self.data_ekl['Bewirtschaftungsform'].unique()):
            data_i = self.data_ekl[self.data_ekl['Bewirtschaftungsform']==i].groupby('Baumart_group').apply(self.wavg,cat,'Repr. Fläche Baumart')
            table_mittel = pd.concat([table_mittel, data_i], axis=1)
            col_name = col_name + [i]

        # give columns their names
        table_mittel.columns = col_name

        # sort table
        table_mittel = pd.DataFrame(table_mittel, index=sorted(self.dic.dic_ba_order, key=self.dic.dic_ba_order.get))
        # filter sorted table
        table_mittel = table_mittel.loc[table_mittel.index.isin(unique)]

        # get index inside data (for printing into docx)
        table_mittel.reset_index(level=0, inplace=True)

        # sort columns
        sequence = ['index', 120, 140, 160, 200, 'W', 'S', 'All']
        table_mittel = table_mittel.reindex(columns=sequence)

        # rename columns
        table_mittel.rename(columns={'index':'Baumart', 'W': 'WW', 'S': 'SW', 'All': 'Ges.'}, inplace=True)

        table_mittel = table_mittel.round(1)

        table_mittel = table_mittel.fillna('-')

        #table_ba_pre = table_ba_pre.astype(int)

        return(table_mittel)


    ###=========================================================================
    ###   3 Natürliche Grundlagen (Seehöhe, Exposition, Neigung)
    ###=========================================================================

    def fuc_tbl_natuerliche_grundlagen(self, cat):
        '''
        cat: 'Seehöhe', 'Exposition', 'NeigGruppe'
        '''

        # make a pivot table out of data
        table = pd.pivot_table(self.data_stoe, index=[cat],columns = ['Forstrevier'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table.columns = table.columns.droplevel()

        table.rename(index={'All': 'Ges.'}, inplace=True)

        #if cat == 'Exposition':
        #    table = pd.DataFrame(table, index=sorted(self.dic.dic_exp_order, key=self.dic.dic_exp_order.get))
        #    table.fillna(0, inplace=True)
        #print(table)

        return table


    ###=========================================================================
    ###   3 Natürliche Grundlagen (Seehöhe, Exposition, Neigung)
    ###=========================================================================

    def fuc_plt_natuerliche_grundlagen(self,table,cat,cat2,name):
        '''
        table: 'table'
        cat: 'Seehöhe', 'Exposition', 'NeigGruppe'
        cat2: ''
        '''

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table[cat2].loc[table.index != 'Ges.'].plot(kind='bar', legend=None, figsize=(4.2, 2.5), title=name[1] + ' ' + name[3] , color=name[4])
        ax.set_xlabel(name[0])
        ax.set_ylabel("Fläche [ha]")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # save plot
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()

        # table
        tableX = table.copy()

        # get index inside data (for printing into docx)
        tableX.reset_index(level=0, inplace=True)

        # del index header (for printing into docx)
        del tableX.index.name

        # filter just needed columns
        tableX = tableX[[cat, cat2]]

        # create % column
        tableX['%'] = tableX[cat2] / tableX[cat2].loc[tableX[cat] != 'Ges.'].sum() * 100

        # rename columns
        tableX.columns = [name[0], '[ha]', '[%]']
        tableX['[ha]'] = tableX['[ha]'].round(1)
        tableX['[%]'] = tableX['[%]'].round(0)

        # change data type
        tableX.loc[:,'[%]'] = tableX.loc[:,'[%]'].astype(int)

        return tableX


    ###=========================================================================
    ###   3.5.1 Standortseinheiten im Teiloperat
    ###=========================================================================

    def fuc_tbl_plt_standortseinheiten(self):

        ## prepare data
        data = self.data_stoe

        # make pivot table
        table_stoe = pd.pivot_table(data, index=['Forstrevier'],columns = ['Standorteinheit'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table_stoe.columns = table_stoe.columns.droplevel()

        # rename columns
        table_stoe.rename(columns=self.dic.dic_stoe_int, inplace=True)
        table_stoe.rename(columns={'All':'Ges.'}, inplace=True)
        table_stoe.rename(index={'All':'Ges.'}, inplace=True)

        # calc percentage
        table_stoe_pre = round(table_stoe.div(table_stoe.iloc[:,-1], axis=0 ),3)*100

        # delete 'All' sum in index
        table_stoe_pre = table_stoe_pre.iloc[:,:-1]
        table_stoe_pre = table_stoe_pre.iloc[::-1]
        #table_stoe_pre = table_stoe_pre[table_stoe.index != 'All']
        #table_stoe_pre = table_stoe_pre.T


        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(11,3.2))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_stoe_pre.plot(kind='barh', stacked=True, ax=ax, width = 0.5, alpha=0.9, color=self.colors, title='Prozentuale Anteile der Standortseinheiten')
        ax.set_xlabel("Prozent [%]")
        #edgecolor='gray', (possible)
        #ax.set_ylabel("Prozent")
        #ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=10, borderaxespad=0., frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        ## table
        table_stoe.index.name = 'FR'
        # get index inside data (for printing into docx)
        table_stoe = round(table_stoe,1)
        #table_stoe.columns = ['sehr gering wüchsig', 'gering wüchsig', 'mittelwüchsig', 'wüchsig', 'sehr wüchsig', 'All']
        table_stoe.reset_index(level=0, inplace=True)

        return(table_stoe)


    ###=========================================================================
    ###   3.5.2 Substrat und Wüchsigkeit im Teiloperat
    ###=========================================================================

    def fuc_tbl_plt_geologie_wuechsigkeit(self, code):
        '''
        creats a table and a plot 'Geologie und Wuechsigkeit'
        input:  code [FR,UZ,Bew]
        output: table | plot
        '''

        ## prepare data
        data_stoe_filter = self.filter_code(self.data_stoe,code)
        name = self.name_code(code)

        # make pivot table
        table_wuech = pd.pivot_table(data_stoe_filter, index=['Wuechsigkeit'],columns = ['Geologie'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table_wuech.columns = table_wuech.columns.droplevel()

        # rename columns
        table_wuech.rename(index={1:'sehr gering wüchsig', 2:'gering wüchsig', 3:'mittelwüchsig', 4:'wüchsig' , 5:'sehr wüchsig', 'All': 'Ges.'}, inplace=True)
        table_wuech.rename(columns={'All':'Ges.'}, inplace=True)

        # calc percentage
        table_wuech_pre = round(table_wuech.div(table_wuech.iloc[-1,:], axis=1 ),3)*100

        # delete 'All' sum in index
        table_wuech_pre = table_wuech_pre[table_wuech_pre.index != 'Ges.']
        table_wuech_pre = table_wuech_pre.T


        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_wuech_pre.plot(kind='bar', stacked=True, ax=ax, width = 0.5, alpha=0.9, title='Wüchsigkeit nach Substrat im ' + name[1], color=['#FC0D1B', '#FDB32B', '#FFFD38', '#80DA29', '#108112'])
        #edgecolor='gray', (possible)
        ax.set_ylabel("Prozent [%]")
        ax.set_xlabel("Substrat")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.35), ncol=3, borderaxespad=0., frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        ## table

        # get index inside data (for printing into docx)
        table_wuech = table_wuech.T
        table_wuech = table_wuech.round(1)
        table_wuech.reset_index(level=0, inplace=True)
        # rename columns
        table_wuech.rename(columns={'All': 'Ges.'}, inplace=True)

        return(table_wuech)


    ###=========================================================================
    ###   5.2 Stichprobe
    ###=========================================================================

    def fuc_tbl_tax(self):

        ## prepare data

        # filter
        data_tax = self.data[(self.data['Best.-Schicht']!=0) & (self.data['Bewirtschaftungsform']=='W') & (self.data['Ertragssituation']=='I')]

        # create pivot table with "Vorrat"
        tax_v = pd.pivot_table(data_tax, index=['Forstrevier'], columns=['TAX: Altersklasse'], values=['Vorrat am Ort'], aggfunc=np.sum, fill_value=0, margins=True)
        tax_v.columns = tax_v.columns.droplevel()

        # create pivot table with "Fläche"
        tax_fl = pd.pivot_table(data_tax, index=['Forstrevier'], columns=['TAX: Altersklasse'], values=['Repr. Fläche Schicht'], aggfunc=np.sum, fill_value=0, margins=True)
        tax_fl.columns = tax_fl.columns.droplevel()

        # add Blöße to AKL I -> for comparibility with SPI
        tax_fl.loc[:,'I'] = tax_fl.loc[:,'BL'] + tax_fl.loc[:,'I']

        tax_v_ha = tax_v/tax_fl

        # round numbers
        tax_v = tax_v.round(0).astype(int)
        tax_fl = tax_fl.round(0).astype(int)
        tax_v_ha = tax_v_ha.round(0).astype(int)

        temp_fr = tax_v.index.values
        fr = temp_fr[:-1].tolist()
        fr = [temp_fr[-1]] + fr

        # reset index
        tax_v.reset_index(level=0, inplace=True)
        tax_fl.reset_index(level=0, inplace=True)
        tax_v_ha.reset_index(level=0, inplace=True)

        temp_tax = []

        for i in fr:
            x = tax_v[tax_v['Forstrevier']==i]
            y = tax_fl[tax_fl['Forstrevier']==i]
            z = tax_v_ha[tax_v_ha['Forstrevier']==i]

            xyz = pd.concat([x,y,z])
            xyz.iloc[:,1:] = xyz.iloc[:,1:].astype(int)
            xyz['Einheit'] = ['Vorrat [Vfm]', 'Fläche [ha]' , 'Vorrat/ Fläche [Vfm/ha]']

            xyz = xyz[['Einheit', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'All']]
            xyz.rename(columns={'All':'Ges.', 'VIII':'VIII+'}, inplace=True)
            temp_tax.append(xyz)

        return temp_tax


    def fuc_tbl_vergleich(self, tax_data, spi_data, fr):

        if fr == '0':
            text_fr = 'TO ' + str(self.dic.to)
        else:
            text_fr = 'FR ' + fr

        tax_data = tax_data[tax_data['Einheit']=='Vorrat/ Fläche [Vfm/ha]']
        tax_data = tax_data[['Einheit', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII+', 'Ges.']]

        spi_data = spi_data[spi_data['Einheit']=='Vorrat/ Fläche [Vfm/ha]']
        spi_data = spi_data[['Einheit', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII+', 'Ges.']]

        temp = pd.concat([tax_data,spi_data])
        temp['Einheit'] = ['Taxation','Stichprobe']

        temp = temp.set_index('Einheit')

        # plot

        fig = plt.figure(figsize=(6,3.2))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        temp.T.plot(kind='bar', ax=ax, linewidth = 2, alpha=0.9, title='Vergleich Stichprobe und Taxation - Wirtschaftswald, ' + text_fr)

        ax.set_ylabel("Masse [Vfm/ha]")
        ax.set_xlabel('Altersklassen')
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)
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

        return temp


    ###=========================================================================
    ###   5.4 Bonitaetsverlauf der BA
    ###=========================================================================

    def fuc_tbl_plt_bonitaetsverlauf(self):

        ## prepare data

        # create dataframe mittlere BG
        table_all_w = self.data_ekl.groupby('Baumart_group').apply(self.wavg,'Ertragsklasse','Repr. Fläche Baumart')

        # add every EKL to dataframe mittlere BG
        for i in np.sort(self.data_ekl['Wuechsigkeit'].unique()):
            data_i = self.data_ekl[self.data_ekl['Wuechsigkeit']==i].groupby('Baumart_group').apply(self.wavg,'Ertragsklasse','Repr. Fläche Baumart')
            table_all_w = pd.concat([table_all_w, data_i], axis=1, sort=False)

        # give columns their names
        table_all_w.columns = ['Ges.', 'sehr gering wüchsig', 'gering wüchsig', 'mittel-wüchsig', 'wüchsig', 'sehr wüchsig']


        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.2))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_all_w.iloc[:,1:].T.plot(kind='line', ax=ax, color=self.colors, alpha=0.9, title='Bonitätsverlauf')
        ax.set_ylabel("Mittlere Ertragsklasse [Vfm/ha/Jahr]")
        ax.set_xlabel("Wüchsigkeit")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)
        x = ['sehr gering', 'gering', 'mittel', 'gut', 'sehr']
        # create an index for each tick position
        xi = list(range(len(x)))
        plt.xticks(xi, x)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='center left', bbox_to_anchor=(1.07, 0.47), frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()

        ## table
        # sort table
        table_all_w = pd.DataFrame(table_all_w, columns=sorted(self.dic.dic_wuech_order, key=self.dic.dic_wuech_order.get))

        # get index inside data (for printing into docx)
        table_all_w.reset_index(level=0, inplace=True)

        # round numbers
        table_all_w = table_all_w.round(1)

        # fill nan with -
        table_all_w = table_all_w.fillna('-')

        # rename columns
        table_all_w.rename(columns={'Baumart_group': 'Baumarten'}, inplace=True)

        return(table_all_w)


    ###=========================================================================
    ###   5.5 Mittelwerte (alter, EKL, BG)
    ###=========================================================================

    def fuc_tbl_mittel(self, cat, code):
        '''
        cat: 'Schichtalter', 'Ertragsklasse', 'BaumartenBestockgrad'
        '''

        ## prepare data
        data_ekl_filter = self.filter_code(self.data_ekl,code)
        name = self.name_code(code)
        unique = self.data_ba['Baumart_group'].unique()

        # create dataframe mittlere EKL
        table_mittel = data_ekl_filter.groupby('Baumart_group').apply(self.wavg,cat,'Repr. Fläche Baumart')

        # add every EKL to dataframe mittlere EKL
        for i in np.sort(self.data_ekl['TAX: Altersklasse'].unique()):
            data_i = self.data_ekl[self.data_ekl['TAX: Altersklasse']==i].groupby('Baumart_group').apply(self.wavg,cat,'Repr. Fläche Baumart')
            table_mittel = pd.concat([table_mittel, data_i], axis=1)

        # give columns their names
        table_mittel.columns = ['Ges.', 'BL', 'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII+']

        # sort table & fill missing
        table_mittel = pd.DataFrame(table_mittel, index=sorted(self.dic.dic_ba_order, key=self.dic.dic_ba_order.get))
        # sort table & fill missing
        table_mittel = pd.DataFrame(table_mittel, columns=sorted(self.dic.dic_akl_order, key=self.dic.dic_akl_order.get))
        # filter sorted table
        table_mittel = table_mittel.loc[table_mittel.index.isin(unique), ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII+', 'Ges.']]


        # get index inside data (for printing into docx)
        table_mittel.reset_index(level=0, inplace=True)

        # round numbers
        table_mittel = table_mittel.round(1)

        # fill nan with -
        table_mittel = table_mittel.fillna('-')

        # rename columns
        table_mittel.rename(columns={'index': 'Baumarten'}, inplace=True)

        return(table_mittel)


    ###=========================================================================
    ###   5.7.1 Umtriebsgruppen nach Wuechsigkeit
    ###=========================================================================

    def fuc_tbl_plt_umtriebsgruppen_wuechsigkeit(self):

        ## prepare data

        # make pivot table
        table_wuech = pd.pivot_table(self.data_stoe, index=['Wuechsigkeit'],columns = ['Umtriebszeit'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table_wuech.columns = table_wuech.columns.droplevel()

        # rename index
        table_wuech.rename(index={1:'sehr gering wüchsig', 2:'gering wüchsig', 3:'mittelwüchsig', 4:'wüchsig' , 5:'sehr wüchsig', 'All': 'Ges.'}, inplace=True)
        table_wuech.rename(columns={'All':'Ges.'}, inplace=True)

        # calc percentage
        table_wuech_pre = (table_wuech.div(table_wuech.iloc[-1,:], axis=1))*100

        # round numbers
        table_wuech = table_wuech.round(1)
        table_wuech_pre_con = round(table_wuech_pre,0).astype(int)

        # rename columns
        header = []
        for i in range(table_wuech.shape[1]):
            temp = ['[ha]']
            header = header + temp
        table_wuech.columns = header
        header = []
        for i in range(table_wuech_pre_con.shape[1]):
            temp = ['[%]']
            header = header + temp
        table_wuech_pre_con.columns = header

        # make ha and percent fit in one table
        table_all = self.fuc_concat_tables(table_wuech, table_wuech_pre_con)

        # delete 'All' sum in index
        table_wuech_pre = table_wuech_pre[table_wuech_pre.index != 'Ges.']


        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_wuech_pre.T.plot(kind='bar', stacked=True, ax=ax, width = 0.5, alpha=0.9, title='Wüchsigkeit nach Umtrieb', color=['#FC0D1B', '#FDB32B', '#FFFD38', '#80DA29', '#108112'])
        #edgecolor='white', (possible)
        ax.set_ylabel("Prozent [%]")
        ax.set_xlabel("Umtriebszeit [a]")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

        # Remove plot frame on the right and upper spines
        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        # Only show ticks on the left and bottom spines
        ax.yaxis.set_ticks_position('left')
        ax.xaxis.set_ticks_position('bottom')

        # add legend
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.25), ncol=3, borderaxespad=0., frameon=False)
        # save plot to file
        plt.savefig('tempx.png', bbox_inches='tight', dpi=300)
        # clear plot
        plt.clf()
        plt.close()


        ## table

        # get index inside data (for printing into docx)
        table_all.reset_index(level=0, inplace=True)

        return(table_all)


    ###=========================================================================
    ###   5.7.2 Umtriebsgruppen nach Standortseinheit
    ###=========================================================================

    def fuc_tbl_plt_umtriebsgruppen_stoe(self):

        ## prepare data

        # make pivot table
        table_stoe = pd.pivot_table(self.data_stoe, index=['Standorteinheit'],columns = ['Umtriebszeit'], \
                           values=['Fläche in HA'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table_stoe.columns = table_stoe.columns.droplevel()

        # rename index
        table_stoe.rename(index=self.dic.dic_stoe_int, inplace=True)
        table_stoe.rename(index={'All':'Ges.'}, inplace=True)
        table_stoe.rename(columns={'All':'Ges.'}, inplace=True)

        # calc percent
        table_stoe_pre = (table_stoe.div(table_stoe.iloc[-1,:], axis=1))*100

        # round numbers
        table_stoe = table_stoe.round(1)
        table_stoe_pre_con = round(table_stoe_pre,0).astype(int)

        # rename columns
        header = []
        for i in range(table_stoe.shape[1]):
            temp = ['[ha]']
            header = header + temp
        table_stoe.columns = header
        header = []
        for i in range(table_stoe_pre_con.shape[1]):
            temp = ['[%]']
            header = header + temp
        table_stoe_pre_con.columns = header

        # make ha and percent fit in one table
        table_all = self.fuc_concat_tables(table_stoe, table_stoe_pre_con)

        # delete 'All' sum in index
        table_stoe_pre = table_stoe_pre[table_stoe_pre.index != 'Ges.']


        ## plot

        # Create a figure of given size
        fig = plt.figure(figsize=(6,3.8))
        # Add a subplot
        ax = fig.add_subplot(111)
        # plot
        table_stoe_pre.T.plot(kind='bar', stacked=True, ax=ax, width = 0.5, color=self.colors, alpha=0.9, title='Standortseinheiten in den Umtriebsgruppen')
        #edgecolor='white', (possible)
        ax.set_ylabel("Prozent [%]")
        ax.set_xlabel("Standortseinheit")
        ax.set_xticklabels(ax.xaxis.get_majorticklabels(), rotation=45)

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
        table_all.reset_index(level=0, inplace=True)
        #Standorteinheit
        table_all = table_all.rename(columns={'Standorteinheit':'STOE'})

        return(table_all)


    ###=========================================================================
    ### 6.2.1 Waldbaulicher HS
    ###=========================================================================

    def fuc_tbl_hs_waldbau(self, kind):
        '''
        kind: 'Forstrevier' | 'Umtriebszeit'
        '''

        # fliter columns
        data = self.data_vn_en[[kind, 'Bewirtschaftungsform', 'Massnahmengruppe', 'Nutzung LH', 'Nutzung NH']]

        # get unique
        kind_values = np.sort(data[kind].unique())
        print(kind_values)

        # add column with LH+NH
        data.loc[:,'Nutzung'] = data.loc[:,'Nutzung LH'] + data.loc[:,'Nutzung NH']

        # make tamplet
        #a = data.groupby([kind, 'Bewirtschaftungsform']).sum()
        #a[:] = 0

        # aggregated EN, VN, EN+VN
        x_list_1 = []
        # aggregated kind_values
        x_list_2 = []
        x_list_3 = []
        x_list_print = []

        for j in kind_values:

            # filter data with kind_value
            x_filter = data[data[kind] == j]

            # goupby Bewirtschaftungsform
            xx = x_filter.groupby([kind, 'Bewirtschaftungsform']).sum()
            # append EN + VN
            x_list_1 = [xx]

            for i in ['EN' ,'VN']:
                # goupby Bewirtschaftungsform & Massnahmengruppe
                xx = x_filter[x_filter['Massnahmengruppe'] == i].groupby([kind, 'Bewirtschaftungsform']).sum()
                # append EN & VN
                x_list_1.append(xx)

            # concat all lists into dataframe
            xxx = pd.concat([x_list_1[1], x_list_1[2], x_list_1[0]], axis=1)
            #print('xxx 1 xxx')
            #print(xxx)
            #x_list_x.append(xxx)
            xxx.reset_index(level=1, inplace=True)
            xxx.reset_index(level=0, inplace=True)
            xxx = xxx.T
            # concat sum
            xxx = pd.concat([xxx, xxx.sum(axis=1)], axis=1)
            xxx.iloc[0,:] = j
            xxx.iloc[1,-1] = 'Ges.'
            xxx.fillna(0, inplace=True)
            # append aggregated
            x_list_2.append(xxx.T)

        ## add missing rows

        # make templet for missing rows
        t = x_list_2[0][x_list_2[0]['Bewirtschaftungsform'] == 'Ges.'].copy()
        t[:] = 0

        for i in x_list_2:
            x_list_temp = []
            #i = i.T
            #print('xxx 2 xxx T')
            #print(i)
            i_unique = i['Bewirtschaftungsform'].unique()
            print('i_unique: ' + i_unique)
            i_uz = i[kind].unique()
            print('i_uz: ' + str(i_uz))
            for j in ['W','S', 'Ges.']:
                # add rows
                if j in i_unique:
                    tt = i[i['Bewirtschaftungsform'] == j]
                    #print('xx 11 xx if ' + j + ' in unique')
                    #print(tt)
                    x_list_temp.append(tt)
                # add missing rows
                else:
                    tt = t.copy()
                    tt[kind] = i_uz[0]
                    tt['Bewirtschaftungsform'] = j
                    #print('xx 12 xx else')
                    #print(tt)
                    x_list_temp.append(tt)

            ttt = pd.concat([x_list_temp[0], x_list_temp[1], x_list_temp[2]], axis=0)
            #print('xxx 3 xxx T')
            #print(ttt)
            x_list_3.append(ttt)

        # concat all kinds
        xxxx = t.copy()
        for i in x_list_3:
            xxxx = pd.concat([xxxx, i], axis=0)
        xxxx = xxxx.iloc[1:,:]

        # set all nan to 0
        xxxx.fillna(0, inplace=True)

        # set to int
        xxxx = xxxx.set_index('Bewirtschaftungsform')
        xxxx = xxxx.round(0)
        xxxx = xxxx.astype('int')
        xxxx.reset_index(level=0, inplace=True)

        # calc sum
        xxxx.columns = ['Bewirtschaftungsform', kind, 'EN_LH', 'EN_NH', 'EN_Sum', 'VN_LH', 'VN_NH', 'VN_Sum', 'Sum_LH', 'Sum_NH', 'Sum_Sum']
        # calc the sum through groupby
        group = xxxx.groupby(['Bewirtschaftungsform'], sort=False).sum()
        # set values from kind_value to TO
        group[kind] = 'TO'
        # inplace index
        group.reset_index(level=0, inplace=True)
        # concatonate with the sum
        xxxx = pd.concat([xxxx, group], axis=0)
        # reset index
        xxxx = xxxx.reset_index(drop=True)

        # sort columns
        sequence = [kind, 'Bewirtschaftungsform', 'EN_LH', 'EN_NH', 'EN_Sum', 'VN_LH', 'VN_NH', 'VN_Sum', 'Sum_LH', 'Sum_NH', 'Sum_Sum']
        xxxx = xxxx.reindex(columns=sequence)

        # replace
        xxxx['Bewirtschaftungsform'] = xxxx['Bewirtschaftungsform'].replace(['S', 'SS'], 'SW')
        xxxx['Bewirtschaftungsform'] = xxxx['Bewirtschaftungsform'].replace(['W'], 'WW')
        xxxx['Bewirtschaftungsform'] = xxxx['Bewirtschaftungsform'].replace(['Summe'], 'Ges.')

        if kind == 'Forstrevier':
            name = 'FR'
        elif kind == 'Umtriebszeit':
            name = 'UZ'
        xxxx.columns = [name, 'WW SW', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.']

        return(xxxx)


    ###=========================================================================
    ### 6.3.1 Waldbaulicher HS nach BKL
    ###=========================================================================

    def fuc_tbl_hs_waldbau_bkl(self,fr):

        x = self.data_vn_en[['Forstrevier', 'Betriebsklasse', 'Umtriebszeit', 'Bewirtschaftungsform', 'Ertragssituation', 'Massnahmengruppe', 'Nutzung LH', 'Nutzung NH']]
        # add column with LH+NH
        x.loc[:,'Nutzung'] = x.loc[:,'Nutzung LH'] + x.loc[:,'Nutzung NH']

        # filter data
        x = x[x['Forstrevier'] == fr]

        # make a pivot table
        table = pd.pivot_table(x, index=['Betriebsklasse', 'Umtriebszeit', 'Bewirtschaftungsform', 'Ertragssituation'],columns = ['Massnahmengruppe'], \
                           values=['Nutzung LH', 'Nutzung NH', 'Nutzung'], \
                          aggfunc=np.sum, fill_value=0, margins=True)
        # drop level
        table.columns = table.columns.droplevel()

        LH_EN = table.iloc[:,3].copy()
        LH_VN = table.iloc[:,4].copy()
        NH_EN = table.iloc[:,6].copy()
        NH_VN = table.iloc[:,7].copy()

        table.iloc[:,0] = LH_EN
        table.iloc[:,1] = NH_EN
        table.iloc[:,2] = LH_EN + NH_EN
        table.iloc[:,3] = LH_VN
        table.iloc[:,4] = NH_VN
        table.iloc[:,5] = LH_VN + NH_VN
        table.iloc[:,6] = LH_EN + LH_VN
        table.iloc[:,7] = NH_EN + NH_VN
        table.iloc[:,8] = LH_EN + LH_VN + NH_EN + NH_VN

        table = round(table,0).astype(int)

        table.reset_index(level=3, inplace=True)
        table.reset_index(level=2, inplace=True)
        table.reset_index(level=1, inplace=True)
        table.reset_index(level=0, inplace=True)

        table['Bewirtschaftungsform'] = table['Bewirtschaftungsform'].replace(['S', 'SS'], 'SW')
        table['Bewirtschaftungsform'] = table['Bewirtschaftungsform'].replace(['W'], 'WW')
        table['Betriebsklasse'] = table['Betriebsklasse'].replace(['All'], 'Ges.')

        table.columns = ['BKL', 'UZ', 'WW SW', 'Ert.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.', 'LH', 'NH', 'Ges.']
        table = table.sort_values(by=['BKL', 'UZ', 'WW SW', 'Ert.'], ascending=[True, True, False, False])

        return table


    ###=========================================================================
    ### 6.3.2 Nutzungsprofil
    ###=========================================================================

    def fuc_tbl_nutzungszeitpunkt(self, indexx, filterx, valuesx):

        table = pd.pivot_table(self.data_vn_en[self.data_vn_en['Massnahmengruppe'] == filterx], index=[indexx], columns = ['Forstrevier'], \
                           values=[valuesx], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table.columns = table.columns.droplevel()

        # calc percent
        table_pre = round(table.div(table.iloc[-1,:], axis=1 ),3)*100

        if valuesx == 'Nutzungssumme':
            unit = '[efm]'
        else:
            unit = '[ha]'

        # rename columns
        header = []
        for i in range(table.shape[1]):
            temp = [unit]
            header = header + temp
        table.columns = header
        header = []
        for i in range(table_pre.shape[1]):
            temp = ['[%]']
            header = header + temp
        table_pre.columns = header

        # make ha and percent fit in one table
        table_all = self.fuc_concat_tables(table, table_pre)

        #table_all.columns = ['FR1 [ha]', 'FR1 [%]', 'FR2 [ha]', 'FR2 [%]', 'TO [ha]', 'TO [%]']
        #table_all.columns = ['[ha]', '[%]', '[ha]', '[%]', '[ha]', '[%]']

        # rename columns
        table_all.rename(index={'TAX: Altersklasse':'AKL', 'All':'Ges.'}, inplace=True)

        # round
        table_all = round(table_all,0)
        table_all = table_all.astype(int)
        table_all.reset_index(level=0, inplace=True)
        table_all.rename(columns={'TAX: Altersklasse':'AKL', 'All':'Ges.', 'Maßnahmenart':'MA', 'Nutzdringlichkeit':'ND', 'Bewpfl.':'BP', 'Zeitpunkt':'NZ', 'Schlägerungsart':'SA', 'Rückungsart':'RA'}, inplace=True)

        return table_all


    ###=========================================================================
    ### 8. Schutzwald
    ###=========================================================================

    def fuc_tbl_schutzwald(self, cat):

        ## prepare data
        if cat == '0':
            data = self.data_wo_fl[self.data_wo_fl['Bewirtschaftungsform']=='S']
        else:
            data = self.data_wo_fl[(self.data_wo_fl['Bewirtschaftungsform']=='S') & (self.data_wo_fl['Schutzwaldkategorie']==cat)]

        # make pivot table
        table = pd.pivot_table(data, index=['Forstrevier'],columns = ['Ertragssituation'], \
                                       values=['Fläche in HA'], \
                                      aggfunc=np.sum, fill_value=0, margins=True)
        print(table)

        # drop level
        table.columns = table.columns.droplevel()

        # get all values from missing data
        uniq = ['A', 'I', 'All']
        #if not np.array_equal(table.T.index.values, uniq):
        if table.columns.values.tolist() != uniq:
            table = table.T.reindex(uniq, fill_value=0).T

        # calc percentage
        table_pre = (table.div(table.iloc[-1,-1], axis=1))*100
        print(table_pre)

        # round numbers
        table = table.round(1)
        table_pre_con = round(table_pre,0).astype(int)

        # rename columns
        header = []
        for i in range(table.shape[1]):
            temp = ['[ha]']
            header = header + temp
        table.columns = header
        header = []
        for i in range(table_pre_con.shape[1]):
            temp = ['[%]']
            header = header + temp
        table_pre_con.columns = header

        # make ha and percent fit in one table
        table_all = self.fuc_concat_tables(table, table_pre_con)

        # rename index
        table_all.rename(index={'All':'Gesamt'}, inplace=True)

        table_all.reset_index(level=0, inplace=True)

        return table_all


    ## weight average (calc)

    def wavg(self,group,avg_name,weight_name):

        d = group[avg_name]
        w = group[weight_name]

        try:
            return (d*w).sum() / w.sum()
        except ZeroDivisionError:
            return 0

    ## for concatonating two tables (ha|%)

    def fuc_concat_tables(self,table_base,table_pre):

        # create table_all with first two columns
        table_all = pd.concat([table_base.iloc[:,0], table_pre.iloc[:,0]], axis=1)

        # concat remaining columns to table_all
        for i in np.arange(1,table_base.shape[1]):
            table_all = pd.concat([table_all.iloc[:,:], table_base.iloc[:,i], table_pre.iloc[:,i]], axis=1)

        return(table_all)


    #***************************************************************************


    ###=========================================================================
    ###   loop for Revier and Umtriebszeit
    ###=========================================================================

    def loop_fr_uz(self):
        '''
        generate array for looping through all cases
        input:  data
        output: FR  UZ  WW/SW
        '''

        frx = np.sort(self.data_V['Forstrevier'].unique())
        y = np.sort(self.data_V['Bewirtschaftungsform'].unique())
        uzx = np.sort(self.data_V['Umtriebszeit'].unique())

        # insert 0 for not
        y = np.insert(y,0,0)

        loop_array = np.array(([[frx[0],'0',y[0]]]))

        for fr in frx:
            # filter
            data = self.data_fl[self.data_fl['Forstrevier']==fr]
            bfx = np.sort(data['Bewirtschaftungsform'].unique())[::-1]
            if (['W'] in bfx) & (['S'] in bfx):
                bfx = np.array(['0', 'W', 'S'])
            for bf in bfx:
                temp = np.array([[fr,'0',bf]])
                loop_array = np.append(loop_array,temp,axis=0)

        for uz in uzx:
            temp = np.array([['0',uz,'0']])
            loop_array = np.append(loop_array,temp,axis=0)

        loop_array = loop_array[1:,:]

        return(loop_array)


    def loop_be(self):
        '''
        generate array for looping through all cases
        input:  data
        output: WW/SW
        '''

        y = np.sort(self.data_V['Bewirtschaftungsform'].unique())[::-1]

        loop_array = np.array(([['0','0','0']]))

        for i in y:
            temp = np.array([['0','0',i]])
            loop_array = np.append(loop_array,temp,axis=0)

        return(loop_array)

    def loop_fr(self):
        '''
        generate array for looping through all cases
        input:  data
        output: FR
        '''

        x = np.sort(self.data_V['Forstrevier'].unique())

        loop_array = np.array(([['0','0','0']]))

        for i in x:
            temp = np.array([[i,'0','0']])
            loop_array = np.append(loop_array,temp,axis=0)

        #loop_array = loop_array[1:,:]

        return(loop_array)


    ###=========================================================================
    ###   filter
    ###=========================================================================

    def filter_code(self,data,code):
        '''
        generate array for looping through all cases
        input:  data
        output: filterd data
        '''

        # Teiloperat
        if (code[0] == '0') & (code[1] == '0'):

            if code[2] == '0':
                data = data.copy()
            else:
                data = data[data['Bewirtschaftungsform']==code[2]]

        # Forstrevier
        elif (code[0] != '0') & (code[1] == '0'):

            if code[2] == '0':

                data = data[(data['Forstrevier']==int(code[0]))]
            else:
                data = data[(data['Forstrevier']==int(code[0])) & (data['Bewirtschaftungsform']==code[2])]

        # Umtriebszeit
        elif (code[0] == '0') & (code[1] != '0'):

            data = data[(data['Umtriebszeit']==int(code[1]))]

        return(data)


    def name_code(self,code):
        '''
        generate array for naming
        input:  [FR | UZ | WW/SW]
        output: FR 6 Frein
        '''

        x = ['0', '0', '0', '0']

        # Forstrevier
        if (code[0] != '0'):
            x[0] = 'FR ' + code[0]
            x[1] = 'FR ' + code[0] + ' ' + self.dic.dic_num_fr[int(code[0])]
        elif (code[1] != '0'):
            x[0] = 'U ' + code[1]
            x[1] = 'Umtrieb ' + code[1] + ' Jahre'
        else:
            x[0] = 'TO ' + str(self.dic.to)
            x[1] = 'TO ' + str(self.dic.to) + ' ' + self.dic.dic_num_fb[self.dic.fb]

        # Bewirtschaftungsform
        if (code[1] == '0'):
            if code[2] == '0':
                x[2] = 'WW+SW'
                x[3] = 'Wirtschafts- und Schutzwald'
            else:
                x[2] = code[2] + 'W'
                if code[2] == 'W':
                    x[3] = 'Wirtschaftswald'
                elif code[2] == 'S':
                    x[3] = 'Schutzwald'
        else:
            x[2] = ''
            x[3] = ''

        return(x)


    def name_ng(self,cat, fr='All'):

        x = ['0', '0', '0', '0', '0']

        #color='#f2500a'
        #self.colors = ['#3976AF', '#EF8536', '#529D3E', '#C53932', '#8D6AB8', '#DEAD3B', '#85584E', '#D57DBE', '#7F7F7F', '#57BBCC', '#4291F5', '#2B6419', '#E96B77', '#741E18', '#FEEC61']

        if cat == 'Seehöhe':
            x[0] = 'Seehöhe [m]'
            x[1] = 'Seehöhenverteilung'
            x[4] = '#f2500a'
        elif cat == 'Exposition':
            x[0] = 'Exposition'
            x[1] = 'Exposition'
            x[4] = '#3976AF'
        elif cat == 'NeigGruppe':
            x[0] = 'Neigungsgruppe [%]'
            x[1] = 'Neigungsverteilung'
            x[4] = '#529D3E'

        if fr == 'All':
            x[2] = 'TO'
            x[3] = 'TO ' + str(self.dic.to)
        else:
            x[2] = self.dic.dic_num_fr[fr]
            x[3] = 'FR ' + self.dic.dic_num_fr[fr]

        return(x)


    def fuc_align_yaxis(self, ax1, v1, ax2, v2):
        """adjust ax2 ylimit so that v2 in ax2 is aligned to v1 in ax1"""
        _, y1 = ax1.transData.transform((0, v1))
        _, y2 = ax2.transData.transform((0, v2))
        inv = ax2.transData.inverted()
        _, dy = inv.transform((0, 0)) - inv.transform((0, y1-y2))
        miny, maxy = ax2.get_ylim()
        ax2.set_ylim(miny+dy, maxy+dy)


    ### 7. Waldpflege - Ziele und Planung für die nächste Periode

    def fuc_tbl_waldpflege(self, filterx, indexx, columnsx):

        ## prepare data
        if filterx == ['DE']:
            # Vornutzungsmaßnahmen = VN (Vornutzung)
            data = self.data_massnahmen[self.data_massnahmen['Massnahmengruppe'] == 'VN']
        else:
            # Waldpflegemaßnahmen = filter WP (Waldpflege)
            data = self.data_massnahmen[self.data_massnahmen['Massnahmengruppe'] == 'WP']


        table = pd.pivot_table(data[data['Maßnahmenart'].isin(filterx)], index=[indexx], columns = [columnsx], \
                           values=['Angriffsfläche'], \
                          aggfunc=np.sum, fill_value=0, margins=True)

        # drop level
        table.columns = table.columns.droplevel()

        # get all values from missing data
        uniq =  np.sort(self.data_massnahmen['Forstrevier'].unique()).tolist() + ['All']
        if not table.T.index.values.tolist() == uniq:
            table = table.T.reindex(uniq, fill_value=0).T

        # calc percent
        table_pre = round(table.div(table.iloc[-1,:], axis=1 ),3)*100

        # rename columns
        header = []
        for i in range(table.shape[1]):
            temp = ['[ha]']
            header = header + temp
        table.columns = header
        header = []
        for i in range(table_pre.shape[1]):
            temp = ['[%]']
            header = header + temp
        table_pre.columns = header

        # make ha and percent fit in one table
        table_all = table_all = self.fuc_concat_tables(table, table_pre)

        table_all = round(table_all,0)
        # rename index
        table_all.rename(index={'All':'Ges.'}, inplace=True)

        table_all.reset_index(level=0, inplace=True)
        table_all.rename(columns={'TAX: Altersklasse':'AKL', 'All':'Ges.', 'Maßnahmenart':'MA', 'Nutzdringlichkeit':'ND', 'Bewpfl.':'BP', 'Zeitpunkt':'NZ', 'Schlägerungsart':'SA', 'Rückungsart':'RA'}, inplace=True)

        # delete 'All' sum in index
        #table_stoe_pre = table_stoe_pre[table_stoe_pre.index != 'All']

        return table_all


    def wpplan_loop(self):

        fr_num = self.data_fl['Forstrevier'].unique()
        fr_name = [self.dic.dic_num_fr[i] for i in fr_num]
        loop = ['Gesamtergebnis']
        loop = loop + fr_name

        return loop

    # aggregated numbers of Abteilungen  - used for Flächentabelle
    def abteilungen_numbers(self, kind):

        '''
        input: 'Forstrevier' | 'Betriebsklasse'
        '''

        abt = []

        fr = np.sort(self.data_fl[kind].unique())


        for j in fr:
            # get data
            data_filtered = self.data_fl[self.data_fl[kind]==j]
            # get unique sorted values
            all_unique_abt = np.sort(data_filtered['Abteilung'].unique())
            # Abteilungsnummern
            abt_temp = ''
            i_old = 0
            i_anf = all_unique_abt[0]
            for i in all_unique_abt:
                if i_old != 0:
                    if i-i_old != 1:
                        if i_anf == i_old:
                            abt_temp = abt_temp + str(i_anf) + ' '
                        else:
                            #print(str(i_anf) + '-' + str(i_old))
                            abt_temp = abt_temp + str(i_anf) + '-' + str(i_old) + ' '
                            i_anf = i
                i_old = i
            #print(str(i_anf) + '-' + str(i_old))
            abt_temp = abt_temp + str(i_anf) + '-' + str(i_old)
            abt.append(abt_temp)
        abt.append('Alle')

        index = fr.tolist()
        index.append('All')

        abt = pd.Series(abt, index=index)

        return abt

    # get min Umtriebszeit out of variable 'code'
    def get_min_uz(self, code):
        uz = self.loop_fr_uz()[:,1].astype(int)
        uz = uz[uz!=0].min()
        return uz
