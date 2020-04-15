## custom dictionaries

class OBFDictionary(object):

    def __init__(self):

        # make a dictionary between number and Forstbetrieb
        self.dic_num_fb = {171:'Wienerwald', 172:'Waldviertel-Voralpen', 173:'Steiermark', 174:'Steyrtal', 175:'Traun-Innviertel', 176:'Inneres Salzkammergut', 177:'Kärnten-Lungau', 178:'Flachgau-Tennengau', 179:'Pongau', 180:'Pinzgau', 181:'Unterinntal', 182:'Oberinntal'}

        # make a dictionary Wüchsigkeit from number to text
        self.dic_num_wuech = {1:'sehr gering wüchsig', 2:'gering wüchsig', 3:'mittelwüchsig', 4:'wüchsig', 5:'sehr wüchsig'}

        # make a dictionary between STOE and Geologie
        self.dic_stoe_geo = {1:'Silikat', 2:'Silikat', 3:'Silikat', 4:'Silikat', 5:'Silikat', 6:'Flysch', 7:'Flysch', 8:'Flysch', 9:'Sediment', 11:'Karbonat', 12:'Karbonat', 13:'Karbonat', 14:'Karbonat', 15:'Karbonat', 21:'Karbonat', 22:'Karbonat', 23:'Karbonat', 25:'Karbonat', 26:'Karbonat', 27:'Karbonat', 28:'Karbonat', 31:'Karbonat', 32:'Karbonat', 33:'Karbonat', 34:'Karbonat', 35:'Karbonat', 36:'Karbonat', 37:'Karbonat', 41:'Karbonat', 42:'Karbonat', 43:'Karbonat', 44:'Karbonat', 51:'Karbonat', 52:'Karbonat', 53:'Karbonat', 54:'Karbonat', 55:'Karbonat', 56:'Karbonat', 57:'Karbonat', 58:'Karbonat', 61:'Silikat', 62:'Silikat', 62:'Silikat', 63:'Silikat', 64:'Silikat', 65:'Silikat', 66:'Flysch', 71:'Silikat', 72:'Silikat', 73:'Silikat', 74:'Silikat', 75:'Silikat', 76:'Flysch', 76:'Sediment', 81:'Silikat', 81:'Silikat', 82:'Silikat', 82:'Silikat', 83:'Silikat', 84:'Silikat', 85:'Silikat', 86:'Flysch', 86:'Flysch', 87:'Flysch', 88:'Flysch', 88:'Flysch', 89:'Flysch', 91:'Silikat', 92:'Sediment', 93:'Sediment', 94:'Sediment', 95:'Sediment', 96:'Flysch', 97:'Flysch', 98:'Flysch'}

        # make a dictionary between STOE and Wüchsigkeit
        self.dic_stoe_wuech = {1:3, 2:2, 3:3, 4:3, 5:3, 6:4, 7:3, 8:5, 9:5, 11:1, 12:2, 13:1, 14:1, 15:1, 21:2, 22:3, 23:3, 25:2, 26:2, 27:3, 28:3, 31:4, 32:5, 33:5, 34:3, 35:3, 36:4, 37:4, 41:5, 42:4, 43:5, 44:5, 51:3, 52:2, 53:4, 54:4, 55:3, 56:4, 57:2, 58:3, 61:1, 62:1, 62:1, 63:1, 64:1, 65:1, 66:1, 71:2, 72:3, 73:2, 74:2, 75:2, 76:2, 76:2, 81:4, 81:4, 82:5, 82:5, 83:2, 84:4, 85:5, 86:4, 86:4, 87:4, 88:5, 88:5, 89:5, 91:3, 92:3, 93:4, 94:4, 95:3, 96:5, 97:5, 98:5}

        # make a dictionary between Baumart and LH | NH
        self.dic_ba_lh_nh = {"AS":"LH", "AZ":"NH", "AH":"LH", "PM":"NH", "BI":"LH", "FB":"NH", "BL":"BL", "AC":"NH", "DG":"NH", "EE":"LH", "EK":"LH", "EB":"NH", "EI":"LH", "EL":"LH", "ES":"LH", "EA":"LH", "FA":"LH", "FI":"NH", "FE":"LH", "FZ":"NH", "GK":"NH", "GB":"LH", "WP":"LH", "GE":"LH", "AG":"NH", "AV":"LH", "HB":"LH", "HT":"NH", "HP":"LH", "JL":"NH", "CJ":"NH", "KK":"NH", "KB":"LH", "KO":"NH", "LA":"NH", "LI":"LH", "ME":"LH", "AN":"NH", "FO":"NH", "PO":"LH", "AB":"NH", "RO":"LH", "RK":"LH", "BU":"LH", "RE":"LH", "SW":"LH", "CH":"NH", "ER":"LH", "SK":"NH", "JN":"LH", "SP":"LH", "SF":"NH", "LS":"LH", "SL":"LH", "SN":"NH", "SG":"LH", "PU":"NH", "SA":"LH", "QR":"LH", "ST":"LH", "KW":"NH", "TH":"NH", "QP":"LH", "TK":"LH", "TB":"LH", "UL":"LH", "NU":"LH", "WD":"LH", "KI":"NH", "TA":"NH", "WO":"LH", "LW":"LH", "EZ":"LH", "ZI":"NH"}

        # ba order in sorting
        self.dic_ba_order = {"FI":0, "TA":1, "LA":2, "KI":3, "SK":4, "ZI":5, "PM":6, "DG":7, "KW":8, "SF":9, "FO":10, "TH":11, "FB":12, "AC":13, "AG":14, "AZ":15, "EB":16, "FZ":17, "GK":18, "HT":19, "JL":20, "CJ":21, "KK":22, "KO":23, "AN":24, "AB":25, "CH":26, "PU":27, "SN":28, "BU":29, "EI":30, "HB":31, "AH":32, "SA":33, "FA":34, "EA":35, "ES":36, "UL":37, "QP":38, "QR":39, "EZ":40, "RE":41, "FE":42, "ER":43, "GE":44, "AV":45, "KB":46, "TK":47, "WO":48, "SG":49, "NU":50, "JN":51, "LI":52, "LS":53, "LW":54, "BI":55, "PO":56, "AS":57, "WP":58, "SP":59, "HP":60, "WD":61, "SW":62, "EK":63, "RK":64, "EE":65, "EL":66, "ME":67, "RO":68, "TB":69, "GB":70, "ST":71, "SL":72, "BL":73}

        # exp order in sorting
        self.dic_exp_order = {"EB":0, "KU":1, "MU":2, "N":3, "NO":4, "O":5, "SO":6, "S":7, "SW":8, "W":9, "NW":10}

        # akl order in sorting
        self.dic_akl_spi_order = {"I":0, "II":1, "III":2, "IV":3, "V":4, "VI":5, "VII":6, "VIII+":7, "Ges.":8}

        # akl order in sorting
        self.dic_akl_plot_order = {"BL":0, "I":1, "II":2, "III":3, "IV":4, "V":5, "VI":6, "VII":7, "VIII+":8}

        # akl order in sorting
        self.dic_akl_order = {"BL":0, "I":1, "II":2, "III":3, "IV":4, "V":5, "VI":6, "VII":7, "VIII+":8, "Ges.":9}

        # wuechsigkeit order in sorting
        self.dic_wuech_order = {'sehr gering wüchsig':1, 'gering wüchsig':2, 'mittel-wüchsig':3, 'wüchsig':4, 'sehr wüchsig':5, 'Ges.':6}

        # convert stoe to int
        self.dic_stoe_int = {1:int(1), 2:int(2), 3:int(3), 4:int(4), 5:int(5), 6:int(6), 7:int(7), 8:int(8), 9:int(9), 11:int(11), 12:int(12), 13:int(13), 14:int(14), 15:int(15), 21:int(21), 22:int(22), 23:int(23), 25:int(25), 26:int(26), 27:int(27), 28:int(28), 31:int(31), 32:int(32), 33:int(33), 34:int(34), 35:int(35), 36:int(36), 37:int(37), 41:int(41), 42:int(42), 43:int(43), 44:int(44), 51:int(51), 52:int(52), 53:int(53), 54:int(54), 55:int(55), 56:int(56), 57:int(57), 58:int(58), 61:int(61), 62:int(62), 63:int(63), 64:int(64), 65:int(65), 66:int(66), 71:int(71), 72:int(72), 73:int(73), 74:int(74), 75:int(75), 76:int(76), 81:int(81), 82:int(82), 83:int(83), 84:int(84), 85:int(85), 86:int(86), 86:int(86), 87:int(87), 88:int(88), 89:int(89), 91:int(91), 92:int(92), 93:int(93), 94:int(94), 95:int(95), 96:int(96), 97:int(97), 98:int(98)}

        # spi dict
        self.dic_schael = {1:'Ernteschaden', 2:'Wegebau', 3:'Steinschlag', 4:'Fegung-/Schlagbaum', 5:'Peitschung', 6:'abiotischer Schaden', 7:'biotischer Schaden', 8:'fr. Borken-käferbefall', 9:'Sonstiges', 10:'kein Schaden', 'All':'Ges.'}
        self.dic_schael_plot = {1:'Ernteschaden', 2:'Wegebau', 3:'Steinschlag', 4:'Fegung/Schlag', 5:'Peitschung', 6:'abiotischer Schaden', 7:'biotischer Schaden', 8:'fr. Borkenkäfer', 9:'Sonstiges', 10:'kein Schaden', 'All':'Ges.'}
        self.dic_kronv = {0:'unter Kluppschwelle' ,1:'keine', 2:'leicht', 3:'deutlich', 4:'stark', 5:'tot', 6:'keine Ansprache', 'All':'Ges.'}
        self.dic_qual = {1:'gut', 2:'mittel', 3:'schlecht', 10:'keine Ansprache', 'All':'Ges.'}

        # Mittelwerte - headings and units
        self.dic_avg_head = {"Schichtalter":"Mittleres Alter", "Ertragsklasse":"Mittlere Ertragsklasse", "BaumartenBestockgrad":"Mittlerer Bestockungsgrad"}
        self.dic_avg_unit = {"Schichtalter":" [a]", "Ertragsklasse":" [Vfm/ha/a]", "BaumartenBestockgrad":""}

        # Einschlag vs Hiebsatz - headings and units
        self.dic_es_hs_head = {"Altersgruppe 1. Schicht":"Alter", "Neigungsgruppe (%)":"Hangneigung", "Seehöhen Gruppe":"Seehöhe", "Umtriebszeit":"Umtriebsgruppen"}
        self.dic_es_hs_unit = {"Altersgruppe 1. Schicht":"Alter [a]", "Neigungsgruppe (%)":"Hangneigung [%]", "Seehöhen Gruppe":"Seehöhe [m]", "Umtriebszeit":"Umtriebsgruppen [a]"}

        # Nutzungs Legende
        self.dic_nutz_dring = {1:'Waldbaulich dringend notwendig (innerhalb von 3 Jahren)', 2:'Waldbaulich notwendig (innerhalb von 10 Jahren)', 3:'Waldbaulich nicht erforderlich (im Planungszeitraum möglich)', 4:'Erst nach Lösung der Wildfrage (nicht im Hiebssatz)', 5:'Erst nach erfolgter Aufschließung (im Hiebssatz)', 6:'Derzeit Nutzung ohne positiven DB I'}
        self.dic_nutz_bewilg = {1:'Keine Bewilligung erforderlich', 2:'Bewilligungspflichtig', 3:'Meldepflichtig'}
        self.dic_nutz_zeit = {1:'Sommer', 2:'Winter', 3:'Ganzjährig'}
        self.dic_nutz_schlaeg = {1:'Sortimentsverfahren', 2:'Stammverfahren', 3:'Baumverfahren', 4:'Harvester', 5:'Baumverfahren abgezopft u. grob geastet', 6:'Seil-Harvester'}
        self.dic_nutz_rueck = {10:'Händisch', 20:'Seil (bis 2009)', 23:'Seil bergauf', 26:'Seil bergab', 29:'Seil Langstrecke', 30:'Schlepper', 31:'Traktor', 35:'Forwarder', 36:'Seil-Forwarder', 38:'Hubschrauber', 40:'Sonstiges', 90:'Holz verbleibt am Waldort'}
        self.dic_nutz_ma = {'AD':'Abdeckung', 'AE':'Aufhieb forstlicher Einteilungslinien', 'AF':'Aufforstung', 'AG':'Astung', 'AS':'Abstockung', 'BU':'Bestandesumwandlung', 'DE':'Erstdurchforstung', 'DF':'Durchforstung', 'DP':'Dickungspflege', 'EG':'Ergänzung', 'FM':'Femelung', 'JF':'Jungwuchsfreistellung, kleinflächig', 'JP':'Jungwuchspflege', 'KE':'Kulturschutz einzeln', 'KF':'Kulturschutz flächig', 'KH':'Kahlhieb', 'LI':'Lichtung', 'LL':'Loslösung', 'NB':'Nachbesserung', 'ND':'Niederdurchforstung', 'PA':'Protzenaushieb', 'PL':'Plenterung', 'RM':'Räumung, großflächig über Verjüngung', 'SB':'Säuberung', 'TR':'Trassenaufhieb', 'UE':'Überhälterentnahme', 'ZE':'Zufällige Ergebnisse im Endnutzungsbestand', 'ZN':'Zielstärkennutzung', 'ZV':'Zufällige Ergebnisse in Vornutzungsbeständen'}

    def print_fb(self):
        print(self.dic_num_fb)
        return(self.dic_num_fb)

    def set_fb(self, fb):
        self.fb = fb

        # make a dictionary between number and Forstrevier

        if self.fb == 171:
            self.dic_num_fr = {1:'Kierling', 2:'Weidlingbach', 3:'Stadlhütte', 4:'Ried', 5:'Pressbaum', 6:'Klausen', 7:'Schöpflgitter', 8:'Alland', 9:'Breitenfurt', 10:'Hinterbrühl', 11:'Haselbach', 12:'Pernitz', 13:'Oberwart'}
        elif self.fb == 172:
            self.dic_num_fr = {1:'Eisenbergeramt', 2:'Droß', 3:'Weißenkirchen', 4:'Münchreith', 5:'Leiben', 6:'Türnitz', 7:'Gaming', 8:'Hollenstein', 9:'Göstling'}
        elif self.fb == 173:
            self.dic_num_fr = {1:'Großreifling', 2:'Wildalpen', 3:'Gußwerk', 4:'Wegscheid', 5:'Mariazell', 6:'Frein', 7:'Mürzsteg', 8:'Neuberg', 9:'Mürzzuschlag', 10:'Lankowitz'}
        elif self.fb == 174:
            self.dic_num_fr = {1:'Steyr', 2:'Großraming', 3:'Brunnbach', 4:'Reichraming', 5:'Sattl', 6:'Molln', 7:'Breitenau', 8:'Windischgarsten', 9:'Pyhrn'}
        elif self.fb == 175:
            self.dic_num_fr = {1:'Offensee', 2:'Ebensee', 3:'Traunstein', 4:'Neukirchen', 5:'Reindlmühl', 6:'Attergau', 7:'Loibichl', 8:'Mondsee', 9:'Schneegattern', 10:'Frauschereck', 11:'Bradirn'}
        elif self.fb == 176:
            self.dic_num_fr = {1:'Mitterweißenbach', 2:'Rettenbach', 3:'Lasern', 4:'Altaussee', 5:'Grundlsee', 6:'Mitterndorf', 7:'Kemetgebirge', 8:'Bad Aussee', 9:'Hallstatt', 10:'Gosau', 11:'Lauffen'}
        elif self.fb == 177:
            self.dic_num_fr = {1:'Millstatt', 2:'Obervellach', 3:'Hermagor', 4:'Bleiberg', 5:'Ossiach', 6:'Tamsweg', 7:'St. Michael', 8:'Mauterndorf', 9:'Zederhaus'}
        elif self.fb == 178:
            self.dic_num_fr = {1:'Wiestal', 2:'Faistenau', 3:'Hintersee', 4:'St.Gilgen', 5:'Strobl', 6:'Osterhorn', 7:'Abtenau', 8:'Annaberg', 9:'St.Martin', 10:'Blühnbach'}
        elif self.fb == 179:
            self.dic_num_fr = {1:'Filzmoos', 2:'Radstadt', 3:'Flachau', 4:'Kleinarl', 5:'Gründeck', 6:'Bischofshofen', 7:'Großarl', 8:'Gastein', 9:'Lend', 10:'Taxenbach'}
        elif self.fb == 180:
            self.dic_num_fr = {1:'Wald', 2:'Habach', 3:'Mühlbach', 4:'Mittersill', 5:'Stubach', 6:'Piesendorf', 7:'Bruck', 8:'Glemmtal', 9:'Alm', 10:'Saalfelden'}
        elif self.fb == 181:
            self.dic_num_fr = {1:'Kössen', 2:'Fieberbrunn', 3:'Kitzbühel', 4:'Brixental', 5:'Gerlos', 6:'Hinteres Zillertal', 7:'Alpbach', 8:'Johannklause', 9:'Marchbach', 10:'Thiersee', 52:'Dielmann'}
        elif self.fb == 182:
            self.dic_num_fr = {1:'Steinberg', 2:'Achenwald-Bächental', 3:'Achensee', 4:'Hinterriß', 5:'Inntal', 6:'Reutte', 7:'Telfs', 8:'Landeck', 9:'Pfunds'}

    def set_to(self, to):
        self.to = to

    def set_wg(self, wg):
        self.wg = wg

    def set_to_old(self, to):
        self.to_old = to_old

    def set_laufzeit(self, laufzeit):
        self.laufzeit = laufzeit

    def set_laufzeit_begin(self, jahr):
        self.laufzeit_begin = jahr

    def set_laufzeit_end(self, jahr):
        self.laufzeit_end = jahr

    def set_sl_sn(self, sl, sn):
        self.sl = sl
        self.sn = sn

    def set_uz(self, uz):
        self.uz = uz

    def set_fr(self, fr):
        self.fr = fr

    def set_dring(self, dring):
        self.dring = dring

    # make dict for flaeche & vorrat
    def set_dic_fr(self):
        key = []
        for i in self.fr:
            key.append(i)
        key.append('TO')
        values = range(len(key))
        self.dic_fr_order = dict(zip(key,values))

    # make legendes
    def nutz_legend(self, unique, dict_name):
        if (dict_name == 'Altersklassen'):
            return ''

        elif (dict_name == 'Nutzungsarten') | (dict_name == 'Maßnahmenart'):
            string = ';'
            for i in unique:
                string = string + ' ' + i + ' = ' + self.dic_nutz_ma[i] + ','
            string = string[:-1] + '.'
            return string

        else:
            if (dict_name == 'Dringlichkeit') | (dict_name == 'Nutzdringlichkeit'):
                dictonary = self.dic_nutz_dring
            elif dict_name == 'Melde-/ Bewilligungspflicht':
                dictonary = self.dic_nutz_bewilg
            elif dict_name == 'Nutzungszeitpunkt':
                dictonary = self.dic_nutz_zeit
            elif dict_name == 'Schlägerungsart':
                dictonary = self.dic_nutz_schlaeg
            elif dict_name == 'Rückungsart':
                dictonary = self.dic_nutz_rueck

            string = ';'
            for i in unique:
                string = string + ' ' + str(int(i)) + ' = ' + dictonary[i] + ','
            string = string[:-1] + '.'
            return string
