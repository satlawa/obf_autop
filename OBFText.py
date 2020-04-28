################################################################################
#
#   fuc_txt_waldbau_grundsatz (doc)
#   fuc_txt_wuchsgebiet (doc,wuchs)
#   fuc_txt_naturschutz (doc)
#
################################################################################

import numpy as np
import pandas as pd

class OBFText(object):

    def __init__(self):

        self

    def set_legend(self, leg_ma_dring, leg_ma_bewill, leg_ma_zeit, leg_ma_schlag, leg_ma_rueck, leg_ma, leg_exp, leg_stoe, leg_es_ze, leg_es_zv):
        self.leg_ma_dring = leg_ma_dring #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Dringlichkeit')
        self.leg_ma_bewill = leg_ma_bewill #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Bewilligung')
        self.leg_ma_zeit = leg_ma_zeit #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Nutzzeitpunkt')
        self.leg_ma_schlag = leg_ma_schlag #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Schlägerungsart')
        self.leg_ma_rueck = leg_ma_rueck #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Rückungsart')

        self.leg_ma = leg_ma #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Massnahmenart')
        self.leg_exp = leg_exp #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Exposition')
        self.leg_stoe = leg_stoe #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'Standort')

        self.leg_es_ze = leg_es_ze #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'ZE')
        self.leg_es_zv = leg_es_zv #pd.read_excel('/Users/philipp/Python/docx/resources/key.xlsx', 'ZV')

    def fuc_leg_ma_dring(self):
        return self.leg_ma_dring

    def fuc_leg_me_bewill(self):
        return self.leg_ma_bewill

    def fuc_leg_ma_zeit(self):
        return self.leg_ma_zeit

    def fuc_leg_ma_schlag(self):
        return self.leg_ma_schlag

    def fuc_leg_ma_rueck(self):
        return self.leg_ma_rueck

    def fuc_loop_nutz(self):
        array = [['TAX: Altersklasse', 'Altersklassen'],
        ['Maßnahmenart', 'Nutzungsarten', self.leg_ma[(self.leg_ma['Massnahmengruppe']=='EN') | (self.leg_ma['Massnahmengruppe']=='VN')].iloc[:,:2]],
        ['Nutzdringlichkeit', 'Dringlichkeit', self.leg_ma_dring],
        ['Bewpfl.', 'Melde-/ Bewilligungspflicht', self.leg_ma_bewill],
        ['Zeitpunkt', 'Nutzungszeitpunkt', self.leg_ma_zeit],
        ['Schlägerungsart', 'Schlägerungsart', self.leg_ma_schlag],
        ['Rückungsart', 'Rückungsart', self.leg_ma_rueck]]

        return array

    ### Text

    def fuc_txt_waldbau_grundsatz(self, doc):

        doc.add_paragraph('Im ÖBf-Waldbauhandbuch sind die waldbaulichen Leitziele formuliert. Demnach werden die Erhaltung und die Verbesserung der Waldsubstanz sowie die nachhaltige Erfüllung der multifunktionalen Anforderungen an den Wald angestrebt. Diese Ziele können nur erreicht werden, wenn die Waldnutzung unter Berücksichtigung folgender ökologischer und ökonomischer Grundsätze erfolgt:')
        doc.add_paragraph('')

        # bullet list
        doc.add_paragraph('Zusammensetzung der Bestände aus standortstauglichen Baumarten (in der Regel hauptsächlich Baumarten der natürlichen Waldgesellschaft)', style='List Bullet')
        doc.add_paragraph('Erneuerung der Bestände möglichst durch Naturverjüngung', style='List Bullet')
        doc.add_paragraph('Erhaltung lokaler Besonderheiten und Kleinbiotope', style='List Bullet')
        doc.add_paragraph('Förderung der Biodiversität', style='List Bullet')
        doc.add_paragraph('Anbau nicht heimischer Baumarten (primär Douglasie) nur in Gebieten, in denen die heimischen Baumarten Probleme bereiten (mangelnde Trockenheitsresistenz, Stabilität, Qualität und Wuchsleistung), und nur mit Beimischung von heimischen Laubbaumarten wie Eiche und Buche', style='List Bullet')
        doc.add_paragraph('optimale Nutzung der wirtschaftlichen Möglichkeiten innerhalb des ökologischen Rahmens bei der Baumartenwahl (so viele wirtschaftlich interessante Baumarten wie möglich und ausreichend ökologisch erforderliche Baumarten) sowie', style='List Bullet')
        doc.add_paragraph('Rücksichtnahme auf das Landschaftsbild bei der forstlichen Nutzung der Wälder', style='List Bullet')
        doc.add_paragraph('')

        doc.add_paragraph('Angestrebt werden gesunde, stabile, standortangepasste, gut strukturierte Waldbestände mit wertvollem Holz.')
        doc.add_paragraph('')

        doc.add_paragraph('Dies wird durch folgende Strategien erreicht:')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Wahl standortstauglicher Baumarten').bold = True
        doc.add_paragraph('Baumarten werden gemäß den Bestockungszielen für die Standortseinheiten gewählt. Bestände, die dem Bestockungsziel nicht entsprechen, werden sukzessiv umgewandelt.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Berücksichtigung der Klimaänderung').bold = True
        doc.add_paragraph('Besonders in den tieferen Lagen wird der zu erwartenden Klimaänderung (Erwärmung) durch verstärkte Verwendung bzw. Förderung von Wärme ertragenden Baumarten Rechnung getragen. Der mit der Klimaänderung zu befürchtenden Erhöhung der Häufigkeit und Intensität von Schadereignissen wird durch den Aufbau stabiler Bestände und erhöhte Sorgfalt bei der Nutzung und Pflege begegnet.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Verfeinerung der Waldstruktur').bold = True
        doc.add_paragraph('Zur Verbesserung der horizontalen Struktur werden große, gleichförmige Waldbestände kleinflächig erneuert. Im schlepperbefahrbaren Gelände kann in geeigneten Beständen die Einzelstammnutzung forciert und durch längere Verjüngungszeiträume eine vertikale Strukturierung angestrebt werden.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Natürliche Bestandeserneuerung').bold = True
        doc.add_paragraph('Bei entsprechendem Verjüngungspotenzial erfolgt die Bestandesbegründung i. d. R. über Naturverjüngung. Davon muss Abstand genommen werden, wenn in Beständen ungeeignete Mutterbäume vorkommen (Baumart, Herkunft etc.) und diese sich verjüngen würden, oder wenn der Standort zur Verunkrautung oder Verwilderung neigt. In solchen Fällen wird aufgeforstet.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Optimale Baumdimensionen').bold = True
        doc.add_paragraph('Ernte und Begründung der Waldbestände erfolgen unter Beachtung der optimalen Baumdimensionen. Der Zieldurchmesser in Brusthöhe liegt in annähernd gleichaltrigen Fichten- und Tannenbeständen bei rund 40 cm ohne Rinde bei schlechten Bonitäten (4 bis 6) bzw. 50 cm ohne Rinde auf besseren Bonitäten (ab 7), bei Laubbaumarten oder Lärche mit hoher Qualität bei 60 cm und mehr.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Angepasste Nutzungsgrößen').bold = True
        doc.add_paragraph('Große, zusammenhängende Nutzungen, die flächige Kahllegungen nach sich ziehen, werden vermieden.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Frühzeitige Pflegeeingriffe').bold = True
        doc.add_paragraph('Dickungspflege und Erstdurchforstung erfolgen rechtzeitig, um das Zuwachspotenzial optimal zu nutzen. Bei Nadelbaumarten werden aus Qualitätsgründen Jahrringbreiten von 3 bis 4 mm angestrebt, was auf besseren Standorten mit einer Stammzahlreduktion bzw. Mischwuchsregulierung und frühzeitiger Erstdurchforstung möglich ist.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Vorsicht bei der Durchforstung von älteren Beständen').bold = True
        doc.add_paragraph('Bei Durchforstungen in älteren Beständen unterbleiben starke Auflockerungen.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Intensivere Nutzung guter Standorte').bold = True
        doc.add_paragraph('Bestände auf guten Standorten werden intensiver bewirtschaftet als solche auf ertragsschwachen. Auf schlechten Standorten sollen die Bestände mit möglichst geringem Aufwand so aufwachsen, dass sie den ökologischen Anforderungen gerecht werden.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Vermeidung von Ernteschäden').bold = True
        doc.add_paragraph('Durch optimierte Erntetechnik werden Schäden an Verjüngung, Bestand und Boden (Wurzelschäden) vermieden.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Abstimmung des Biomasseentzugs auf den Standort').bold = True
        doc.add_paragraph('Grundsätzlich bleiben Feinäste, Nadeln und Blätter am Fällungsort. Auf Böden mit hohem Nährstoffpotenzial ist die Entnahme dieser Biomasse vertretbar. Hochmechanisierte Verfahren werden auf Standorten mit mittlerer oder geringer Nährstoffversorgung modifiziert, zum Beispiel Abwipfeln oder Grobentastung auf der Oberseite des gefällten Baumes.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Schädlingsvorbeugung').bold = True
        doc.add_paragraph('Schadensprävention und natürlicher Waldschutz durch Waldhygiene und Förderung der Nützlinge haben Vorrang vor Bekämpfungsmaßnahmen, insbesondere vor dem Einsatz von Pestiziden.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Priorität der Waldbewirtschaftung vor der Jagd').bold = True
        doc.add_paragraph('Wo es möglich ist, werden waldbauliche Maßnahmen auf die Bedürfnisse des Wildes bzw. der Jagd abgestimmt, um Wildschäden zu vermeiden und Bejagungsmöglichkeiten zu erhalten oder zu schaffen. Dabei wird eine Optimierung angestrebt, wobei der Wald stets Vorrang hat.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Verminderung von Schäden durch Waldweide').bold = True
        doc.add_paragraph('Die Waldweide wird im Rahmen von Vereinbarungen zur Verminderung der Schäden reduziert.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Integraler Naturschutz').bold = True
        doc.add_paragraph('Der ökologisch orientierte, naturnahe Waldbau entspricht den Anforderungen des Natur- und Landschaftsschutzes. Naturschutzanliegen werden im operativen Handeln berücksichtigt.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Integraler Landschaftsschutz').bold = True
        doc.add_paragraph('Die Auswirkungen waldbaulicher Maßnahmen auf das Landschaftsbild werden beachtet. Großflächige Nutzungen, die weit eingesehen werden können, werden vermieden. Bestandesränder werden wo möglich der Landschaft angepasst und lange, gerade Linien möglichst vermieden. In besonders sensiblen Gebieten (intensiver Tourismus) werden große Kahllegungen unterlassen und stattdessen so genannte „Rauschläge“ bevorzugt oder ein vorhandener Zwischenbestand bis zur Dickungspflege belassen.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Naturwaldreservate').bold = True
        doc.add_paragraph('Im Rahmen der nationalen Programme wirken die ÖBf an der Einrichtung von Naturwaldreservaten und deren Beobachtung sowie der Sicherung der genetischen Vielfalt unseres Waldes durch natürliche Verjüngung mit.')
        doc.add_paragraph('')

        p = doc.add_paragraph()
        p.add_run('Einbeziehung der Biodiversität').bold = True
        doc.add_paragraph('Die Biodiversität wird verstärkt in die Planung und Umsetzung einbezogen, Waldwiesen und wertvolle Biotope werden erhalten. Die Gestaltung der Bestandesränder erfolgt bewusst naturnah unter Förderung vorhandener Sträucher. Die standortsbezogene Baumartenwahl, die Verfeinerung der Waldstruktur, die Vielfalt der Nutzungsformen und die Förderung der Naturverjüngung begünstigen die biologische Vielfalt.')
        doc.add_paragraph('')


    def fuc_txt_wuchsgebiet(self, doc, wuchs):


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 1.1: Kontinentale Kernzone
#-------------------------------------------------------------------------------------------------------------------------------

        if wuchs == '1.1':

            doc.add_heading('Wuchsgebiet 1.1: Kontinentale Kernzone', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Stark vergletscherte Hochgebirgslandschaft mit großer Reliefenergie. Die Kammlagen befinden sich durchwegs um 3000 m bis weit darüber. Tief eingeschnittene Kerbtäler und Trogtäler mit weiten Hochtalböden, ausgedehnte Steilhanglagen kennzeichnen das Wuchsgebiet.')
            doc.add_paragraph('Als Grundgestein findet man vorwiegend saures Kristallin (Paragneis), nur im Oberinntal auch basenreichere Bündner Schiefer. Nördlich des Inn liegt eine schmale Zone dolomitischen Kalkalpins.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Insbesondere auf nährstoffarmem Kristallin herrscht Semipodsol vor (39%).')
            doc.add_paragraph('Die klimatische untere Höhengrenze der Podsolverbreitung liegt wegen Trockenheit relativ hoch. Wegen der Höhenlage des Wuchsgebietes und der sehr hohen Waldgrenze ist dennoch auch Podsol vergleichsweise stark verbreitet (18%). Er tritt hier oft in Verbindung mit mächtigem, aber zoogenem Feinmoderhumus bis in große Höhen auf.')
            doc.add_paragraph('Ranker und magere Braunerde aus saurem Kristallingestein sind relativ wenig verbreitet (8%).')
            doc.add_paragraph('Nährstoffreiche Braunerde auf basenreichem, z.T. karbonathaltigem Kristallin reicht bis in sehr hohe Lagen und ist etwas häufiger (15%).')
            doc.add_paragraph('Karbonatgesteinsböden machen immerhin 14% der Waldfläche aus, vor allem Rendsina und Braunlehm- Rendsina (13%) an den Südhängen zum Inntal.')
            doc.add_paragraph('Ferner treten auf: leichtere, auch karbonathaltige Lockersedimentbraunerden auf Talschotter und Moränenmaterial (5%), Hanggley, Fluß- und Bachauen.')

            doc.add_heading('Natürliche Waldgesellschaften', 4) # how to add a Paragraph to the List Bullet paragraph? add p = doc.add_paragraph('') und p.add_run('...')?
            doc.add_paragraph('Das Wuchsgebiet ist ein Zentrum der Lärchen-Zirbenwälder. Zentralalpine Kiefernwälder und andere Trockenvegetation sind verbreitet.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Eichentrockenwald-Fragmente mit Rotföhre im Inntal.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichtenwald mit Lärche (Leitgesellschaft) in der montanen Stufe vorherrschend; submontan bis tief(-mittel)montan in trockener Ausbildung mit Rotföhre, z.T. auch anthropogen durch Rotföhren- Ersatzgesellschaften vertreten. Auf Silikatstandorten vor allem Hainsimsen-Fichtenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic = True
            p.add_run('), auf Karbonatstandorten Buntreitgras- Fichtenwald (')
            p.add_run('Calamagrostio variae-Piceetum').italic = True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder als Dauergesellschaften an flachgründigen, sonnigen Standorten submontan bis mittel(-hoch)montan sehr stark hervortretend. Rotföhre steigt nach oben ausdünnend bis ca. 1900 m an.')
            p.add_run('Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') über karbonatischem Bergsturzschutt (Tschirgant) und an Dolomit- Steilhängen im Inntal. Silikat-Rotföhrenwald (')
            p.add_run('Vaccinio vitisideae-Pinetum').italic=True
            p.add_run('). Hauhechel-Rotföhrenwald (')
            p.add_run('Ononido-Pinetum').italic=True
            p.add_run(') über Bündner Schiefer im obersten Inntal zwischen Prutz und Pfunds.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die mittel (-hoch)montane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald mit höherem Lärchenanteil und Zirbe. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Homogyno-Piceetum').italic=True
            p.add_run('über Silikat, Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabra-Piceetum').italic=True
            p.add_run('). Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Substraten, z.B. Kössener Schichten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Lärchen-Zirbenwald im Silikatgebiet (')
            p.add_run('Larici-Pinetum cembrae').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Über Karbonaten ersetzen Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe großﬂächig die LärchenZirbenwälder und steigen außerdem an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinab. Silikat-Latschengebüsche mit Rostroter Alpenrose (')
            p.add_run('Rhododendron ferrugineum').italic=True
            p.add_run(') nur lokal an blockreichen Standorten im Waldgrenzbereich.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche), bis in die montane Stufe herabsteigend.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 1.2: Subkontinentale Innenalpen - Westteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '1.2':

            doc.add_heading('Wuchsgebiet 1.2: Subkontinentale Innenalpen - Westteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Die hochalpine, vergletscherte Landschaft ist ähnlich dem Wuchsgebiet 1.1: getreppte Trogtäler und V-Täler mit ausgedehnten, wenig gegliederten Steilﬂanken.')
            doc.add_paragraph('Das Grundgestein hat neben Gneisen jedoch höheren Anteil an basenreichen Silikaten als das Wuchsgebiet 1.1: Kalkschiefer, Kalkphyllit und kristalline Kalke. Die Kalkalpen werden hingegen nur kleinräumig im Westen erfaßt; dazu kommen die Kalke des Brenner-Mesozoikums.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Semipodsol ist mit Abstand am weitesten verbreitet (<50%) mit Schwerpunkt auf nährstoffarmem Kristallin; in Steillagen auch Ranker.')
            doc.add_paragraph('Es handelt sich um ein Hauptverbreitungsgebiet des klimabedingten Podsol, der hier auch auf basenreichem Gestein auftritt (Anteil am Schutzwald allein knapp 30%; fast 1/3 aller Probeflächen der österreichischen Waldbodenzustandsinventur mit Podsol liegen in diesem Wuchsgebiet). Die Höhengrenze zwischen Semipodsol und Podsol auf vergleichbarem Gestein liegt etwas tiefer als im Wuchsraum 1.1. Basenarme Braunerde ist auf tiefere Lagen beschränkt (5%).')
            doc.add_paragraph('Relativ verbreitet (20%) ist hingegen basenreiche Braunerde bis in Hochlagen, Kalkbraunerde auf Kalkglimmerschiefer und Kalk.')
            doc.add_paragraph('Nur untergeordnet ﬁndet man ferner: Rendsina auf Kalkfels und Kalkschotter, Lockersedimentbraunerden auf Moränen und Schotter, Hanggley und Anmoore.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Stieleichen-Waldreste mit Rotföhre, Winterlinde im Inntal (z.B. Stams) und im unteren Wipptal.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichtenwald mit Lärche in der submontanen und montanen Stufe vorherrschend (Leitgesellschaft), lokal mit Beteiligung der Tanne (z.B. im Gschnitztal auf Karbonaten). Auf ärmeren Silikatstandorten vorwiegend Hainsimsen-Fichtenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf reicheren Böden Sauerklee-Fichtenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Oxalido-Piceetum').italic=True
            p.add_run('), auf Karbonatstandorten Buntreitgras-(Tannen-)Fichtenwald (')
            p.add_run('Calamagrostio variae-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder als Dauergesellschaften an ﬂachgründigen, sonnigen Standorten submontan bis hochmontan.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In luftfeuchtem Lokalklima (Grabeneinhang) an frisch-feuchten Hangstandorten lokales Vorkommen von Bergahorn-Bergulmen-Eschenwäldern.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die hochmontane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald mit höherem Lärchenanteil und Zirbe. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') über Silikat, Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum')
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lärchenwald (Larif Karbonatgestein (z.B. BrennerMesozoikum) in der montanen-subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Lärchen-Zirbenwald im Silikatgebiet (')
            p.add_run('Larici-Pinetum cembrae').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Über Karbonaten ersetzen Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe großﬂächig die LärchenZirbenwälder und steigen außerdem an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinab. Silikat-Latschengebüsche (')
            p.add_run('Rhododendro ferruginei-Pinetum prostratae').italic=True
            p.add_run(') an blockreichen Standorten in der subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 1.3: Subkontinentale Innenalpen - Ostteil
#-------------------------------------------------------------------------------------------------------------------------------

        if wuchs == '1.3':

            doc.add_heading('Wuchsgebiet 1.3: Subkontinentale Innenalpen - Ostteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Nach Osten zu kennzeichnet merklich abnehmende Reliefenergie mit niedrigeren Kammlinien und höheren Tallagen das Wuchsgebiet. Nur im Westen ist es vergletschert; im Osten herrschen runde Altlandschaftsformen mit Gipfeln unter 2500 m vor.')
            doc.add_paragraph('Es treten fast ausschließlich Silikatgesteine mit basenarmen (Gneis, Granit, Quarzphyllit, Schiefer) und basenreichen (Kalkglimmerschiefer, basische Vulkanite) Komponenten auf; nur lokal kommen Kalkmarmor und Kalk (Radstädter Tauern) vor.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Ranker ist relativ weit verbreitet.')
            doc.add_paragraph('Häuﬁgster Bodentyp ist Semipodsol (>50%).')
            doc.add_paragraph('Im Wuchsgebiet liegt noch ein Schwerpunkt des klimabedingten Podsol. Er nimmt jedoch gegenüber den westlichen Innenalpen ab, vor allem weil die Waldgrenze tiefer liegt. Der Anteil an Semipodsol und magerer Braunerde nimmt nach Osten entsprechend zu.')
            doc.add_paragraph('Basenreiche Braunerde und Kalkbraunerde sind bis in Hochlagen relativ weit verbreitet (>20%).')
            doc.add_paragraph('Untergeordnet treten auf: Lockersedimentbraunerden auf Moräne und Schotter (ebenfalls häuﬁg basenreich), Hanggley, Hangmoore, Hochmoore, Niedermoore (in Hochtälern).')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Es handelt sich um ein Übergangsgebiet zwischen Fichten-Tannenwald und Fichtenwald als Leitgesellschaft. Durch anthropogene Förderung der Fichte ist die Abgrenzung des natürlichen Tannenanteils schwierig.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichtenwald (Leitgesellschaft) bzw. Fichten-Tannenwald submontan bis hochmontan. Tannenfreier montaner Fichtenwald am Rande des Wuchsgebietes v.a. lokalklimatisch (Frostbeckenlagen) oder edaphisch (anmoorige Standorte, Blockhalden) bedingt. Randlich geringwüchsige Buchen lokal beigemischt. Auf ärmeren Silikatstandorten Hainsimsen-(Tannen-)Fichtenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf reicheren Böden Sauerklee-(Tannen-)Fichtenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder als montane Dauergesellschaften an flachgründigen, sonnigen Standorten nur kleinﬂächig.')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge).')
            p.add_run('In luftfeuchtem Lokalklima (Grabeneinhang) an frisch-feuchten Hangstandorten lokales Vorkommen von Bergahorn-Bergulmen-Eschenwäldern. Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(') mit Rasenschmiele (')
            p.add_run('Deschampsia cespitosa').italic=True
            p.add_run(') tief-mittelmontan (z.B. Stubachtal); Hochstauden-Ahornwald (')
            p.add_run('Ulmo-Aceretum').italic=True
            p.add_run(') mittel-hochmontan (z.B. Gößgraben, Radlgraben bei Gmünd).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald (v.a. ')
            p.add_run('Larici-Piceetum')
            p.add_run(') und hochsubalpiner Lärchen-Zirbenwald (')
            p.add_run('LariciPinetum cembrae')
            p.add_run(') sind noch gut ausgebildet.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Silikat-Latschengebüsche mit Rostroter Alpenrose (')
            p.add_run('Rhododendro ferruginei-Pinetum prostratae')
            p.add_run(') in der subalpinen Stufe gut entwickelt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 2.1: Nördliche Zwischenalpen - Westteil
#-------------------------------------------------------------------------------------------------------------------------------

        if wuchs == '2.1':

            doc.add_heading('Wuchsgebiet 2.1: Nördliche Zwischenalpen - Westteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Die Gipfellagen liegen zwischen 3000 und 2000 m und sinken von Westen nach Osten ab. Nur im Westen ist das Gebiet geringfügig vergletschert. Die Haupttäler verlaufen von West nach Ost.')
            doc.add_paragraph('Dem dominierenden Klimacharakter der Nördlichen Zwischenalpen wurden die recht vielfältigen geochemisch-edaphischen Gegebenheiten untergeordnet: Das Wuchsgebiet umfaßt die Leelagen der Nördlichen Kalkalpen vor allem im Westen in einer breiten Zone, randliche Bereiche der zentralalpinen Gneise, die Innsbrucker Quarzphyllitberge sowie Teile der Kitzbühler Schieferalpen und die Sedimente des Inntales.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Etwa die Hälfte aller Böden liegt auf Silikatgestein.')
            doc.add_paragraph('Auch in diesem Wuchsgebiet herrscht auf Silikat Semipodsol vor (23% des Wuchsgebietes bzw. über 40% der Silikatböden).')
            doc.add_paragraph('Relativ weit (14% bzw. 26% der Silikatböden) und in tieferen Lagen als in den Innenalpen verbreitet ist Podsol - sowohl klimatisch begünstigt als auch wegen des hohen Anteils an basenarmem Gestein (Quarzphyllit).')
            doc.add_paragraph('Auch magere Braunerde ist im Silikatgebiet vergleichsweise stärker vertreten (insgesamt 5%), während basenreiche Braunerde zurücktritt (7%).')
            doc.add_paragraph('Auf silikatischem Substrat ferner Ranker sowie Braunerde auf Moräne, Terrassenschottern etc.')
            doc.add_paragraph('Ein relativ großer Teil des Wuchsgebietes fällt auf Kalkböden (43%) mit Rendsina (13%), BraunlehmRendsina (18%) und Kalkbraunlehm (12%) sowie etwas Kalkbraunerde.')
            doc.add_paragraph('Ferner kommen vor: Hanggley (4%), Pseudogley auf Lockersedimenten (Terrassen) und tonigem Festgestein.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Eichenmischwald-Fragmente mit Stieleiche, Rotföhre und Winterlinde (z.B. Ampass). Bei Innsbruck und Zirl an wärmebegünstigten Stellen (Föhn) isolierte Vorkommen von Hopfenbuche und Blumenesche.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald (Leitgesellschaft) in der submontanen und montanen Stufe, häuﬁg anthropogen durch Fichten-Ersatzgesellschaften vertreten. Der Kitzbüheler Raum (Brixental) ist besonders tannenreich. Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run('). Auf Karbonat Buntreitgras-Fichten-Tannenwald (')
            p.add_run('Calamagrostio variae-Piceetum').italic=True
            p.add_run(', trockener) und Alpendost Fichten-Tannenwald (')
            p.add_run('Adenostylo glabrae-Abietetum').italic=True
            p.add_run(', frischer).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tannenfreier montaner Fichtenwald auf lokalklimatisch (Frostbeckenlagen) oder edaphisch (anmoorige Standorte, Blockhalden) bedingten Sonderstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontan und tiefmontan auf warmen, gut durchlüfteten Karbonatstandorten (“laubbaumfördernde Unterlage”) verstärkter Buchenanteil (Fichten-Tannen-Buchenwald).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(' mit Schneeheide,')
            p.add_run(' Carici humilis-Pinetum sylvestris').italic=True
            p.add_run(' mit Erdsegge, extremere Standorte) als Dauergesellschaften an flachgründigen, sonnigen DolomitSteilhängen submontan bis mittelmontan besonders im Inntal häuﬁg auftretend.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Spirkenwald als Dauergesellschaft an schattigen Steilhängen (')
            p.add_run('Rhododendro hirsuti-Pinetum montanae').italic=True
            p.add_run(' mit Wimper-Alpenrose auf Dolomit, ')
            p.add_run('Lycopodio annotini-Pinetum uncinatae').italic=True
            p.add_run(' mit Torfmoos auf Bergsturzschutt) oder als Pionier- bzw. anthropogene Degradationsgesellschaft (')
            p.add_run('Erico carneaePinetum uncinatae').italic=True
            p.add_run(') auf sonnigen (Schutt-)Standorten mit Rotföhre, Steinröslein (')
            p.add_run('Daphne striata').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchte und Bergulme (z.B. ')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(' mit Wald-Ziest und Rasenschmiele, ')
            p.add_run('Lunario-Aceretum').italic=True
            p.add_run(' mit Mondviole,')
            p.add_run(' Arunco-Aceretum').italic=True
            p.add_run(' mit Geißbart).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die hochmontane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald gut entwickelt. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') über Silikat. Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run('). Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der (tief-)hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinabreichend. Silikat-Latschengebüsche (')
            p.add_run('Rhododendro ferruginei-Pinetum prostratae').italic=True
            p.add_run(') mit Rostroter Alpenrose auf skelettreichen Böden in der subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Silikat-Lärchen-Zirbenwald (')
            p.add_run('Larici-Pinetum cembrae').italic=True
            p.add_run(') nur kleinﬂächig, gebietsweise auch fehlend (Kitzbüheler Alpen: ausgedehnte Almgebiete). Karbonat-Lärchen-Zirbenwald (')
            p.add_run('Pinetum cembrae').italic=True
            p.add_run(') und Karbonat-Lärchenwald (')
            p.add_run('Laricetum deciduae').italic=True
            p.add_run(') sind kleinﬂächig vorhanden.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 2.2: Nördliche Zwischenalpen - Ostteil
#-------------------------------------------------------------------------------------------------------------------------------

        if wuchs == '2.2':

            doc.add_heading('Wuchsgebiet 2.2: Nördliche Zwischenalpen - Ostteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Die Landschaft entlang des alpinen Längstales Salzach-Enns besteht vornehmlich aus bodensauren Quarzphylliten, Quarziten und Glimmerschiefern der Grauwackenzone (Salzburger Schieferberge) und der Niederen Tauern. Im Pongau gibt es auch Kalkglimmerschiefer sowie (kristalline) Kalke und Dolomit. Die Kammlinien liegen in den Salzburger Schieferbergen nur um 2000 m, in den Niederen Tauern (z.T. außerhalb des Wuchsgebietes) um 2400 (bis 2800) m.')
            doc.add_paragraph('Teilweise werden von dem Wuchsgebiet noch die Südhänge der nördlichen Kalkalpen erfaßt, der Flächenanteil ist aber geringer als in den westlichen Zwischenalpen (Wuchsgebiet 2.1). Hier werden auch die größten Gipfelhöhen (Dachstein 2995 m) erreicht. Verbreitet sind erosionsgefährdete Steilhänge aus mürbem, tiefgründig aufgewittertem Dolomit.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Auf Silikatgestein dominiert wiederum Semipodsol (ca. 40%), gefolgt von reicher Braunerde (20%), welche hier etwas häufiger auf basenreichem Substrat bis in Hochlagen vorkommt.')
            doc.add_paragraph('Basenarme Braunerde (10%) ist relativ weniger und nur in Talnähe verbreitet.')
            doc.add_paragraph('Die klimatischen Verbreitungsbedingungen des Podsols rücken auf vergleichbarem Substrat von Westen nach Osten in größere Höhe, gleichzeitig sinkt die durchschnittliche Gipfelhöhe nach Osten zu ab. Die klimabedingte Podsolzone ist deshalb vergleichsweise schmal. Anderseits begünstigt das bodensaure Substrat v.a. am Nordabfall der Niederen Tauern die Podsolverbreitung bis in Tallagen. Insgesamt ist Podsol weniger häufig (ca. 10%) als im westlichen Wuchsgebiet.')
            doc.add_paragraph('Die kalkalpinen Südhänge und zentralalpinen Marmorzüge machen etwas über 25% der Waldfläche aus. Mehr als ein Drittel davon sind Extremstandorte mit Dolomitrendsina. Auf Kalk überwiegen Braunlehmrendsina und Kalkbraunlehm.')
            doc.add_paragraph('Relativ häufig ist weiters Hanggley und Pseudogley (4%), v.a. auf Gosau und Werfener Schichten, untergeordnet ferner bindige Braunerde auf den Lockersedimenten der Haupttäler.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Zwischenalpines Fichten-Tannenwaldgebiet. An begünstigten Stellen kommt Buche vor. An lokalklimatischen und edaphischen Sonderstandorten gibt es noch Zirbenvorkommen (Dachsteinplateau).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Eichenmischwald-Fragmente kleinflächig.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald (Leitgesellschaft) in der submontanen und montanen Stufe, häufig anthropogen durch Fichten-Ersatzgesellschaften vertreten. Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic = True
            p.add_run(', auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Oxalido-Abietetum').italic=True
            p.add_run('). Karbonat-Alpendost-Fichten-Tannenwald (')
            p.add_run('Adenostylo glabrae-Abietetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tannenfreier montaner Fichtenwald auf lokalklimatisch (Frostbeckenlagen) oder edaphisch (anmoorige Standorte, Blockhalden) bedingten Sonderstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontan und Tief(-mittel)-montan auf warmen, gut durchlüfteten Karbonatstandorten (“laubbaumfördernde Unterlage”) verstärkter Buchenanteil (Fichten-Tannen-Buchenwald).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schneeheide-Rotföhrenwälder (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') als Dauergesellschaften an ﬂachgründigen, sonnigen Dolomit-Steilhängen submontan bis mittelmontan kleinﬂächig auftretend.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima Laubmischwälder mit Bergahorn, Esche und Bergulme (z.B. ')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die hochmontane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald gut entwickelt. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Homogyno-Piceetum').italic=True
            p.add_run(') über Silikat und subalpiner Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinabreichend. Silikat-Latschengebüsche (')
            p.add_run('Rhododendro ferruginei-Pinetum prostratae').italic=True
            p.add_run(') mit Rostroter Alpenrose auf skelettreichen Böden in der subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Silikat-Lärchen-Zirbenwald (')
            p.add_run('Larici-Pinetum cembrae').italic=True
            p.add_run(') an Sonderstandorten, gebietsweise (Kitzbüheler Alpen) fehlend. KarbonatLärchen-Zirbenwald (')
            p.add_run('Pinetum cembrae').italic=True
            p.add_run(') und Karbonat-Lärchenwald (')
            p.add_run('Laricetum deciduae').italic=True
            p.add_run(') sind kleinﬂächig vorhanden.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 3.1: Östliche Zwischenalpen - Nordteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '3.1':

            doc.add_heading('Wuchsgebiet 3.1: Östliche Zwischenalpen - Nordteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Entlang der dominierenden Liesing-Mur-Mürzfurche liegen nur kleinere hochalpine Bergmassive; die Kammlagen sind meist unter 2000 m.')
            doc.add_paragraph('Vor allem der Osten ist periglazialer Raum mit Resten einer alten Rumpﬂandschaft. Das Gebiet ist geologisch sehr vielfältig, weist jedoch vorwiegend basenarme Gesteine wie Ortho- und Paragneise, Quarzphyllite und Quarzite sowie saure Ergußgesteine auf. Nur zum kleinen Teil kommen paläozoische (Eisenerzer Alpen) und andere (Veitsch, Semmeringtrias) Kalke/Dolomite vor. Die tertiären Beckenfüllungen sind vornehmlich landwirtschaftlich genutzt.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Magere, podsolige Braunerde und Semipodsol (zusammen über 60%) auf intermediärem oder basenarmem Silikat herrschen vor.')
            doc.add_paragraph('Die Zone klimabedingten Podsols wird nur mehr in den höchsten Lagen erreicht. Verbreiteter tritt Podsol aber höhenunabhängig auf sehr quarzreichem Schiefer, Quarz-Phyllit, Quarzit etc. auf (Podsol insgesamt in Wuchsgebiet 2.2 und 3.1 10%).')
            doc.add_paragraph('Braunerde auf Amphibolit und anderem basenreicherem Silikatgestein reicht bis in große Höhen.')
            doc.add_paragraph('Ferner treten auf: Hanggley, Pseudogley; der Anteil an Rendsina und Braunlehm-Rendsina in den kalkalpinen Randgebieten ist gering.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Das Wuchsgebiet ist Verbreitungsgebiet der natürlichen Fichten-Tannenwälder mit Buche und Lärche. An begünstigten Stellen (Kalk) ist die Buche auch bestandsbildend; Zirbe fehlt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Eichen-Rotföhrenwald-Fragmente (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run('), z.B. bei Leoben.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald (Leitgesellschaft) mit Lärche, Buche und Bergahorn in der submontanen und montanen Stufe, häuﬁg anthropogen durch Fichten-Ersatzgesellschaften vertreten. In den submontanen bis mittelmontanen Ausbildungen mit Rotföhre und stärkerer Beimischung von Buche; Bergahorn an feuchteren Standorten. In den hochmontanen Homogyne-Ausbildungen Tanne zurücktretend, Buche nur mehr auf karbonatischen Böden im Nebenbestand. Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run('). Karbonat-Alpendost-Fichten-Tannenwald (')
            p.add_run('Adenostylo glabrae-Abietetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tannenfreier montaner Fichtenwald auf lokalklimatisch (Frostbeckenlagen) oder edaphisch (anmoorige Standorte, Blockhalden) bedingten Sonderstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf Karbonatstandorten (“laubbaumfördernde Unterlage”) und in der submontanen bis tiefmontanen Stufe auch Fichten-Tannen-Buchenwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Silikat-Rotföhrenwald (')
            p.add_run('Vaccinio vitis-idaeae-Pinetum').italic=True
            p.add_run(') kleinﬂächig als montane Dauergesellschaften an ﬂachgründigen, sonnigen Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die hochmontane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Homogyno-Piceetum').italic=True
            p.add_run(') über Silikat. Subalpiner Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run('). Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Böden.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe, über ﬂachgründigen Karbonatböden sowie an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) in die montane Stufe hinabreichend. Silikat-Latschengebüsche (')
            p.add_run('Rhododendro ferruginei-Pinetum prostratae').italic=True
            p.add_run(') mit Rostroter Alpenrose beschränken sich im wesentlichen auf skelettreiche Böden in der subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 3.2: Östliche Zwischenalpen - Südteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '3.2':

            doc.add_heading('Wuchsgebiet 3.2: Östliche Zwischenalpen - Südteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Es handelt sich um Hochgebirge mit Gipfelﬂuren wenig über 2000 m, mit weiten, offenen Tälern und mäßig steilen Hängen. Mit Ausnahme der Seetaler Alpen und der Niederen Tauern besteht das Gelände aus ﬂachen Bergrücken und Kuppen. Es kommt fast ausschließlich Silikatgestein vor: basenarme Gneise mit Marmor- und Amphibolitzügen sowie Quarzphyllit. Im Raum Neumarkter Sattel ﬁndet man auch paläozoischen Kalk und metamorphe basische Ergußgesteine. In weiten Talbecken gibt es tertiäre Sedimente.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Hier herrscht die Braunerde-Podsolreihe auf Kristallin vor. Kalkbeeinﬂußte Böden fehlen fast vollständig.')
            doc.add_paragraph('Am weitesten verbreitet ist Semipodsol (55%*). Auf basenarmem Kristallin reicht er einerseits bis in tiefe Lagen, anderseits bis etwa 1200 m, an Sonnhängen bis über 1500 m.')
            doc.add_paragraph('Die tief gelegenen Täler erlauben dennoch eine gewisse Verbreitung von Braunerde auf saurem Substrat.')
            doc.add_paragraph('Die klimatische Höhenzone des Podsol ist nur schmal und/oder an sehr saures Substrat (Quarzitgänge etc.) gebunden (zusammen ca. 5%* der Waldﬂäche).')
            doc.add_paragraph('Auf basenreichem Kristallin ist nährstoffreiche Braunerde weit verbreitet (>20%*), die Höhengrenze zum Semipodsol liegt dort sehr hoch: Auf Amphibolit beginnt Semipodsol erst in Kammlagen gegen 1800 m und somit an oder über der Waldgrenze.')
            doc.add_paragraph('Ferner treten auf: Anmoore, Hanggley und Karbonatböden (jeweils unter 2%).')

            doc.add_heading('Natürliche Waldgesellschaften', 4)

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Eichen-Rotföhrenwald-Fragmente.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald (Leitgesellschaft) mit Lärche und Buche in der submontanen und montanen Stufe. Tanne heute jedoch weitgehend aus den anthropogenen Fichten-Ersatzgesellschaften verdrängt. In den tief-mittelmontanen Ausbildungen mit Rotföhre und stärkerer Beimischung von Buche, in den hochmontanen Alpenlattich-(Homogyne-)Ausbildungen Tanne zurücktretend. Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run('). Karbonat-Alpendost-Fichten-Tannenwald (')
            p.add_run('Adenostylo glabrae-Abietetum').italic=True
            p.add_run(') nur lokal.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tannenfreier montaner Fichtenwald auf lokalklimatisch (Frostbeckenlagen) oder edaphisch (anmoorige Standorte, Blockhalden) bedingten Sonderstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf Karbonatstandorten (“laubbaumfördernde Unterlage”, z.B. bei Unzmarkt) und in der submontanen bis tief(-mittel)montanen Stufe auch FichtenTannen-Buchenwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Silikat-Rotföhrenwald (')
            p.add_run('Vaccinio vitis-idaeae-Pinetum').italic=True
            p.add_run(') kleinﬂächig als montane Dauergesellschaften an ﬂachgründigen, sonnigen Standorten. Auf Serpentinit bei Kraubath auch Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die hochmontane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In luftfeuchtem Lokalklima an nährstoffreichen Unterhängen Laubmischwälder mit Bergahorn und Esche (lokal).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald. V.a. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Homogyno-Piceetum').italic=True
            p.add_run(') über Silikat, auch Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Böden.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Lärchen-Zirbenwald nur lokal (z.B. Zirbitzkogel).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Silikat-Latschengebüsche (').italic=True
            p.add_run('Rhododendro ferruginei-Pinetum prostratae').italic=True
            p.add_run(') mit Rostroter Alpenrose auf skelettreichen Böden in der subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 3.3: Südliche Zwischenalpen
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '3.3':

            doc.add_heading('Wuchsgebiet 3.3: Südliche Zwischenalpen', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Vorwiegend hochalpines Gebiet an der Südabdachung der Zentralalpen mit Kammlinien bis 3000 m, breiten, tief ausgeschürften Trogtälern bis 500 m Seehöhe herab und V-Gräben mit steilen Flanken.')
            doc.add_paragraph('Ziemlich einheitliche Schiefergneise und Glimmerschiefer herrschen vor. Im Süden treten Triasdolomite/-kalke der Lienzer Dolomiten und Gailtaler Alpen sowie paläozoische Schiefer der Karnischen Alpen auf.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Abgesehen von inneralpin orientierten N-Hängen der Lienzer Dolomiten und Gailtaler Alpen sowie lokalen Marmorzügen dominieren die Böden der Felsbraunerde-Podsol-Reihe. Konkrete Daten der Forstinventur können für dieses Wuchsgebiet nicht abgeleitet werden. Abgesehen von den etwas stärker vertretenen Karbonatgesteinsböden ist aber die Verteilung der Bodenformen jener des Wuchsgebiets 3.2 ähnlich. Infolge der größeren Massenhebung dürfte die Verbreitung von Braunerde auf saurem Substrat etwas geringer sein.')
            doc.add_paragraph('Das steilere Relief bedingt das häuﬁge Vorkommen von Ranker unter Wald.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            p = doc.add_paragraph('Durch das vorgeschobene Vorkommen von Blumenesche, Hopfenbuche und Dreiblatt-Windröschen (')
            p.add_run('Anemone trifolia').italic=True
            p.add_run(') in den Tallagen (z.B. Drautal) wird in diesem Wuchsgebiet bereits ein stärkerer submediterran-illyrischer Einﬂuß spürbar. Die Höhenstufengrenzen sind gegenüber dem Wuchsgebiet 3.2 deutlich (ca. 100-150 m) nach oben verschoben.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontane Eichen-Rotföhrenwald-Fragmente und submontan-tiefmontane Vorposten von Hopfenbuchen-Blumeneschenwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald (Leitgesellschaft) in der submontanen und montanen Stufe, häuﬁg anthropogen an Tanne verarmt. In den submontanen bis mittelmontanen Ausbildungen mit stärkerer Beimischung von Buche. Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run('). Auf Karbonat z.B. Alpendost-FichtenTannenwald (')
            p.add_run('Adenostylo glabrae-Abietetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tannenfreier montaner Fichtenwald auf lokalklimatisch (Frostbeckenlagen) oder edaphisch (anmoorige Standorte, Blockhalden) bedingten Sonderstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf Karbonatstandorten (“laubbaumfördernde Unterlage”) und in der submontanen bis tief(-mittel)montanen Stufe auch Fichten-Tannen-Buchenwald. Z.B. Dreiblatt-Windröschen-Fichten-Tannen-Buchenwald (')
            p.add_run('Anemono trifoliae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf Karbonat, Hainsimsen-Fichten-Tannen-Buchenwald (')
            p.add_run('Luzulo nemorosae(Abieti-)Fagetum').italic=True
            p.add_run(') auf Silikat.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder als submontane bis mittel(-hoch)montane Dauergesellschaften an flachgründigen, trockenen Standorten. Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') über Karbonat und Silikat-Rotföhrenwald (')
            p.add_run('Vaccinio vitis-idaeae Pinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald und an feuchten Hängen (z.B. Muren, Lawinenzüge) von der submontanen bis in die hochmontane Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald. V.a. Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') über Silikat, auch Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Silikat-Lärchen-Zirbenwald (')
            p.add_run('Larici-Pinetum cembrae').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinabreichend. Silikat-Latschengebüsche mit Rostroter Alpenrose (')
            p.add_run('Rhododendron ferrugineum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 4.1: Nördliche Randalpen - Westteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '4.1':

            doc.add_heading('Wuchsgebiet 4.1: Nördliche Randalpen - Westteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Im Bregenzer Wald reicht eine ausgedehnte Zone mit helvetischen und Flyschgesteinen bis in subalpine Lagen;  häuﬁg tritt Mergel bis mergeliger Kalk auf. Die Landschaft hat Mittelgebirgscharakter mit weiten, ungegliederten Hängen; sie läuft im Westen ins MolasseHügelland aus. Das Gebiet ist außerordentlich stark entwaldet; der Wald ist auf die Grabeneinhänge konzentriert. (Das wenig bewaldete Rheintal ist dem Wuchsgebiet zugeordnet).')
            doc.add_paragraph('Die Kalkalpen-Hauptkette hat Gipfelhöhen zwischen 2000 und 3000 m und tief eingeschnittene Täler. Sie ist fast ausschließlich aus (Trias-) Karbonatgesteinen aufgebaut; im Westen herrscht Dolomit vor, ab Salzburg Kalk. Im Westen liegen Kammgebirge, ab Leoganger und Loferer Steinberge ostwärts Karsthochﬂächen mit steilen Felsﬂanken. Um Werfen und Abtenau, am Fuß des Hochkönigs und im Ausseerland beherrschen sanftere Formen aus leichter verwitterbaren Werfener und Gosauschichten die Landschaft.')
            doc.add_paragraph('Ab Salzburg ostwärts breitet sich wieder eine vorgelagerte schmale Flyschzone mit Mittelgebirgscharakter aus, bestehend aus Mergel und Sandstein.')


            doc.add_heading('Böden', 4)
            doc.add_paragraph('Das Wuchsgebiet umfaßt insgesamt 16% Pseudogleyund Gleyböden sowie 55% Böden auf Karbonatgestein. Innerhalb der Flyschzone überwiegen Pseudogley (51%) und Hanggley (4%), sowie bindige, z.T. kalkhaltige Braunerde (8%) und braunlehmartige Böden (21%) aus Mergel; seltener auf Sandstein saure, z.T podsolige Braunerde (insgesamt ca. 5%).')
            doc.add_paragraph('In den Kalkalpen dominieren Rendsina und Braunlehm-Rendsina (zusammen 63%) und Kalkbraunlehm (24%).')
            doc.add_paragraph('Auf Geschiebelehm (Moränen, etc.), Tertiär, Werfener Schichten etc. tritt auch hier Pseudogley (5%) sowie basenreiche, z.T. kalkhältige Braunerde (4%) auf.')
            doc.add_paragraph('Vor allem in Tallagen Niedermoore, Anmoore. Nur ganz vereinzelt podsolige Braunerde auf Silikatgestein.')


            doc.add_heading('Natürliche Waldgesellschaften', 4)
            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontaner Stieleichen-Hainbuchenwald (')
            p.add_run('Galio sylvatici-Carpinetum').italic=True
            p.add_run(') an wärmebegünstigten Hängen am Alpenrand.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Beimischung von Tanne (auf Flyschpseudogley stärker), Bergahorn, Esche (Fichte). Fichten-Tannen-Buchenwald (Leitgesellschaft) mit Quirl-Weißwurz (')
            p.add_run('Polygonatum verticillatum').italic=True
            p.add_run(') in der mittel- bis hochmontanen Stufe. Häuﬁg anthropogene Entmischung  zu Fichte-Tanne bzw. Fichte-Buche oder zu Fichten- bzw. Buchen-Reinbeständen. Auf Karbonatgesteinen Hainsalat-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Aposerido-(Abieti-)Fagetum').italic=True
            p.add_run(') vorherrschend, mittelmontan mit Grünem Alpendost (')
            p.add_run('Adenostyles glabra').italic=True
            p.add_run('), hochmontan außerdem mit Rostsegge (')
            p.add_run('Carex ferruginea').italic=True
            p.add_run('), von Salzburg nach Osten in den Schneerosen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Helleboro(Abieti-)Fagetum').italic=True
            p.add_run(') übergehend. Weißseggen-Buchenwald (')
            p.add_run('Carici albae-Fagetum').italic=True
            p.add_run(') submontan bis tiefmontan auf trockeneren Karbonatstandorten, Bergahorn-Buchenwald (')
            p.add_run('Aceri-Fagetum').italic=True
            p.add_run(') hochmontan in sehr schneereichen, aber frostgeschützten Lagen. Waldmeister-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf leichter verwitternden, basenreichen Substraten (z.B. Flysch), Hainsimsen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montaner Fichten-Tannenwald als edaphisch bedingte Dauergesellschaft, submontan bis tiefmontan z.T. mit Stieleiche gemischt. Peitschenmoos-Tannen-Fichtenwald (')
            p.add_run('Mastigobryo-Piceetum').italic=True
            p.add_run(') mit Torfmoos auf anmoorigen Standorten oder Waldschachtelhalm-Fichten-Tannenwald (')
            p.add_run('Equiseto sylvatici-Abietetum').italic=True
            p.add_run(') auf Gleystandorten an vernäßten, tonreichen Flachhängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montaner Fichtenwald als lokalklimatisch (Kaltluftdolinen) oder edaphisch bedingte Dauergesellschaft. Kalk-Block-Fichtenwald (')
            p.add_run('Asplenio-Piceetum').italic=True
            p.add_run(') auf Blockhalden. Kalkfels-Fichtenwald (')
            p.add_run('Carici albae-Piceetum').italic=True
            p.add_run(') an ﬂachgründigen Felshängen. Torfmoos-Fichtenwald (')
            p.add_run('Sphagno girgensohnii-Piceetum').italic=True
            p.add_run(') an Moorrändern.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') kleinflächig als Dauergesellschaft an flachgründigen, sonnigen Dolomit-Steilhängen submontan bis mittelmontan auftretend.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Spirkenwald (z.B. ')
            p.add_run('Rhododendro hirsuti-Pinetum montanae').italic=True
            p.add_run(') an schattigen Dolomit-Steilhängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') als Auwald vorherrschend, an den größeren Flüssen (z.B. Rheintal) auch Silberweiden-Au (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') und Hartholz-Au mit Esche')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima Laubmischwälder mit Bergahorn, Esche und Bergulme. Submontan bis mittelmontan Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(') mit Waldziest und Rasenschmiele auf wasserzügigen Unterhängen; auf skelettreicheren Schluchtstandorten Hirschzungen-Ahornwald (')
            p.add_run('Scolopendrio-Fraxinetum').italic=True
            p.add_run('), Mondviolen-Ahornwald (')
            p.add_run('Lunario-Aceretum').italic=True
            p.add_run(') und GeißbartAhornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run('). Hochstauden-Ahornwald (')
            p.add_run('Ulmo-Aceretum').italic=True
            p.add_run(') mit Grauem Alpendost (')
            p.add_run('Adenostyles alliariae').italic=True
            p.add_run(') und Alpen-Milchlattich (')
            p.add_run('Cicerbita alpina').italic=True
            p.add_run(') (mittel-)hochmontan.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lindenmischwald mit Sommerlinde submontan-tiefmontan auf trockeneren kalkreichen Schutthängen. Kalkschutthalden-Lindenwald (')
            p.add_run('Cynancho-Tilietum').italic=True
            p.add_run(') weiter verbreitet. Turinermeister-Lindenwald (')
            p.add_run('Asperulo taurinae-Tilietum').italic=True
            p.add_run(') submontan an wärmebegünstigten Hängen (Föhn!) in Vorarlberg.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald als schmaler Höhengürtel, reichlich mit Lärchen gemischt. Überwiegend Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(') über skelettreichen Karbonatböden. Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Substraten, seltener Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') auf bodensauren Standorten (z.B. Tangelhumus).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Lärchenwald (')
            p.add_run('Laricetum deciduae').italic=True
            p.add_run(') kleinﬂächig in der subalpinen Stufe, an schattigen Steilhängen bis ca. 800 m hinabsteigend.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpiner Karbonat-Lärchen-Zirbenwald (')
            p.add_run('Pinetum cembrae').italic=True
            p.add_run(') nur fragmentarisch.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinabreichend, häuﬁg anthropogen gefördert.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet Wuchsgebiet 4.2: Nördliche Randalpen - Ostteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '4.2':

            doc.add_heading('Wuchsgebiet 4.2: Nördliche Randalpen - Ostteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Hochgebirge besteht fast ausschließlich aus Kalk und Dolomit. Es weist ausgedehnte Karsthochﬂächen (Altlandschaften) mit steilen Felsﬂanken, tief eingeschnittenen Tälern und Schluchten auf. Die Gipfelfluren liegen wenig über 2000 m und sinken nach Osten zu ab. Die im Nordosten vorgelagerte Kette der Kalkvoralpen bildet eher Kämme und erreicht nur um 1700 m, im Osten bis 1300 m.')
            doc.add_paragraph('Am Nordrand liegt ein schmales, nach Osten (Wienerwald) zu breiter werdendes Band aus Flyschgesteinen mit runden Formen. Es handelt sich um Mittelgebirge mit Gipfeln unter 1500 m, im Osten unter 900 m.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Der Anteil der Flyschzone an der Waldfläche des Wuchsgebietes beträgt knapp 20%. Die für Flysch und Werfener Schichten typischen schweren Pseudogley- und Hanggley-Böden machen 14% aus. Karbonatböden nehmen einen Anteil von 73% ein.')
            doc.add_paragraph('In der Flyschzone dominiert wiederum Pseudogley und Gley (59%) - etwas mehr als im westlichen Wuchsgebiet; kalkbraunlehmartige Böden treten demgegenüber deutlich zurück (4%!); untergeordnet wie dort sind Rendsina/Pararendsina mit ca. 5%.')
            doc.add_paragraph('Dafür sind silikatische, saure Braunerden mit 15% häuﬁger. Auf Greifensteiner Sandstein auch sandige, podsolige Braunerde. Podsol ist in diesen Höhenlagen auffällig, aber insgesamt selten (1%).')
            doc.add_paragraph('Vor allem im Wienerwald verbreitet sind sehr schwere, alte Bodenbildungen mit sehr tieﬂiegendem Stauhorizont und leichterem Oberboden, der zu oberﬂächlicher Austrocknung neigt.')
            doc.add_paragraph('Die Kalkalpen werden fast ausschließlich von Kalkböden beherrscht, mit einer stärkeren Dominanz von Rendsina (39%) und Braunlehm-Rendsina (29%) als in den westlichen Kalkalpen; Kalkbraunlehm 20%. Auf unreinem Kalk und Dolomit auch Kalkbraunerde (4%). Immerhin nehmen auch hier Pseudogley (Werfener Schichten, Gosau) und Hanggley etwa 9000 ha Waldﬂäche ein.')
            doc.add_paragraph('Der Anteil an saurer Braunerde und Semipodsol auf Silikatgestein (Lunzer Schichten etc.) ist mit 3% gering.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Typisches Fichten-Tannen-Buchenwaldgebiet. Gegenüber dem Wuchsgebiet 4.1 ist ein verstärktes Auftreten von Rotföhrenwäldern auf Dolomit zu beobachten. Die östliche Grenze des Wuchsgebietes wird von der Verbreitungsgrenze der Tanne in der tief-/submontanen Stufe festgelegt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontaner Stieleichen-Hainbuchenwald (')
            p.add_run('Galio sylvatici-Carpinetum').italic=True
            p.add_run(') an wärmebegünstigten Hängen v.a. am Alpenrand.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Beimischung von Tanne (auf Flyschpseudogley stärker), Bergahorn, Esche (Fichte, Rotföhre, Eiche). Fichten-Tannen-Buchenwald (Leitgesellschaft) mit Quirl-Weißwurz (')
            p.add_run('Polygonatum verticillatum').italic=True
            p.add_run(') in der mittel- bis hochmontanen Stufe. Häuﬁg anthropogene Entmischung  zu FichteTanne bzw. Fichte-Buche oder zu Fichten- bzw. Buchen-Reinbeständen. Auf Karbonatgesteinen Schneerosen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Helleboro nigri-(Abieti-)Fagetum').italic=True
            p.add_run(') vorherrschend, mittelmontan mit Grünem Alpendost (')
            p.add_run('Adenostyles glabra').italic=True
            p.add_run('), hochmontan außerdem mit Rostsegge (')
            p.add_run('Carex ferruginea').italic=True
            p.add_run(') und Großer Hainsimse (')
            p.add_run('Luzula sylvatica').ialic=True
            p.add_run('). Weißseggen-Buchenwald (')
            p.add_run('Carici albae-Fagetum').italic=True
            p.add_run(') submontan bis tiefmontan auf trockeneren Karbonatstandorten. Bergahorn-Buchenwald (')
            p.add_run('Aceri-Fagetum').italic=True
            p.add_run(') hochmontan in schneereichen, aber frostgeschützten Lagen. Waldmeister-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf leichter verwitternden, basenreichen Substraten (z.B. Flysch), Hainsimsen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montaner Fichten-Tannenwald als edaphisch bedingte Dauergesellschaft, submontan bis tiefmontan z.T. mit Stieleiche gemischt. Z.B. Waldschachtelhalm-Fichten-Tannenwald (')
            p.add_run('Equiseto sylvatici-Abietetum').italic=True
            p.add_run(') auf Gleystandorten an vernäßten, tonreichen Flachhängen mit Übergängen zu Erlenbeständen (')
            p.add_run('Carici remotae-Fraxinetum ').italic=True
            p.add_run('s.lat.).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montaner Fichtenwald als lokalklimatisch (Kaltluftdolinen) oder edaphisch bedingte Dauergesellschaft. Kalk-Block-Fichtenwald (')
            p.add_run('Asplenio-Piceetum').italic=True
            p.add_run(') auf Blockhalden. Kalkfels-Fichtenwald (')
            p.add_run('Carici albae-Piceetum').italic=True
            p.add_run(') an flachgründigen Felshängen. Torfmoos-Fichtenwald (')
            p.add_run('Sphagno girgensohnii-Piceetum').italic=True
            p.add_run(') an Moorrändern.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') als Dauergesellschaft an ﬂachgründigen, sonnigen Dolomit-Steilhängen submontan bis mittelmontan häuﬁg auftretend.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run('), an den größeren Flüssen auch Silberweidenbestände (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') als Auwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima Laubmischwälder mit Bergahorn, Esche und Bergulme submontan bis mittelmontan. Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(') mit Waldziest und Rasenschmiele auf wasserzügigen Unterhängen; auf skelettreicheren Schluchtstandorten Hirschzungen-Ahornwald (')
            p.add_run('Scolopendrio-Fraxinetum').italic=True
            p.add_run('), Mondviolen-Ahornwald (')
            p.add_run('Lunario-Aceretum').italic=True
            p.add_run(') und Geißbart-Ahornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lindenmischwald (')
            p.add_run('Cynancho-Tilietum').italic=True
            p.add_run(') submontan bis tiefmontan auf trockeneren kalkreichen Schutthängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald als schmaler Höhengürtel, reichlich mit Lärchen gemischt. Überwiegend Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(') über skelettreichen Karbonatböden. Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Substraten, seltener Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') auf bodensauren Standorten (z.B. Tangelhumus).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Lärchenwald (')
            p.add_run('Laricetum deciduae').italic=True
            p.add_run(') kleinﬂächig in der subalpinen Stufe, an schattigen Steilhängen bis ca. 800 m hinabsteigend.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run(' Karbonat-Latschengebüsche in der hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinabreichend, häuﬁg anthropogen gefördert.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Subalpines Grünerlengebüsch (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 5.1: Niederösterreichischer Alpenostrand (Thermenalpen)
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '5.1':

            doc.add_heading('Wuchsgebiet 5.1: Niederösterreichischer Alpenostrand (Thermenalpen)', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Wuchsgebiet umfaßt zwei geomorphologisch unterschiedliche Areale, die allenfalls als Wuchsbezirke ausgeschieden werden könnten:')
            doc.add_paragraph('a) den Östlichen Flyschwienerwald mit Mergel und Sandstein und entsprechend tiefgründigen, schweren Böden; gerundete Landformen mit steilen Kerbtälern, rutsch- und hochwassergefährdet.')
            doc.add_paragraph('b) den Ostrand der Kalkalpen bzw. Kalkvoralpen im Schwarzföhrengebiet: Kalk- und Dolomitstandorte sowie tertiäre Schotter. Außer dem Schneeberg an der Wuchsgebietsgrenze Gipfel nur 700 bis 1300 m, z.T. Karsthochﬂächen und deutliche Hangverebnungen (Rumpftreppe) mit Reliktböden.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('a) Flysch-Wienerwald')
            doc.add_paragraph('Im Flysch-Wienerwald überwiegen Pseudogley (insgesamt 15% des Wuchsgebietes) und schwere Parabraunerde (knapp 10%), z.T. extrem verhagert. Zum Teil sind es alte Reliktböden.')
            doc.add_paragraph('Auf Greifensteiner Sandstein kommt auch arme, sandige Braunerde (8%) und lokal substratbedingter Podsol vor.')

            doc.add_paragraph('b) Kalkalpen')
            doc.add_paragraph('Hier herrscht Rendsina vor (insgesamt 33%), meist trockene Dolomitrendsina, Braunlehm-Rendsina (20%). Kalkbraunlehm (Terra fusca) ﬁndest man vor allem auf Verebnungen und Gipfelplateaus (21%).')
            doc.add_paragraph('Seltener ist Silikat-Braunlehm (auf Triestingschotter).')


            doc.add_heading('Natürliche Waldgesellschaften', 4)
            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Wärmeliebende Traubeneichen-Hainbuchenwälder (')
            p.add_run('Galio sylvatici-Carpinetum').italic=True
            p.add_run(',')
            p.add_run(' Carici pilosaeCarpinetum').italic=True
            p.add_run(') z.T. mit Zerreiche in der kollinen Stufe vorherrschend; submontan mit Buche, meist an wärmebegünstigten Hängen. Beimischung von Stieleiche in Talsohlen (Mulden) auf schweren vergleyten Böden.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf warmen, mäßig bodensauren Standorten Traubeneichen-Zerreichenwald (z.B. ')
            p.add_run('Quercetum petraeae-cerris').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Flaumeichenwald auf sonnigen, trockenen, kalkreichen Standorten in der kollinen Stufe. Flaumeichen-Buschwald (')
            p.add_run('Geranio sanguinei-Quercetum pubescentis').italic=True
            p.add_run(') auf flachgründigen Extremstandorten, FlaumeichenTraubeneichen-Hochwald (')
            p.add_run('Euphorbio angulatae-Quercetum pubescentis').italic=True
            p.add_run(', ')
            p.add_run('Corno-Quercetum').italic=True
            p.add_run(') auf tiefergründigen Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Beimischung von Traubeneiche, Esche, Bergahorn (Hainbuche, Kirsche, Tanne, Lärche). Fichten-Tannen-Buchenwald mit Quirl-Weißwurz (')
            p.add_run('Polygonatum verticillatum').italic=True
            p.add_run(') in der mittel- bis hochmontanen Stufe, hochmontan Fichte stärker hervortretend. Auf Karbonatgesteinen Schneerosen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Helleboro nigri-(Abieti-)Fagetum').italic=True
            p.add_run(') submontan bis hochmontan vorherrschend, mittelmontan bis hochmontan mit Grünem Alpendost (')
            p.add_run('Adenostyles glabra').italic=True
            p.add_run('). Weißseggen-Buchenwald (')
            p.add_run('Carici albae-Fagetum').italic=True
            p.add_run(') mit Schwarz- und Rotföhre submontan bis tiefmontan auf trockeneren Standorten. Bergahorn-Buchenwald (')
            p.add_run('Aceri-Fagetum').italic=True
            p.add_run(') nur sehr lokal (Schneeberg) hochmontan in schneereichen, frostgeschützten Lagen. Waldmeister-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum) submontan-montan').italic=True
            p.add_run('und Wimperseggen-Buchenwald (')
            p.add_run('Carici pilosae-Fagetum').italic=True
            p.add_run(') submontan auf leichter verwitternden, basenreichen Substraten (z.B. Flysch). Hainsimsen-(Fichten-Tannen-) Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzföhrenwälder als Dauergesellschaften an flachgründigen, sonnigen Dolomit-Steilhängen (kollin-)submontan bis mittelmontan. Auf Laubwaldstandorten sehr häuﬁg Schwarzföhren-Forste. Felsenwolfsmilch-Schwarzföhrenwald (')
            p.add_run('Euphorbio saxatilis-Pinetum nigrae').italic=True
            p.add_run(') mit Schneeheide submontan-mittelmontan zwischen Baden und Payerbach, Blaugras-Schwarzföhrenwald (')
            p.add_run('Seslerio-Pinetum nigrae').italic=True
            p.add_run(') kollin-submontan am nördlichen Alpenostrand (z.B. bei Mödling, Perchtoldsdorf). Übergänge zum Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') von Pernitz nach Westen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima (z.B. Schluchten) Laubmischwälder mit Bergahorn, Esche und Bergulme. Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(') mit Waldziest und Rasenschmiele auf wasserzügigen Unterhängen. Mondviolen-Ahornwald (')
            p.add_run('Lunario-Aceretum').italic=True
            p.add_run(') auf skelettreicheren Schluchtstandorten. Gipfel-Eschenwald (')
            p.add_run('Violo albae-Fraxinetum').italic=True
            p.add_run(') auf lockeren Pararendsinen über Kalkmergel im Gipfelbereich einiger Berge im Flysch-Wienerwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lindenmischwald (')
            p.add_run('Cynancho-Tilietum').italic=True
            p.add_run(') kollin-tiefmontan lokal (z.B. Leopoldsberg) auf trockeneren kalkreichen Schutthängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Eschen-Bestände (z.B. ')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(') als Auwald an Bächen und an quelligen, feuchten Unterhängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montaner Fichtenwald als edaphisch (Felshänge, Blockhalden) bedingte Dauergesellschaft nur lokal.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald mit Lärche und auslaufender Tanne als schmaler Höhengürtel. Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(') über skelettreichen Böden vorherrschend. HochstaudenFichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründigen, basenreichen Böden (z.B. Kalk-Braunlehm).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum')
            p.add_run(') in der hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) in die montane Stufe hinabreichend.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 5.2: Bucklige Welt
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '5.2':

            doc.add_heading('Wuchsgebiet 5.2: Bucklige Welt', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Grundgestein sind überwiegend basenarmes Silikatgestein, Gneis, Quarzphyllit; Amphibolitzüge sind untergeordnet. Abgesehen vom Hochwechselkomplex herrscht eine mittelgebirgsartige Rumpﬂandschaft mit ausgedehnten Hochﬂächen vor. Im Norden ist noch ein geringer Anteil an Kalk und pliozänen Schotterﬂuren zu ﬁnden. Oft sind tief aufgemürbte, z.T. kaolinisierte Verwitterungsdecken, erhalten.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Die Böden sind vorwiegend  basenarme Braunerden (35%), meist leicht und grusig, an Sonnhängen zur Trockenheit neigend.')
            doc.add_paragraph('Erst in relativ hohen Lagen kommt auch Semipodsol (>20%) hinzu.')
            doc.add_paragraph('Podsol tritt nur substratbedingt auf Quarzit, alten Quarzschottern und Quarzsand auf (3%).')
            doc.add_paragraph('Weiters ﬁndet man basenreiche Braunerde (>10%) auf Amphibolit und kalkbeeinﬂußtem Substrat.')
            doc.add_paragraph('Auf Altlandschaftsresten (Hochflächen, Hangstufen) sind silikatischer Relikt-Braunlehm (>10%) und Pseudogley (5%)  weit verbreitet.')
            doc.add_paragraph('Auf Semmeringtrias, Kalkschotter und ähnlichem Substrat werden circa 1% Rendsina ausgewiesen, auf welche sich die Schwarzkiefernvorkommen konzentrieren.')


            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Trotz des etwas kühleren, trockeneren Klimas gibt es in begünstigten Lagen immer noch Edelkastanie, am Nordrand Flaumeiche, Schwarzkiefer. Das Vorkommen der Tanne ist betont, sie ist z.T. vorwüchsig. Rotföhre ist stärker beigemischt als in Wuchsgebiet 5.3.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe Eichen-Hainbuchenwald und bodensaurer Rotföhren-Eichenwald (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run(') mit Besenheide (z.T. mit Edelkastanie).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lokal Schwarzföhrenwald (auf Karbonatgestein) und Flaumeichenrelikte.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Tannen-Buchenwald mit Beimischung von Eichen, Edelkastanie und Rotföhre. Föhrenanteil anthropogen erhöht. Fichten-Tannen-Buchenwald mit hohem Tannenanteil (Leitgesellschaft) in der mittelmontanen Stufe.')
            p.add_run('Vorwiegend Hainsimsen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten, auch Waldmeister-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf basenreicheren Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald in der hochmontanen Stufe vorherrschend.')
            p.add_run('Vorwiegend Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run(') auf ärmeren Silikatstandorten, auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') als schmaler Höhengürtel, nicht typisch entwickelt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpines Grünerlengebüsch kleinräumig in Kammlage (Hochwechsel).')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 5.3: Ost- und Mittelsteirisches Bergland
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '5.3':

            doc.add_heading('Wuchsgebiet 5.3: Ost- und Mittelsteirisches Bergland', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Wuchsgebiet besteht aus einer Mittelgebirgslandschaft mit ausgedehnten, ﬂachkuppigen Hochﬂächen und Hangverebnungen (Altlandschaftsresten) zum Gebirgsrand hin, in welche steilhängige V-Gräben eingeschnitten sind. Die Kammlinien liegen zwischen 1700 m und 1100 m, im Günser Gebirge bei 900 m. Sie sind meist in ebenen Gipfelﬂuren als Reste der alten Rumpftreppe gestaffelt. Nur die Kalkstöcke des Grazer Paläozoikums bilden markantere Gipfel und steile Wandabbrüche.')
            doc.add_paragraph('Das Grundgestein ist vielfältig: Ortho- und Paragneise mit Amphibolitzügen, saure Schiefer bis Kalkphyllit, paläozoischer Kalk und Quarzit (SemmeringTrias und Grazer Paläozoikum).')
            doc.add_paragraph('Auf den Altlandschaftsﬂächen des Kristallins kommen verbreitet tiefgründige, z.T. kaolinisierte Aufmürbungszonen, Verwitterungsdecken und alte, ausgewitterte Schotterdecken vor, auf Kalk-BraunlehmDecken. Im weichen Tonschiefer und Phyllit  dominieren steile V-Täler mit jungen, tiefgründigen, kolluvialen Böden.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Entsprechend vielgestaltig sind die Böden:')
            doc.add_paragraph('Basenarme, podsolige Braunerde (18%) ist vor allem im Burgenland steinig-grusig und neigt dort zur Trockenheit.')
            doc.add_paragraph('Basenreiche Braunerden und Kalkbraunerden (16%) gibt es vor allem im Grazer Bergland, nur selten östlich der Feistritz.')
            doc.add_paragraph('Semipodsol (36%*) ist vor allem auf Gneis verbreitet.')
            doc.add_paragraph('Podsol (3%) ist auf Quarzit (Semmeringtrias) beschränkt; die klimatische Podsolstufe wird nicht erreicht.')
            doc.add_paragraph('Im Grazer Paläozoikum gibt es ferner Pararendsina, Rendsina und Braunlehm-Rendsina (zusammen 11%) sowie Kalkbraunlehm (7%).')
            doc.add_paragraph('Auf den alten Abtragungsﬂächen im Kristallin treten ausgewitterte, saure Lockersedimentbraunerde, reliktischer Braun- und Rotlehm (10%) und Pseudogley (4%) auf, stellenweise ﬁndet man Podsol auf Quarzschotter.')


            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Gegenüber den nördlichen Wuchsgebieten 5.1 und 5.2 ist Tanne vitaler; Rotföhre tritt zurück.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An wärmebegünstigten Hängen in der submontanen Stufe Eichen-Hainbuchenwald (z.B. ')
            p.add_run('Asperulo odoratae-Carpinetum').italic=True
            p.add_run(') mit Buche über basenreichen Substraten und bodensaurer Eichenwald mit Rotföhre (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run(') auf ärmeren Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lokal (bei Graz) Flaumeichen-Buschwald (')
            p.add_run('Geranio sanguinei-Quercetum pubescentis').italic=True
            p.add_run(') auf Kalk.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hopfenbuchenwald (z.B. ')
            p.add_run('Ostryo-Fagetum').italic=True
            p.add_run('), z.T. mit Rotföhre, Fichte und Buche submontan bis tiefmontan an steilen, wärmebegünstigten Hängen auf Kalk (Weizklamm).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An der Mur Auwaldreste mit Silberweide (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') und Grauerle (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Tanne, Rotföhre (Edelkastanie, Eichen). In der (tief-)mittelmontanen Stufe FichtenTannen-Buchenwald (Leitgesellschaft) mit QuirlWeißwurz (')
            p.add_run('Polygonatum verticillatum').italic=True
            p.add_run('), seltener auf Karbonatstandorten auch in die hochmontane Stufe reichend.')
            p.add_run('Hainsimsen-(Fichten-Tannen-)Buchenwald (').italic=True
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren und Waldmeister-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf basenreichen silikatischen Substraten. Auf Karbonatgestein Mittelsteirischer Kalk-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Poo stiriacae-(Abieti-)Fagetum').italic=True
            p.add_run(') vorherrschend. Trockenwarmer Kalk-Buchenwald (')
            p.add_run('Carici albae-Fagetum').italic=True
            p.add_run(' s.lat.) mit Weißem Waldvöglein (')
            p.add_run('Cephalanthera damasonium').italic=True
            p.add_run(') submontan bis tiefmontan auf trockeneren Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder lokal als Dauergesellschaften an flachgründigen Sonderstandorten submontan bis mittelmontan.')
            p.add_run('Karbonat-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(' s.lat. ) mit Blaugras (Sesleria) im Grazer Bergland. Silikat-Rotföhrenwald (')
            p.add_run('Vaccinio vitis-idaeae-Pinetum')
            p.add_run(') auf Quarzit und auch auf Schatthängen über Serpentinit. Serpentin-Rotföhrenwald (')
            p.add_run('Festuco eggleri-Pinetum').italic=True
            p.add_run(' im Murtal, ')
            p.add_run('Festuco guestfalicae-Pinetum').italic=True
            p.add_run(' bei Bernstein) auf sonnigen Serpentinit-Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima in der submontanen bis mittelmontanen Stufe Laubmischwälder mit Bergahorn, Esche und Bergulme.')
            p.add_run('Z.B. Geißbart-Ahornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run(') und Hirschzungen-Ahornwald (')
            p.add_run('Scolopendrio-Fraxinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lindenmischwald (')
            p.add_run('Cynancho-Tilietum').italic=True
            p.add_run(') mit Sommerlinde auf trockeneren kalkreichen Felshängen im Hochlantschgebiet.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald mit Lärche, Bergahorn und Buche in der hochmontanen Stufe, selten tief-mittelmontan als edaphisch bedingte Dauergesellschaft (häuﬁger allerdings anthropogen entstanden).')
            p.add_run('Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Oxali-Abietetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald.')
            p.add_run('Überwiegend Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Homogyno-Piceetum').italic=True
            p.add_run(') auf Silikat. Auf Kalk (Hochlantsch) auch Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(') und Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpine Stufe nur schlecht ausgebildet (z.B. Gleinalpe, Stuhleck, Hochlantsch). Latschen- und Grünerlengebüsche (auch in tieferen Lagen), meist ersetzt durch subalpine Zwergstrauchheiden.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 5.4: Weststeirisches Bergland
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '5.4':

            doc.add_heading('Wuchsgebiet 5.4: Weststeirisches Bergland', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Gebiet ist geomorphologisch und bodenkundlich dem Wuchsgebiet 5.3 ähnlich. Lediglich im Süden hat es zusätzlich hochalpinen Charakter. Alte Landoberﬂächen (Ebenen) bzw. Reste alter Verwitterungsdecken sind dort entsprechend weniger verbreitet.')
            doc.add_paragraph('Die Grundgesteine sind vor allem Gneise, Glimmerschiefer und Amphibolit. Kalk und Quarzit sind wenig verbreitet.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Die Bodenverhältnisse entsprechen jenen im Wuchsgebiet 5.3, jedoch ohne die Karbonatböden des Grazer Paläozoikums und mit geringerer Verbreitung von Reliktböden.')
            doc.add_paragraph('Der häufigster Bodentyp ist Semipodsol (über 40%); seine untere Verbreitungsgrenze liegt auf saurem Substrat (v.a. Koralpe) schon bei ca. 600 m, sonst eher hoch.')
            doc.add_paragraph('Magere Braunerde ﬁndet sich auf nährstoffarmem Kristallin (18%), nährstoffreiche Braunerde auf Amphibolit und anderen basenreichen Kristallingesteinen (ca 18%).')
            doc.add_paragraph('Die klimatische Podsolstufe wird im Gebiet kaum mehr erreicht. Auch substratbedingter Podsol auf Quarzgängen, Quarzit, Quarzschotter (in allen Höhenstufen) ist seltener als im Wuchsgebiet 5.3.')
            doc.add_paragraph('Am Gebirgsrand gibt es in Hangverebnungen Reste alter Verwitterungsdecken. Bindiges Reliktbodenmaterial (Braunlehm) ist auch in relativ steilen Hanglagen tiefgründig erhalten (in der Soboth bis 10 m mächtig!), meist jedoch nur als Gemengeanteil in Hangdeckschichten und Ausgangsmaterial für arme, bindige Braunerde oder Pseudogley (insgesamt mit über 10% ausgewiesen).')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Die Tanne ist in diesem Wuchsgebiet begünstigt, z.T. vorwüchsig; im südlichsten Teil gibt es spitzkronige Formen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An wärmebegünstigten Hängen in der submontanen Stufe Eichen-Hainbuchenwald (')
            p.add_run('Asperulo odoratae-Carpinetum').italic=True
            p.add_run(') mit Buche über basenreicheren Substraten und bodensaurer Eichenwald mit Rotföhre (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run(') auf ärmeren Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Tanne, Rotföhre (Edelkastanie, Eichen). In der mittelmontanen Stufe Fichten-Tannen-Buchenwald (Leitgesellschaft), seltener auf Karbonatstandorten auch hochmontan.')
            p.add_run('Hainsimsen-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren und Waldmeister-(Fichten-Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf basenreichen silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima in der submontanen bis mittelmontanen Stufe Laubmischwälder mit Bergahorn, Esche, Bergulme und Sommerlinde, z.B. GeißbartAhornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Eschen-Bestände (')
            p.add_run('Stellario bulbosae-Fraxinetum').italic=True
            p.add_run(') als Auwald an Bächen und an quelligen, feuchten Unterhängen in der submontanen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannenwald mit Buche, Lärche und Bergahorn in der hochmontanen Stufe, seltener tiefmittelmontan (meist anthropogen entstanden).')
            p.add_run('Auf ärmeren Silikatstandorten Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run('), auf tiefergründigen, basenreichen Böden Sauerklee-Fichten-Tannenwald (')
            p.add_run('Galio rotundifolii-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald mit wenig Lärche.')
            p.add_run('Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(' = ')
            p.add_run('Homogyno-Piceetum').italic=True
            p.add_run(') mit Woll-Reitgras (')
            p.add_run('Calamagrostis villosa').italic=True
            p.add_run(') auf Silikat.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochsubalpine Latschen- und Grünerlengebüsche (auch in tieferen Lagen vorkommend), meist ersetzt durch subalpine Zwergstrauchheiden.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 6.1: Südliches Randgebirge
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '6.1':

            doc.add_heading('Wuchsgebiet 6.1: Südliches Randgebirge', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Die Karawanken sind vornehmlich aus Kalk und Dolomit aufgebaut, daneben kommen auch sehr saures Kristallin und Quarzit vor.')
            doc.add_paragraph('Die Karnischen Alpen bestehen - von einzelnen Hochgebirgsstöcken abgesehen - vornehmlich aus (paläozoischem) mesotrophem Silikatgestein.')
            doc.add_paragraph('Die Gailtaler Alpen umfassen v. a. im Norden (Lienzer Dolomiten bis Villacher Alpe) Karbonatgestein, am S-Abhang karbonathaltiges Silikatgestein, Schiefergneise und Glimmerschiefer.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Die häuﬁgsten Bodenformen sind Rendsina und Braunlehm-Rendsina (40%) sowie Kalkbraunlehm (20%).')
            doc.add_paragraph('Auf nährstoffreichem Silikat gibt es nährstoffreiche, zum Teil schwach kalkbeeinflußte Braunerde, auch durch Überrollung von höher gelegenen Kalkzügen (11%).')
            doc.add_paragraph('Auf nährstoffärmerem Silikatgestein kommt Semipodsol (18%) vor.')
            doc.add_paragraph('Podsol  kommt  in Hochlagen,  untergeordnet  auf Quarzitzügen oder Quarzschotter substratbedingt auch in tieferen Lagen vor (zusammen ca. 2%).')
            doc.add_paragraph('Weitere Böden des Wuchsgebiets sind Hanggley, Pseudogley, und meist bindige LockersedimentBraunerde auf Moränen und Talterrassen.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Das Wuchsgebiet ist charakterisiert durch optimales Wachstum fast aller Hauptbaumarten (Fichte, Tanne, Buche, Lärche) sowie der Nebenbaumarten Ahorn, Esche, Bergulme. An wärmebegünstigten Standorten kommen die typisch illyrischen Baumarten Schwarzföhre, Hopfenbuche und Blumenesche vor.')
            doc.add_paragraph('Die Tanne hat in diesem Wuchsgebiet ein Optimum, insbesondere auf Silikatböden und in Schattlagen. Wegen ihrer hohen Vitalität haben sich sogar viele Reinbestände erhalten. Sie geht gutwüchsig bis 1500 m, Einzelvorkommen reichen bis 1850 m!')
            doc.add_paragraph('Buche ist durchgehend als Hauptbaumart beteiligt, mit Schwerpunkt auf Kalk. Die Kalk-Buchenwälder sind durchwegs als eigene, stark illyrisch geprägte Gesellschaften ausgebildet, die von den Buchenwäldern des übrigen Österreich deutlich abgesetzt sind.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe illyrischer BuchenMischwald (')
            p.add_run('Hacquetio-Fagetum').italic=True
            p.add_run('s.lat.) auf Karbonaten. Hopfenbuchen-Buchenwald (')
            p.add_run('Ostryo-Fagetum').italic=True
            p.add_run(') submontan bis tiefmontan an wärmebegünstigten Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hopfenbuchen-Blumeneschen-Wald (')
            p.add_run('Ostryo carpinifoliae-Fraxinetum orni').italic=True
            p.add_run(') in der submontanen bis tiefmontanen Stufe an warmen, trockenen Steilhängen über Kalk und Dolomit.')
            p.add_run('Fichten-Tannen-Buchenwald in der tief- bis mittelmontanen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Überwiegend Dreiblattwindröschen-Fichten-Tannen-Buchenwald (')
            p.add_run('Anemono trifoliae-(Abieti-)Fagetum').italic=True
            p.add_run(', Leitgesellschaft) auf Karbonaten. Braunerde-Fichten-Tannen-Buchenwald (')
            p.add_run('Lamio orvalae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf tiefergründig verwitternden Kalk/Silikat-Mischsubstraten. Hainsimsen-FichtenTannen-Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Föhrenwälder als kleinﬂächige Dauergesellschaften submontan bis mittelmontan an Extremstandorten über Karbonatgestein.')
            p.add_run('Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') an sonnigen Dolomit-Steilhängen und auf Karbonatschutt weiter verbreitet. Hopfenbuchen-Schwarzföhrenwald (')
            p.add_run('Fraxino orni-Pinetum nigrae').italic=True
            p.add_run(') mit Blumenesche in Gebieten mit lokal verstärktem illyrischen Einﬂuß (z.B. Loibltal, Dobratsch) an ﬂachgründigen, sonnigen Kalk- und Dolomit-Steilhängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (Alnetum incanae) als Auwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima (unterer Bereich von Grabeneinhängen, Schluchten)  der  submontanen bis mittel(-hoch)montanen Stufe Laubmischwälder mit Bergahorn, Esche, Bergulme.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochmontaner, illyrisch geprägter Buchenwald auf basenreichen Standorten.')
            p.add_run('Auf Karbonatgestein Lanzenfarn-(Tannen-)Buchenwald (')
            p.add_run('Polysticho lonchitis-Fagetum ').italic=True
            p.add_run('= ')
            p.add_run('Saxifrago rotundifoliae-Fagetum').italic=True
            p.add_run(') vorherrschend, Süßdolden-Bergahorn-Buchenwald (')
            p.add_run('Aconiti paniculati-Fagetum').italic=True
            p.add_run(') lokal in schneereichen Lagen. Braunerde(Fichten-Tannen-)Buchenwald (')
            p.add_run('Ranunculo platanifolii-(Abieti-)Fagetum').italic=True
            p.add_run(' auf tiefergründig verwitternden Kalk/SilikatMischsubstraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Hochmontaner Fichten-Tannenwald auf ärmeren Silikatstandorten, z.B. Hainsimsen-Fichten-Tannenwald (')
            p.add_run('Luzulo nemorosae-Piceetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montaner Fichtenwald als edaphisch (Felshänge, Blockhalden) bedingte Dauergesellschaft nur lokal.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald mit Lärche, in den Karawanken nur schlecht entwickelt.')
            p.add_run('Karbonat-Alpendost-Fichtenwald (')
            p.add_run('Adenostylo glabrae-Piceetum').italic=True
            p.add_run(') über skelettreichen Karbonatböden, Hochstauden-Fichtenwald (')
            p.add_run('Adenostylo alliariae-Abietetum').italic=True
            p.add_run(') auf tiefergründig verwitternden, basenreichen Substraten, Alpenlattich-Fichtenwald (')
            p.add_run('Larici-Piceetum').italic=True
            p.add_run(') auf Silikat.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Latschengebüsche mit Wimper-Alpenrose (')
            p.add_run('Rhododendron hirsutum').italic=True
            p.add_run(') in der hochsubalpinen Stufe, an ungünstigen Standorten (z.B. Schuttriesen, Lawinenzüge) weit in die montane Stufe hinabreichend. Silikat-Latschengebüsche mit Rostroter Alpenrose (')
            p.add_run('Rhododendron ferrugineum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Karbonat-Lärchenwald (')
            p.add_run('Laricetum deciduae').italic=True
            p.add_run(') kleinﬂächig in der (montanen-)subalpinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lärchen-Zirbenwald nur lokal (westliche Karnische Alpen, Petzen).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grünerlengebüsche (')
            p.add_run('Alnetum viridis').italic=True
            p.add_run(') an feuchten, schneereichen Standorten (Lawinenstriche) in der montanen bis hochsubalpinen Stufe.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 6.2: Klagenfurter Becken
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '6.2':

            doc.add_heading('Wuchsgebiet 6.2: Klagenfurter Becken', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Wuchsgebiet ist ein inneralpines, nach allen Seiten abgeschirmtes Becken mit Lockersedimentfüllung aus Moränenmaterial und fluvioglazialen Schotterfluren, z.T. Seetonen. Teilweise ist es grundwasserfern und trocken, teilweise grundwassernahe mit Mooren und Seen. Vorherrschend sind Hügel und Inselberge aus Moräne oder anstehendem Fels, im SW liegt das Sattnitzplateau aus tertiärem Konglomerat.')
            doc.add_paragraph('Der Beckenrand umaßt die Hangfüße der Gurktaler Alpen sowie der Sau- und Koralpe.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Vorherrschend sind tiefgründige, skelettreiche Braunerde und Parabraunerde auf Moränen und Schotter; insbesondere auf Grundmoräne auch bindig und vergleyt (16%); auf Schotter seicht- bis mittelgründig und leicht (Dobrova) (36%), z.T. stark kalkhaltig (Pararendsina), z.T. tiefgründig entkalkt; bes. im Westen auf sandigem Material sauer und podsoliert (3%).')
            doc.add_paragraph('Auf tertiären Sedimenten und Altlandschaftsresten beﬁnden sich  Relikte alter Verwitterungsdecken, z.T. Braunlehm und insbesondere am Ostrand Rotlehm (insgesamt 7%).')
            doc.add_paragraph('Die Hanglagen tragen auf silikatischem Fels Braunerde (12%) und Semipodsol (13%) sowie Böden aus Karbonatgestein (8%).')
            doc.add_paragraph('Ferner gibt es Auböden, Gley sowie Anmoore und Moore (3%).')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('In Beckenlagen scheiden frostempﬁndliche Baumarten wie die Tanne aus; Buche ist labil, kommt jedoch vor, insbesondere in den Einhängen zum Drautal.')
            doc.add_paragraph('Die Waldgesellschaften in der submontanen(-tiefmontanen) Stufe sind ﬂächig durch sekundäre Rotföhren- und Fichtenwälder ersetzt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe bodensaure Eichenwälder mit Rotföhre (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run(') über silikatischen und Eichen-Hainbuchenwälder (')
            p.add_run('Helleboro nigri-Carpinetum').italic=True
            p.add_run(' s.lat.) über karbonathältigen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An wärmebegünstigten, ﬂachgründigen Steilhängen über Kalk und Dolomit Hopfenbuchen-Blumeneschen-Wald (')
            p.add_run('Ostryo carpinifoliae-Fraxinetum orni').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Purpurweiden-Filzweiden-Gebüsch (')
            p.add_run('Salicetum incano-purpureae').italic=True
            p.add_run(') als Pioniergesellschaft auf Flußschotter.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Grauerlenbestände (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') und auf durchlässigen Schotterböden auch Fichten-Rotföhrenbestände als Auwald. Bei weiter fortgeschrittener Bodenentwicklung Fichten-Eschenwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Eschen-Auwald (z.B. ')
            p.add_run('Stellario bulbosae-Fraxinetum').italic=True
            p.add_run(') an Bächen und an quelligen, feuchten Unterhängen in der submontanen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Bruchwald (z.B. ')
            p.add_run('Carici elongatae-Alnetum glutinosae').italic=True
            p.add_run(') auf Standorten mit hochanstehendem, stagnierendem Grundwasser (z.B. Seeufer) gut entwickelt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannen-Buchenwald in der tief- bis mittelmontanen Stufe, v.a. Hainsimsen-Fichten-Tannen-Buchenwald (')
            p.add_run('Luzulo nemorosae-Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima Laubmischwälder mit Bergahorn, Esche und Bergulme.')
            p.add_run('Z.B. Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(') an wasserzügigen Unterhängen, Hirschzungen-BergahornSchluchtwald (')
            p.add_run('Scolopendrio-Fraxinetum').italic=True
            p.add_run(').')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 7.1: Nördliches Alpenvorland - Westteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '7.1':

            doc.add_heading('Wuchsgebiet 7.1: Nördliches Alpenvorland - Westteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Vorwiegend ﬂachwelliges Hügelland aus tertiären Sedimenten, im Südwesten Moränenlandschaft. Vor den Endmoränenwällen liegen Sander- und Schotterﬂuren. Entlang des Inn und der Traun beﬁnden sich Schotterterrassen. Nur einzelne Flyschklippen und die tertiäre, zertalte Schotterplatte des Hausruck - Kobernaußerwaldes bilden markantere Höhenzüge.')
            doc.add_paragraph('Der nördliche Teil trägt eine fast durchgehende Lößund Staublehmdecke. Im Innviertel treten unter der Lößdecke die tertiären, tonigen Sedimente (=Schlier) zutage. Im Süden tritt an ihre Stelle Moränenmaterial.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Bindige Braunerde und Parabraunerde ﬁndet man auf Löß (8%) oder auf Staublehm und Moräne (9%); auf Grundmoräne ist sie sehr dichtgelagert, selbst seichtgründige Böden neigen dort zu Wasserstau.')
            doc.add_paragraph('Einen großen Anteil nimmt Pseudogley auf Schlier, Staublehm und v.a. älterem Löß, seltener auf Moräne, sowie Grundwassergley ein (zusammen 24%).')
            doc.add_paragraph('Pararendsina (1%) und leichte Braunerden (24%) sind auf Moräne, Schotter und Sand entwickelt.')
            doc.add_paragraph('Die tertiären Schotter des Hausruck tragen saure, steinige, meist podsolige Braunerde bis Podsol. Während die fruchtbaren Böden unter Acker- und Grünlandkultur stehen, sind die podsoligen Böden dem Wald verblieben. Ihr Anteil an der Waldﬂäche beträgt daher 25%!')
            doc.add_paragraph('Ferner gibt es Auböden (5%), Anmoore, Niedermoore und Hochmoore (3% der Waldﬂäche).')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Von Natur aus sind hier nährstoffreiche, leistungsfähige Laubmischwald-Standorte verbreitet; die besseren Standorte sind allerdings unter landwirtschaftlicher Nutzung (Äcker, Grünland).')
            p = doc.add_paragraph('Ersatzgesellschaften mit Fichte (Rotföhre) nehmen den größten Anteil an der Waldﬂäche ein. Die natürliche Waldvegetation ist daher vielfach nur schwer erkennbar. Häuﬁg sind Vergrasungen mit Seegras (')
            p.add_run('Carex brizoides').italic=True
            p.add_run('), z.T. gibt es auch Degradationen mit Torfmoos (')
            p.add_run('Sphagnum').italic=True
            p.add_run('), Pfeifengras (')
            p.add_run('Molinia').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Submontaner Stieleichen-Hainbuchenwald (')
            p.add_run('Galio sylvatici-Carpinetum').italic=True
            p.add_run(') an wärmebegünstigten, trockenen Standorten oder auf schlecht durchlüfteten, bindigen, staunassen Böden; meist durch Fichtenbestände ersetzt.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe Buchenwald mit Tanne (Edellaubbaumarten, Stieleiche, Rotföhre), tiefmontan (Fichten-)Tannen-Buchenwald.')
            p.add_run('Hainsimsen-(Tannen-)Buchenwald (')
            p.add_run('Luzulo nemorosae (Abieti-)Fagetum').italic=True
            p.add_run(' auf ärmeren, bodensauren und Waldmeister-(Tannen-)Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf basenreicheren Standorten. Auf den KalkschotterTerrassen (z.B. Traun, Salzach) auch Kalk-Buchenwälder (z.B. ')
            p.add_run('Carici albae-Fagetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Peitschenmoos-Fichten-Tannenwald (')
            p.add_run('Mastigobryo-Piceetum').italic=True
            p.add_run(') mit Torfmoos auf bodensauren, staunassen Standorten wohl meist anthropogen entstanden, ursprünglich mit höherem Buchen- und Stieleichenanteil; kleinflächig vielleicht auch als edaphisch bedingte Dauergesellschaft.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auwälder der größeren Flußtäler: Silberweiden-Au (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') als Pioniergesellschaft auf schlufﬁg-sandigen Anlandungen, Purpurweiden-Filzweiden-Gebüsch (')
            p.add_run('Salicetum incano-purpureae').italic=True
            p.add_run(') auf Schotter. Grauerlen-Au (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') gut entwickelt.')

            p.add_run('Bei weiter fortgeschrittener Bodenentwicklung und nur mehr seltener Überschwemmung Hartholz-Au mit Esche, Bergahorn, Grauerle, Stieleiche, Winterlinde: In Alpennähe (z.B. Salzach) mit Bergulme (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(' = ')
            p.add_run('Aceri-Fraxinetum').italic=True
            p.add_run('), am Inn auch mit Feldulme (')
            p.add_run('Querco-Ulmetum').italic=True
            p.add_run(').')

            p.add_run('Auf durchlässigen Schotterböden (Alm-Auen) außerdem (Fichten-)Rotföhrenbestände (')
            p.add_run('Dorycnio-Pinetum').italic=True
            p.add_run(' s.lat.).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Entlang der kleineren Bäche Grauerlen-Au (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') und Eschen-Schwarzerlen-Bachauwälder (')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(',')
            p.add_run(' Pruno-Fraxinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Entlang der kleineren Bäche Grauerlen-Au (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') und Eschen-Schwarzerlen-Bachauwälder (')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(', ')
            p.add_run('Pruno-Fraxinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Bruchwald (')
            p.add_run('Carici elongatae-Alnetum glutinosae').italic=True
            p.add_run(') auf Standorten mit hochanstehendem, stagnierendem Grundwasser.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schneeheide-Rotföhrenwald (')
            p.add_run('Erico-Pinetum sylvestris').italic=True
            p.add_run(') kleinﬂächig als Dauergesellschaft an Konglomeratschutt-Steilhängen (Traunschlucht).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An nährstoffreichen, frischen, meist rutschgefährdeten Standorten (z.B. Grabeneinhänge) Laubmischwälder mit Bergahorn, Esche und Bergulme, z.B. Geißbart-Ahornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run(') und Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(').')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 7.2: Nördliches Alpenvorland - Ostteil
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '7.2':

            doc.add_heading('Wuchsgebiet 7.2: Nördliches Alpenvorland - Ostteil', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Das Gebiet besteht aus Hügelland und Terrassenﬂuren. Den Untergrund bilden tertiäre Sedimente: Ton, Sand, Tonmergel; sie sind weithin in Terrassenstufen gegliedert und mit Schotter, Löß und Staublehm bedeckt. Nur lokal treten Flyschinseln zutage oder es reicht sogar der kristalline Untergrund an die Oberﬂäche (Hiesberg).')


            doc.add_heading('Böden', 4)
            doc.add_paragraph('Auf anstehendem Tertiär liegen meist Pseudogley und vergleyte Braunerde.')
            doc.add_paragraph('Auf den Löß- und Staublehmdecken (gleichen Alters) der Terrassenlandschaft kommt das klimatische Ost-West-Gefälle auch in der Bodenbildung zum Ausdruck: Braunerde und Parabraunerde auf Löß (15%) liegen eher im Osten, Pseudogley auf (verlehmtem) Löß sind vorwiegend im Westen zu ﬁnden. Insgesamt nimmt Pseudogley 30% der Waldfläche ein, bindige Braunerde auf Löß und anderen Lockersedimenten 20%.')
            doc.add_paragraph('Auf jungem Terrassenschotter tritt Pararendsina und seichtgründige bzw. skelettreiche, leichte Braunerde (14%) in den Vordergrund, besonders an den Terrassenrändern.')
            doc.add_paragraph('Auf anstehendem Silikatgestein (z.B. Hiesberg) sind Fels-Braunerden unterschiedlicher Trophie relativ weit verbreitet (18%).')
            doc.add_paragraph('Bedeutung haben die außerordentlich fruchtbaren Fluß- und Stromauböden (Donauauen! 16%). Die Gleyböden der Talsohlen tragen relativ wenig Wald.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            p=doc.add_paragraph('Von Natur aus überwiegen nährstoffreiche, leistungsfähige Laubmischwald-Standorte. Verbreitet sind Ersatzgesellschaften mit Fichte (Rotföhre), häuﬁg Vergrasungen mit Seegras (')
            p.add_run('Carex brizoides').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der kollinen Stufe Stieleichen-Hainbuchenwald (')
            p.add_run('Galio sylvatici-Carpinetum').italic=True
            p.add_run(') vorherrschend; submontan mit Buche, meist an wärmebegünstigten Standorten. Natürlicher Rotföhrenanteil v.a. an den Kanten der Schotterterrassen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe Buchenwald mit Tanne (Edellaubbaumarten, Stieleiche).')
            p.add_run('Meist Hainsimsen-Buchenwald (')
            p.add_run('Luzulo nemorosae-Fagetum').italic=True
            p.add_run(') auf ärmeren, bodensauren Standorten. Auf den KalkschotterTerrassen (z.B. Traun, Enns) auch Kalk-Buchenwälder (z.B. ')
            p.add_run('Carici albae-Fagetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auwälder der größeren Flußtäler und der Donau: Silberweiden-Au (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') als Pioniergesellschaft auf schlufﬁg-sandigen Anlandungen, Purpurweiden-Filzweiden-Gebüsch (')
            p.add_run('Salicetum incano-purpureae').italic=True
            p.add_run(', ')
            p.add_run('Salix purpurea-Ges.').italic=True
            p.add_run(') auf Schotter. Grauerlen-Au (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') an den Flüssen gut entwickelt. An der Donau Silberpappel-Au (')
            p.add_run('Fraxino-Populetum').italic=True
            p.add_run('), Grauerlen-Au dort hauptsächlich an Uferwällen oder durch Niederwaldwirtschaft (Ersatzgesellschaft) entstanden.')

            p.add_run('Bei weiter fortgeschrittener Bodenentwicklung und nur mehr seltener Überschwemmung Hartholz-Au mit Esche, Bergahorn, Grauerle, Stieleiche, Winterlinde: An den Flüssen mit Bergulme (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(' = ')
            p.add_run('Aceri-Fraxinetum').italic=True
            p.add_run('), an der Donau auch mit Feldulme, Flatterulme (')
            p.add_run('Querco-Ulmetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Entlang der kleineren Bäche Grauerlen-Au (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(') und Eschen-Schwarzerlen-Auwälder (')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(', ')
            p.add_run('Pruno-Fraxinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An nährstoffreichen, frisch-feuchten Standorten (z.B. Grabeneinhänge) Laubmischwälder mit Bergahorn, Esche und Bergulme, z.B. Bergahorn-Eschenwald (')
            p.add_run('Carici pendulae-Aceretum').italic=True
            p.add_run(').')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 8.1: Pannonisches Tief- und Hügelland
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '8.1':

            doc.add_heading('Wuchsgebiet 8.1: Pannonisches Tief- und Hügelland', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Der Raum umfaßt im wesentlichen tertiäres Hügelland und Schotterterrassen. Beide Landschaftselemente sind zum Teil mit Löß oder kalkfreiem Flugstaub bedeckt.')
            doc.add_paragraph('Dagegen bilden ältere, ausgewitterte Quarzschotter (Hollabrunn, Rauchenwarter Platte), Kalkklippen (Leiser Berge, Hainburger Berge) und alpin-karpatische Kristallinsockel (Leithagebirge, Hainburger Berge) vielfältige Standortsbedingungen.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Im Westen (gegen das Waldviertel zu) und im Hügelland überwiegen Braunerde und Parabraunerde auf Löß und tertiären Sedimenten (insgesamt 31%).')
            doc.add_paragraph('Im Osten überwiegt Tschernosem, der kennzeichnende und häuﬁgste Bodentyp des Wuchsgebietes. Er nimmt als vorzüglicher Ackerboden aber nur 11% der Waldﬂäche - v.a. Ausschlagwald - ein.')
            doc.add_paragraph('Daneben kommen bindige Reliktlehme auf älteren Schotterterrassen und vor allem im Leithagebirge vor (insgesamt 6%).')
            doc.add_paragraph('Besonders im Wiener Becken sind grundwassernahe, schwere Böden - Gley und Feuchtschwarzerde verbreitet, welche allerdings nur kleinere Waldkomplexe tragen.')
            doc.add_paragraph('Auf kalkfreiem Flugstaub über Schotter (z. B. Gänserndorfer Terrasse) liegt Paratschernosem, der ökologisch der mageren, trockenen Braunerde auf Sand und Schotter nahesteht und forstlich sehr unproduktiv ist. Sanddünen sind dort häuﬁg.')
            doc.add_paragraph('Die seichtgründigen, rendsinaartigen Böden des Steinfeldes sind Grenzstandorte für Wald. Die älteren Schotter des Alpenrandes (z. B. Hernstein) tragen Kalkbraunlehm-Reste (zusammen 3%) und sind standörtlich günstiger (wärmeliebender Laubmischwald).')
            doc.add_paragraph('Rendsina und Kalkbraunlehm treten auch auf den Kalkklippen und im Leithagebirge auf.')
            doc.add_paragraph('Dort und auf anderen Kristallinsockeln, ebenso wie auf Quarzschotterﬂuren, ist magere, saure, basenarme Braunerde überraschend häuﬁg (zusammen mit Paratschernosem 14%).')
            doc.add_paragraph('Einen großen Flächenanteil nehmen die hochproduktiven Böden der March- und Donauauen (24%) ein.')
            doc.add_paragraph('Die Salzböden des Seewinkels sind Nichtholzböden.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Das Wuchsgebiet ist vorzüglich für landwirtschaftliche Kulturen geeignet und dementsprechend überwiegend landwirtschaftlich genutzt. Dennoch beträgt die Waldﬂäche weit über 100.000 ha, 60% davon sind Ausschlagwald.')
            doc.add_paragraph('Eine Sonderstellung nehmen die überaus produktiven Stromauwälder der Donau, March und Thaya mit ca. 24.000 ha ein.')
            doc.add_paragraph('Natürliche Wald-Grenzstandorte (Rendsinen im Steinfeld, Sanddünen im Marchfeld) wurden großﬂächig v.a. mit Schwarzföhre (Robinie) aufgeforstet.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Kollin-planar auf warmen, mäßig bodensauren Standorten Zerreichen-Traubeneichenwald (')
            p.add_run('Quercetum petraeae-cerris').italic=True
            p.add_run('). Auf kalkhältigen Löß-Standorten nur mehr fragmentarisch (z.B. Parndorfer Platte) Löß-Eichenwald (')
            p.add_run('Aceri tatarici-Quercetum').italic=True
            p.add_run(') mit Zerreiche, Stieleiche, Flaumeiche, Feldahorn.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Wärmeliebende Eichen-Hainbuchenwälder (')
            p.add_run('Primulo veris-Carpinetum').italic=True
            p.add_run(', ')
            p.add_run('Carici pilosae-Carpinetum').italic=True
            p.add_run(') in der kollinen und submontanen Stufe vorherrschend, an grundwasserfernen Standorten mit Traubeneiche, besonders in Talsohlen und Muldenlagen mit Stieleiche; submontan mit Buche.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Flaumeichenwald auf sonnigen, trockenen, kalkreichen Standorten in der kollinen Stufe, v.a. in Gebieten mit Hartgesteinen (Hainburger Berge, Leithagebirge, Klippenzone im Weinviertel).')

            p.add_run('Flaumeichen-Buschwald (')
            p.add_run('Pruno mahaleb-Quercetum pubescentis').italic=True
            p.add_run(', ')
            p.add_run('Geranio sanguinei-Quercetum pubescentis').italic=True
            p.add_run(') auf ﬂachgründigen Extremstandorten. Flaumeichen-TraubeneichenHochwald (')
            p.add_run('Euphorbio angulatae-Quercetum pubescentis').italic=True
            p.add_run(', ')
            p.add_run('Corno-Quercetum').italic=True
            p.add_run(') auf tiefergründigen Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe Buchenwald (')
            p.add_run('Melittio-Fagetum').italic=True
            p.add_run(') mit Traubeneiche und Hainbuche an kühleren Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auwälder der größeren Flußtäler und der Donau: Silberweiden-Au (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') als Pioniergesellschaft auf schlufﬁg-sandigen Anlandungen, Purpurweiden-Gebüsch (')
            p.add_run('Salix purpurea').italic=True
            p.add_run('-Ges.')
            p.add_run(') auf Schotter, Mandelweiden-Gebüsch (')
            p.add_run('Salicetum triandrae').italic=True
            p.add_run(') auf Schlick.')

            p.add_run('Silberpappel-Au (')
            p.add_run('Fraxino-Populetum').italic=True
            p.add_run(') an der Donau großflächig entwickelt. Grauerlen-Au kleinﬂächig an Uferwällen oder durch Niederwaldwirtschaft entstanden.')

            p.add_run('Hartholz-Au mit Eschen, Stieleiche, Feldulme und Flatterulme bei weiter fortgeschrittener Bodenentwicklung und nur mehr seltener Überschwemmung. An der Donau mit Gewöhnlicher Esche (')
            p.add_run('Querco-Ulmetum').italic=True
            p.add_run('), an March und Leitha mit Quirlesche (')
            p.add_run('Fraxino pannonicae-Ulmetum').italic=True
            p.add_run('). Die am seltensten überschwemmten Austandorte mit Winterlinde und Hainbuche.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Entlang kleinerer Bäche Eschen-SchwarzerlenBachauwälder (z.B. ')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run('). Bruchwaldartige Schwarzerlenbestände auf Niedermoor-Standorten (z.B. Marchegg, Wiener Becken, Neusiedlersee, Hanság).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Laubmischwälder mit Esche, Sommerlinde, Bergahorn, Bergulme an kühl-schattigen Standorten nur selten vorhanden, z.B. Lerchensporn-AhornEschenwald (')
            p.add_run('Corydalido cavae-Aceretum').italic=True
            p.add_run('), Lindenmischwald (')
            p.add_run('Cynancho-Tilietum').italic=True
            p.add_run(').')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 8.2: Subillyrisches Hügel- und Terrassenland
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '8.2':

            doc.add_heading('Wuchsgebiet 8.2: Subillyrisches Hügel- und Terrassenland', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Vom Alpenrand nach Südosten auslaufende Riedel (250 - 500 m) prägen das Gebiet. Im Süden liegt die Murebene.')
            doc.add_paragraph('Den Untergrund bilden tertiäre Sedimente aus Schotter, Sand, Ton, Tonmergel. Diese sind in Terrassen und Täler zergliedert. Dabei ist zum Teil das tertiäre Substrat freigelegt, zum Teil ist es mit jüngeren Terrassenschottern, Staublehm und Reliktböden bedeckt. Kleinräumig treten Inseln aus Quarzphyllit (Sausal) und vulkanischem Gestein (Gleichenberg) zutage.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Die Böden sind im Gegensatz zu Wuchsgebiet 8.1 karbonatfrei und im allgemeinen sauer.')
            doc.add_paragraph('Besonders am Gebirgsrand sind großﬂächig Reste alter Verwitterungsdecken - meist tiefergründig silikatischer Braunlehm, seltener Rotlehm - erhalten (8%).')
            doc.add_paragraph('Daneben gibt es auf Quarzschotter auch podsolige Braunerde bis Podsol (1%).')
            doc.add_paragraph('Im tieferen Hügelland selbst überwiegt extremer Pseudogley aus Staublehm (“Opok”), in den Talsohlen sind schwere Gleyböden verbreitet (zusammen 53%!).')
            doc.add_paragraph('Dazu kommen schwere Braunerde, vor allem auf Hangrücken (20%), und leichte Braunerden auf Schotter oder tertiärem Sand (9%).')
            doc.add_paragraph('Ferner kommen vor: Anmoore, Niedermoore, Auböden (3%) sowie magere Felsbraunerden auf Quarzphyllit und sauren vulkanischen Gesteinen.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Anthropogene Rotföhrenwälder und Fichtenforste sind im Gebiet weit verbreitet.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf wärmebegünstigten, mäßig bodensauren Standorten Traubeneichenwald mit Zerreiche (')
            p.add_run('Quercetum petraeae-cerris').italic=True
            p.add_run(') randlich in der kollinen Stufe im Osten des Gebiets.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der kollinen und submontanen Stufe EichenHainbuchenwälder (z.B. ')
            p.add_run('Asperulo odoratae-Carpinetum').italic=True
            p.add_run(' mit Waldmeister, ')
            p.add_run('Fraxino pannonicae-Carpinetum').italic=True
            p.add_run('mit Stieleiche und Seegras-Segge) auf tiefergründigen, basenreicheren Standorten, submontan mit Buche.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf stark bodensauren Standorten Rotföhren-Eichenwälder.')
            p.add_run('Drahtschmielen-Eichenwald (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run(') auf trockeneren Standorten, Pfeifengras-Stieleichenwald (')
            p.add_run('Molinio arundinaceae-Quercetum').italic=True
            p.add_run(') mit Schwarzerle auf vernäßten Standorten (z.B. Mur-Terrassen).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe Buchenwald mit Eichen, Tanne, Edelkastanie, Rotföhre vorherrschend. Auf bindigen Böden höherer Tannen-Anteil bis in tiefe Lagen.')
            p.add_run('Überwiegend Hainsimsen-(Tannen-)Buchenwald (')
            p.add_run('Luzulo nemorosae-Fagetum').italic=True
            p.add_run(') auf ärmeren silikatischen Substraten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auwälder der größeren Flußtäler: Silberweiden-Au (')
            p.add_run('Salicetum albae').italic=True
            p.add_run(') als Pioniergesellschaft auf schlufﬁg-sandigen Anlandungen. Silberpappel-, Grauerlen- und Schwarzerlen-Auwälder. Hartholz-Au mit Flatterulme, Stieleiche und Esche bei weiter fortgeschrittener Bodenentwicklung und nur mehr seltener Überschwemmung.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Eschen-Bestände (')
            p.add_run('Stellario bulbosae-Fraxinetum').italic=True
            p.add_run('), ')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(') als Auwald an Bächen und an quelligen, feuchten Unterhängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Bruchwald (')
            p.add_run('Carici elongatae-Alnetum glutinosae').italic=True
            p.add_run(') auf Standorten mit hochanstehendem, stagnierendem Grundwasser.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('An nährstoffreichen, frisch-feuchten Standorten (z.B. Grabeneinhänge) Laubmischwälder mit Bergahorn, Esche und Bergulme.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 9.1: Mühlviertel
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '9.1':

            doc.add_heading('Wuchsgebiet 9.1: Mühlviertel', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Die Landschaft wird durch kristallines Rumpfgebirge mit ﬂachen, nur zum Böhmerwald-Hauptkamm hin ausgeprägteren Mittelgebirgsformen und Steilhängen zum Donautal (Schutzwald!) geprägt. Mit den Abtragungsformen verknüpft sind alte Verwitterungsdecken: tiefgründig aufgemürbtes, kaolinisiertes Grundgestein und Braunlehmdecken, im Sauwald auch Rotlehmreste. Weiters sind Blockﬂuren und Soliﬂuktionsdecken verbreitet.')
            doc.add_paragraph('Im Freistädter Becken und am Südrand des Wuchsgebietes liegen tertiäre Tone und Sande sowie Lößund Flugsanddecken.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Böden der Braunerde-Podsolreihe aus Kristallin herrschen vor. In tiefen Lagen (z. B. Donautal) sowie auf nährstoffreicherem Granit und Gneis überwiegt Braunerde.')
            doc.add_paragraph('Es ﬁndet sich magere Braunerde auf saurem Granit und Gneis (27%). Reichere und meist auch bindigere Braunerde gibt es auf Hornblendegneis (Julbach) u.ä. nährstoffreichem Silikatgestein, leichte, aber basenreiche Braunerde auch auf Weinsberger Granit (zusammen 11%).')
            doc.add_paragraph('In mittleren Lagen herrscht Semipodsol (45%) vor.')
            doc.add_paragraph('In den höchsten Lagen des Böhmerwaldes und des Grenzkammes zwischen Mühl- und Waldviertel gibt es klimatisch bedingt Podsol, auf Eisgarner Granit und Quarzsand - vor allem nördlich von Linz und um Freistadt - substratbedingt auch in tieferen Lagen (zusammen 6%).')
            doc.add_paragraph('Vor allem im Sauwald und im westlichsten Mühlviertel kommen bindige Reliktböden hinzu, tiefgründiger Braunlehm (3%) und Pseudogley (3%).')
            doc.add_paragraph('Im Freistädter Becken und im Linzer Raum ﬁndet man Braunerde und Parabraunerde auf Löß und lößähnlichen Sedimenten (2%).')
            doc.add_paragraph('Anmoore und Hochmoore machen immerhin etwa 3% der Waldﬂäche aus.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)
            doc.add_paragraph('Das Wuchsgebiet 9.1 ist subherzynisches Fichten-Tannen-Buchen-Mischwaldgebiet. Auf reicher Braunerde (Hornblendegneis) reicht Buche bis in Hochlagen. Reichere, bindige Braunerden (Perlgneis) begünstigen die Tanne. In den tieferen Randlagen sind auch reiche (Eichen-)Buchen-Mischwaldgesellschaften entwickelt.')
            doc.add_paragraph('Verbreitet sind anthropogene Fichten-Ersatzgesellschaften und besonders in den tieferen Lagen sekundäre Rotföhrenwälder.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen Stufe Stieleichen-Hainbuchenwald (')
            p.add_run('Galio sylvatici-Carpinetum').italic=True
            p.add_run(') z.T. mit Traubeneiche, Buche an wärmebegünstigten Hängen auf reicheren Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Bodensaure, nährstoffarme submontaneRotföhrenEichenwälder.')

            p.add_run('Geißklee-Traubeneichenwald (')
            p.add_run('Cytiso nigricantis-Quercetum').italic=True
            p.add_run(') auf wärmebegünstigten Silikatstandorten (Donautal zwischen Passau und Linz) und Drahtschmielen-Stieleichenwald (')
            p.add_run('Deschampsio ﬂexuosae-Quercetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lindenmischwälder an Sonderstandorten in der submontanen Stufe.')
            p.add_run('Schlucht-Lindenwald (')
            p.add_run('Aceri-Carpinetum').italic=True
            p.add_run(') mit Spitzahorn, Hainbuche an meist schattigen Hangschuttstandorten; SilikatBlock-Lindenwald (')
            p.add_run('Poo nemoralis-Tilietum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Tanne (Fichte, Eichen) vorherrschend.')

            p.add_run('Vorwiegend Hainsimsen-Buchenwald (')
            p.add_run('Luzulo nemorosae-Fagetum').italic=True
            p.add_run(') mit Rotföhre auf ärmeren, Waldmeister-Buchenwald (')
            p.add_run('Asperulo odoratae-Fagetum').italic=True
            p.add_run(') auf basen- und nährstoffreicheren Silikatstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Bodensaurer Rotföhrenwald (')
            p.add_run('Dicrano-Pinetum').italic=True
            p.add_run(') als kleinﬂächige Dauergesellschaft submontan bis tief(-mittel)montan an flachgründigen Felskuppen; anthropogen entstanden (z.B. Streunutzung) oft auch an besseren Standorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen bis tiefmontanen Stufe EschenSchwarzerlen-Auwälder an Bächen und Flüssen. Waldsternmieren-Schwarzerlenwald (')
            p.add_run('Stellario nemori-Alnetum glutinosae').italic=True
            p.add_run(') mit Bruchweide und Geißfuß (')
            p.add_run('Aegopodium').italic=True
            p.add_run(') auf Schwemmböden, Winkelseggen-Eschen-Schwarzerlenwald (')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(') an quelligen Stellen (Gleyböden).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auwaldreste mit Grauerle (')
            p.add_run('Alnetum incanae').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Laubmischwälder mit Esche, Bergahorn, Spitzahorn, Bergulme und Buche an frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima (Grabeneinhänge, Schluchten).')
            p.add_run('Z.B. Bingelkraut-Ahorn-Eschenwald (')
            p.add_run('Mercuriali-Fraxinetum').italic=True
            p.add_run(') und Geißbart-Ahornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run(') submontanmittelmontan; Hochstauden-Ulmen-Bergahornwald (')
            p.add_run('Ulmo-Aceretum').italic=True
            p.add_run(') hochmontan (Böhmerwald).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannen-Buchenwald (Leitgesellschaft) in der mittel-hochmontanen Stufe.')
            p.add_run('Vorwiegend Wollreitgras-Fichten-Tannen-Buchenwald (')
            p.add_run('Calamagrostio villosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren Standorten. Auf basen- und nährstoffreicheren Silikatstandorten (Weinsberger Granit) Quirlzahnwurz-Fichten-Tannen-Buchenwald (')
            p.add_run('Dentario enneaphylli-(Abieti-)Fagetum').italic=True
            p.add_run(' mit Kleeschaumkraut (')
            p.add_run('Cardamine trifolia').italic=True
            p.add_run(') und Quirl-Weißwurz (')
            p.add_run('Polygonatum verticillatum').italic=True
            p.add_run('). Degradation der Bodenvegetation zum AstmoosHeidelbeer-Drahtschmiele-Typ ist jedoch auch dort möglich.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montane Fichten- und Fichten-Tannenwälder als edaphisch oder lokalklimatisch bedingte Dauergesellschaften.')
            p.add_run('Waldschachtelhalm-Tannen-Fichtenwald (')
            p.add_run('Equiseto sylvatici-Abietetum').italic=True
            p.add_run(') auf vernäßten Flachhängen (Gleystandorte, “Fichten-Au”), Peitschenmoos-Tannen-Fichtenwald (')
            p.add_run('Mastigobryo-Piceetum').italic=True
            p.add_run(') mit Torfmoos auf anmoorigen Standorten, oft in Inversionslagen; Blockﬂur-Fichtenwald, Moorrand-Fichtenwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auf Torfböden (Hochmoore) Fichten-Rotföhrenwald (')
            p.add_run('Vaccinio uliginosi-Pinetum sylvestris').italic=True
            p.add_run(') sowie Latschen-, Spirken- und Moorbirken-Bestände.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Tiefsubalpiner Fichtenwald (')
            p.add_run('Soldanello montanae-Piceetum').italic=True
            p.add_run(') mit Woll-Reitgras (')
            p.add_run('Calamagrostis villosa').italic=True
            p.add_run(') nur lokal entwickelt (z.B. Böhmerwald).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lokal im Gipfelbereich des Plöckensteins Latschengebüsch auf Blockschutt.')


#-------------------------------------------------------------------------------------------------------------------------------
#   Wuchsgebiet 9.2: Waldviertel
#-------------------------------------------------------------------------------------------------------------------------------

        elif wuchs == '9.2':

            doc.add_heading('Wuchsgebiet 9.2: Waldviertel', 3)

            doc.add_heading('Geomorphologie', 4)
            doc.add_paragraph('Kristallines Rumpfgebirge mit ﬂachen Mittelgebirgsformen und Hochﬂächen in von Westen nach Nordosten zu abnehmender Höhenlage. Tief eingeschnittene Gräben und Schluchten sowie Steilhänge zum Donautal (Schutzwald!) kennzeichnen das Gebiet. Mit den Abtragungsformen verknüpft sind alte Verwitterungsdecken: tiefgründig aufgemürbtes, kaolinisiertes Grundgstein und Braunlehmdecken, Blockﬂuren und Soliﬂuktionsdecken.')
            doc.add_paragraph('Im Norden und Osten treten tertiäre Sedimente, Tone und im Raum Gmünd-Litschau Quarzsand auf. Im Norden und vom Ostrand her greifen Löß- und Flugsanddecken von der Niederung in die tieferen Lagen des Wuchsgebietes über. Dadurch ist auch landschaftlich der Ostrand mit der angrenzenden pannonischen Niederung verzahnt.')

            doc.add_heading('Böden', 4)
            doc.add_paragraph('Es herrschen meist leichte, sandig grusige Böden der Braunerde-Podsolreihe aus Kristallin vor.')
            doc.add_paragraph('In tiefen Lagen (z. B. Donautal) sowie auf nährstoffreicherem Granit (Weinsberger Wald) und Gneis überwiegt Braunerde: basenarme, magere Braunerde auf saurem Granit und Gneis (40%); reichere Braunerde auf Hornblendegneis u.ä nährstoffreichem Silikatgestein (auch Weinsberger Granit, 10%).')
            doc.add_paragraph('In mittleren Lagen tritt Semipodsol (20%) auf.')
            doc.add_paragraph('Der Podsol in den höchsten Lagen des Weinsberger Waldes und des Ostrong ist eher klimatisch bedingt auf Eisgarner Granit und Quarzsand kommt er substratbedingt auch in tieferen Lagen (zusammen 10%) vor.')
            doc.add_paragraph('Auf den Abtragungsﬂächen sind bindige Reliktböden verbreitet, tiefgründige bindige Braunerde und Braunlehm (3%) und vor allem im Raum Zwettl-Allentsteig auf tertiären Sedimenten - Pseudogley (9%).')
            doc.add_paragraph('Auf diesen Ebenen sind auch Anmoore und Hochmoore konzentriert (3%).')
            doc.add_paragraph('Verbreitet ist Humus-Verhagerung (30% der Waldfläche); Rohhumus ist hingegen seltener (1% der Waldﬂäche mit Auﬂagen > 9 cm).')
            doc.add_paragraph('Im Nordosten des Wuchsgebiets ist in verzahnten Vorkommen Braunerde und Parabraunerde auf Löß relativ weit verbreitet (5%).')
            doc.add_paragraph('Ferner gibt es magere Braunerde auf Flugsand und Schotter, Bachau- und Schwemmböden.')

            doc.add_heading('Natürliche Waldgesellschaften', 4)

            p=doc.add_paragraph('Subherzynisches Fichten-Tannen-Buchen-Mischwaldgebiet mit vergleichsweise hohem Fichtenanteil und kühl-borealen Florenelementen, z.B. Siebenstern (')
            p.add_run('Trientalis europaea').italic=True
            p.add_run(') und Woll-Reitgras (')
            p.add_run('Calamagrostis villosa').italic=True
            p.add_run('). Der Effekt der Klimadepression auf die Vegetation wird durch das saure Substrat (Granit, Gneis) verstärkt.')

            p.add_run('In den tieferen Randlagen gibt es auch reiche Eichen-Buchen-Mischwaldgesellschaften. Fichtenforste sind hier besonders gefährdet.')

            p.add_run(' Verbreitet sind sekundäre Rotföhrenwälder und anthropogene Fichten-Ersatzgesellschaften.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lokal (Wachau) Flaumeichen-Buschwald auf trockenwarmen Karbonatstandorten der kollinen Stufe.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Auwälder an der Donau (Wachau): Silberpappel-Au (')
            p.add_run('Fraxino-Populetum').italic=True
            p.add_run('), Schwarzpappel-Au und Hartholz-Au (')
            p.add_run('Querco-Ulmetum').italic=True
            p.add_run(') mit Esche, Feldulme und Stieleiche.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Traubeneichen-Hainbuchenwälder (v.a. ')
            p.add_run('Melampyro nemorosi-Carpinetum').italic=True
            p.add_run(') auf reicheren Standorten in der kollinen Stufe, submontan an wärmebegünstigten Hängen.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Stark bodensaure, nährstoffarme Rotföhren-Eichenwälder kollin bis submontan.')

            p.add_run('Haarginster-Traubeneichenwald (')
            p.add_run('Genisto pilosae-Quercetum petraeae').italic=True
            p.add_run(') und Elsbeeren-Traubeneichenwald (')
            p.add_run('Sorbo torminalis-Quercetum').italic=True
            p.add_run(') auf trockenen, sonnigen Silikatstandorten; Drahtschmielen-Stieleichenwald (')
            p.add_run('Deschampsio flexuosae-Quercetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Lindenmischwälder in der (kollinen-)submontanen Stufe an Sonderstandorten.')

            p.add_run('Schlucht-Lindenwald (')
            p.add_run('Aceri-Carpinetum').italic=True
            p.add_run(') mit Spitzahorn, Hainbuche an meist schattigen Hangschuttstandorten, SilikatBlock-Lindenwald (')
            p.add_run('Poo nemoralis-Tilietum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('In der submontanen und tiefmontanen Stufe Buchenwald mit Tanne, Fichte (Eichen) als Leitgesellschaft.')

            p.add_run('Vorherrschend Hainsimsen-(Fichten-Tannen-) Buchenwald (')
            p.add_run('Luzulo nemorosae-(Abieti-)Fagetum').italic=True
            p.add_run(') mit Rotföhre auf ärmeren Silikatstandorten; auf substratbedingtem Podsol sehr labil, weitgehend degradiert zu Besenheide-(Calluna-)Föhrenwald. Besonders auf Podsol über grundwassernahem Sand ist die Amplitude der Zustandsformen innerhalb eines Standortes (Sauerklee- bis Besenheide- oder Torfmoos-Typ) sehr weit. Auf basen- und nährstoffreicheren Standorten v.a. Waldmeister-Buchenwald (')
            p.add_run('Asperulo odoratae-(Abieti-)Fagetum').italic=True
            p.add_run('), seltener Wimperseggen-Buchenwald (')
            p.add_run('Carici pilosae-Fagetum').italic=True
            p.add_run(') am Rand des Gebiets (z.B. Dunkelsteiner Wald, submontan).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Rotföhrenwälder als Dauergesellschaften an ﬂachgründigen Silikat-Sonderstandorten.')

            p.add_run('Moos-Föhrenwald (')
            p.add_run('Dicrano-Pinetum').italic=True
            p.add_run(') submontan bis tief(-mittel)-montan kleinflächig auf Quarzsand und an flachgründigen Felskuppen; verbreitet auch als Degradationsform auf weniger extremen Standorten. Wachauer Gneis-Föhrenwald (')
            p.add_run('Cardaminopsio petraeae-Pinetum').italic=True
            p.add_run(') kollin-submontan an sonnigen Felsabbrüchen. Serpentin-Föhrenwald (')
            p.add_run('Festuco guestfalicae-Pinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Eschen-Schwarzerlen-Auwälder in der submontanen bis tiefmontanen Stufe.')

            p.add_run('Auf Schwemmböden Waldsternmieren-Schwarzerlenwald (')
            p.add_run('Stellario nemori-Alnetum glutinosae').italic=True
            p.add_run(') mit Bruchweide und Geißfuß (')
            p.add_run('Aegopodium').italic=True
            p.add_run('), an quelligen Stellen (Gleyboden) Winkelseggen-Eschen-Schwarzerlenwald (')
            p.add_run('Carici remotae-Fraxinetum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Schwarzerlen-Bruchwald (')
            p.add_run('Carici elongatae-Alnetum glutinosae').italic=True
            p.add_run(') auf Standorten mit hochanstehendem, stagnierendem Grundwasser (z.B. Teichufer).')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Laubmischwälder mit Esche, Bergahorn, Spitzahorn, Bergulme und Buche submontan bis mittelmontan an frisch-feuchten (Schutt-)Hängen in luftfeuchtem Lokalklima (Grabeneinhänge, Schluchten), z.B. Bingelkraut-Ahorn-Eschenwald (')
            p.add_run('Mercuriali-Fraxinetum').italic=True
            p.add_run('), Geißbart-Ahornwald (')
            p.add_run('Arunco-Aceretum').italic=True
            p.add_run(').')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Fichten-Tannen-Buchenwald (Leitgesellschaft) in der (tief-)mittel-hochmontanen Stufe.')
            p.add_run('Vorwiegend Wollreitgras-Fichten-Tannen-Buchenwald (')
            p.add_run('Calamagrostio villosae-(Abieti-)Fagetum').italic=True
            p.add_run(') auf ärmeren, Quirlzahnwurz-Fichten-Tannen-Buchenwald (')
            p.add_run('Dentario enneaphylli(Abieti-)Fagetum').italic=True
            p.add_run(' mit Kleeschaumkraut (')
            p.add_run('Cardamine trifolia').italic=True
            p.add_run(') und Quirl-Weißwurz (')
            p.add_run('Polygonatum verticillatum').italic=True
            p.add_run(') auf basenund nährstoffreicheren Silikatstandorten.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Montane Fichten- und Fichten-Tannenwälder als edaphisch oder lokalklimatisch bedingte Dauergesellschaften.')

            p.add_run('Waldschachtelhalm-Tannen-Fichtenwald (')
            p.add_run('Equiseto sylvaticiAbietetum').italic=True
            p.add_run(') auf vernäßten Flachhängen (Gleystandorte, “Fichten-Au”). Peitschenmoos-Fichtenwald (')
            p.add_run('Mastigobryo-Piceetum').italic=True

            p.add_run(') mit Torfmoos und Woll-Reitgras (')
            p.add_run('Calamagrostis villosa').italic=True
            p.add_run(') auf anmoorigen Standorten, oft in Inversionslagen; MoorrandFichtenwald.')

            p = doc.add_paragraph('', style='List Bullet')
            p.add_run('Im Bereich von Hochmooren auf Torfböden Fichten-Rotföhrenwald (')
            p.add_run('Vaccinio uliginosi-Pinetum sylvestris').italic=True
            p.add_run(') sowie Latschen- und Moorbirken-Bestände.')


    def fuc_txt_sez_met_1(self, doc):

        doc.add_paragraph('Im Rahmen der Schutzwaldstrategie der ÖBf AG (Stand: 22.10.2018) wird der Erhaltungszustand des Schutzwaldes im Zuge der Forsteinrichtung erhoben. Die Erhebung erfolgt bestandesweise in Anlehnung an die Jagdstrategie der ÖBf AG nach einem Ampelsystem.')

        doc.add_paragraph('')
        p = doc.add_paragraph('')
        p.add_run('GRÜN: ').bold=True
        p.add_run('Die Schutzwirkung ist für die nächsten 20 Jahre gegeben, es besteht kein unmittelbarer Handlungsbedarf. Der Schutzwald ist stabil, gut geschichtet, nicht überaltert, und/oder Naturverjüngung stellt sich ein.')

        doc.add_paragraph('')
        p = doc.add_paragraph('')
        p.add_run('GELB: ').bold=True
        p.add_run('Die Schutzwirkung ist noch gegeben, jedoch werden negative Entwicklungen sichtbar. Es besteht mittelbarer forstlicher/jagdlicher Handlungsbedarf innerhalb der nächsten 20 Jahre. Die Bestände beginnen zu überaltern, Strukturen lösen sich auf, Naturverjüngung stellt sich nicht im gewünschten Maß ein oder fällt aus.')

        doc.add_paragraph('')
        p = doc.add_paragraph('')
        p.add_run('ROT: ').bold=True
        p.add_run('Die Schutzwirkung nimmt zusehends ab. Mehrere negative Faktoren können wirksam werden. Es besteht Handlungsbedarf innerhalb der nächsten 10 Jahre. Überalterte Bestände lösen sich auf, ohne dass Verjüngung nachkommt, das Gelände ist schwierig, es bestehen Belastungen durch Wild und Weidevieh.')

        doc.add_paragraph('')
        doc.add_paragraph('In Vorbereitung zur Forsteinrichtung wurden vor der Einrichtungsperiode die bereits verfügbaren, fortgeschriebenen Manuale automatisiert auf Basis eines definierten Zahlenschlüssels bewertet. Dafür wird der folgende Schlüssel herangezogen.')


    def fuc_txt_sez_met_2(self, doc):

        doc.add_paragraph('In der Gesamtbewertung kann die Dringlichkeit der Schutzwaldverbesserung zwischen den Werten -5 bis +5 liegen. Für die Ampelphasen wurden folgende Grenzwerte festgelegt:')

        doc.add_paragraph('')
        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('GRÜN: ').bold=True
        p.add_run('Dringlichkeit > 2')
        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('GELB: ').bold=True
        p.add_run('Dringlichkeit ≤ 2 und > -2')
        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('ROT: ').bold=True
        p.add_run('Dringlichkeit ≤ -2')

        doc.add_paragraph('')
        doc.add_paragraph('Während der FE soll in jedem Revier vom Forsteinrichter in enger Abstimmung mit dem Betrieb bzw. Revier die aktuelle Ausscheidung der Schutzwälder überprüft und nach dem oben dargestellten Ampelsystem nach Sanierungsdringlichkeit eingestuft werden. Davon ausgehend ergeben sich die jährlichen Planungen für die SW-Maßnahmen. Bei Ansprache der Kategorie „Rot“ (also dringlichem Handlungsbedarf) wurde eine unbedingte Maßnahmenplanung entsprechend der Schutzwaldverbesserung durchgeführt. Durch diese Vorgehensweise ist ein regelmäßiges Up-Date der Schutzwaldsituation durch die Forsteinrichtung bei den ÖBf gewährleistet.')

        doc.add_paragraph('')
        doc.add_paragraph('Maßnahmen, die nicht bestandesweise erhoben werden können, werden abschließend verbal beurteilt. Das betrifft vor allem den Einfluss von Wild, Weide und Käfer, der sich in der Regel nur an größeren Bereichen festmachen lässt.')


    def fuc_txt_sez_wp(self, doc):

        doc.add_paragraph('Der Waldpflegebedarf im Schutzwald mit dem Erhaltungszustand „Rot“ erstreckt sich zum größten Teil auf Aufforstungen (AF) und Ergänzungen (EG). Außerdem wurde, aufgrund der Wildschadenssituation ein entsprechender Anteil an Einzelschutzmaßnahmen gegen Wild (KE) geplant. Das FR 6 Hinteres Zillertal weist hier, aufgrund der umfangreichsten Fläche in dieser Kategorie den höchsten Anteil auf. Darüber hinaus wurde im FR 7 Alpbach, aufgrund der Weidebelastung, ein entsprechend hoher Anteil an flächigem Kulturschutz (KF) geplant. Die Maßnahmenplanung in der Kategorie „Rot“ ist als dringlich zu erachten.')
        doc.add_paragraph('')
        doc.add_paragraph('Flächen in deutlich besseren Zustand (Kategorie „Gelb“) weisen zusätzlich einen höheren Anteil an Jungwuchs- und Dickungspflegen (JP, DP) auf. Auch hier wurden Ergänzungen, Aufforstungen und Maßnahmen zum Kulturschutz geplant. Die Maßnahmenplanung in der Kategorie „Gelb“ sollte mittelfristig durchgeführt werden.')
        doc.add_paragraph('')
        doc.add_paragraph('Die Waldpflegemaßnahmen der Kategorie „Grün“ ergeben sich im Zuge einer entsprechenden Waldbewirtschaftung.')


    def fuc_txt_sez_vn(self, doc):

        doc.add_paragraph('Vornutzungen im Schutzwald wurden im Sinne der Stabilitäts- und Biodiversitätsförderung geplant. Außerdem ergaben sich Nutzungen im Schutzwald häufig in sinnvoller Kombination mit unterliegenden Nutzungen im Wirtschaftswald (Seil bergab). Im Erhaltungszustand „Rot“ können Vornutzungen meist, aufgrund des Alters keinen Beitrag mehr zum Erhalt des Schutzwaldzustandes leisten. Dahingegen stellen Vornutzungen in der Kategorie „Gelb“ einen wertvollen Beitrag zur zukünftigen Bestandesstabilität dar. Hier wurden Vornutzungseingriffe auch verstärkt zum Erhalt von wertvollen Mischbaumarten geplant. Vor allem im am stärksten mit Schutzwald ausgestatteten Revier, dem FR 6 Hinteres Zillertal können entsprechende Mengen realisiert werden.')
        doc.add_paragraph('')
        doc.add_paragraph('In der Nutzungsplanung nehmen naturgemäß seilgebundene Technologien eine zentrale Rolle ein. Sofern ein rechtzeitiger und sinnvoller Eingriff keinen positiven Deckungsbeitrag erwarten lässt, wurde die Maßnahme mit der Rückungsart 90 (Holz verbleibt am Waldort) geplant.')


    def fuc_txt_sez_en(self, doc):

        doc.add_paragraph('Vor allem in der Kategorie „Gelb“ stellen Nutzungen zur Verjüngungseinleitung- und -förderung, wie Lichtungen und Femelungen eine sinnvolle Maßnahme dar. Im Erhaltungszustand „Rot“ können Endnutzung zumeist keinen sinnvollen Beitrag zur Schutzwaldverbesserung leisten.')


    def fuc_txt_naturschutz(self, doc):

        doc.add_paragraph('Die ÖBf tragen besondere Verantwortung für den Erhalt natürlicher Ressourcen und Lebensräume Österreichs. Eine nachhaltige Bewirtschaftung der Flächen wird den wirtschaftlichen Erfordernissen gerecht und erhält ökologisch besonders wertvolle und sensible Gebiete. Im Projekt „Ökologie und Ökonomie“ wurden Strategien entwickelt, um sowohl den wirtschaftlichen als auch den ökologischen Anforderungen an das Flächenmanagement der ÖBf gerecht zu werden. Zur Berücksichtigung von naturschutzfachlichen Aspekten auf der gesamten Fläche der ÖBf wurde im Rahmen von „Ökologie und Ökonomie“ das vorliegende „Ökologische Landschaftsmanagement“ (Ö.L.) eingeführt.')
        doc.add_paragraph('')

        doc.add_paragraph('Das Ökologische Landschaftsmanagement besteht aus vier Handlungsfeldern:')
        doc.add_paragraph('')

        # bullet list
        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('Schutzgutbuch\n').bold = True
        p.add_run('Das Schutzgutbuch umfasst eine Auflistung und Darstellung der hoheitlich verordneten Schutzgebiete sowie Flächen mit Vertragsnaturschutz und nennt Bewirtschaftungserfordernisse und Einschränkungen.').bold = False
        doc.add_paragraph('')

        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('Erhaltung und Renaturierung\n').bold = True
        p.add_run('Dieses Handlungsfeld beinhaltet die Auflistung und Verortung bedeutsamer im Forstrevier vorkommender Lebensraumtypen und Arten und gibt Handlungsempfehlungen für das Flächenmanagement mit konkreten Vorschlägen für Naturschutzmaßnahmen.').bold = False
        doc.add_paragraph('')

        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('Lebensraumvernetzung\n').bold = True
        p.add_run('Hierbei handelt es sich um die langfristige Sicherung eines Netzwerks an hochwertigen Waldstrukturen für ausgewählte, altholzgebundene Zielarten und deren Erhalt im Zuge der forstlichen Nutzungen. Wenn es übergeordnete Raumordnungspläne der Bundesländer zum Thema Lebensraumvernetzung gibt, werden diese hier ebenfalls angeführt.').bold = False
        doc.add_paragraph('')

        p = doc.add_paragraph('', style='List Bullet')
        p.add_run('Prozessschutz\n').bold = True
        p.add_run('Dieses Handlungsfeld enthält die Auflistung und Beschreibung bestehender Prozessschutzflächen, die von kleineren Biodiversitätsinseln über Naturwaldreservate bis hin zu großflächigen Wildnisgebieten reichen können. Außerdem werden Wildnis-Potenzialflächen, die sich aus einer GIS-Daten-Analyse des WWF 2016 ergeben haben, beschrieben. Sie dienen als Information und Diskussionsgrundlage bei geplanten Erschließungs- und Infrastrukturprojekten.').bold = False
        doc.add_paragraph('')

        doc.add_paragraph('Ö.L. ist ein Werkzeug für im Revier tätige MitarbeiterInnen, um Schutzgüter zu erkennen und zu erhalten. Es hilft bei der Umsetzung von Naturschutzmaßnahmen vor Ort. Selbstständig oder mit Unterstützung der KollegInnen vom Kompetenzfeld Naturschutz und dem Naturraummanagement können daraus Ausgleichsmaßnahmen und Projekte abgeleitet, entwickelt und umgesetzt werden. Mit Hilfe des Ö.L. sollen auch Konfliktfelder mit Naturschutz und Stakeholdern frühzeitig erkannt und wenn möglich vermieden werden.')
        doc.add_paragraph('')

        doc.add_paragraph('Durch die Zusammenarbeit zwischen dem Kompetenzfeld Naturschutz, dem Naturraummanagement, Forsteinrichtung und den RevierleiterInnen sowie den ForstbetriebsleiterInnen soll sichergestellt werden, dass alle Handlungsfelder in der langfristigen Planung und auch in der täglichen Umsetzung berücksichtigt werden.')
        doc.add_paragraph('')

        doc.add_paragraph('Die wichtigsten Informationen zu den jeweiligen Handlungsfeldern sind im Schutzgutbuch zu finden das im Forstbetrieb auch analog abgelegt wird.')
        doc.add_paragraph('')

        doc.add_paragraph('Bei Fragen stehen die KollegInnen aus dem Kompetenzfeld Naturschutz und dem Naturraummanagement gerne zur Verfügung.')
        doc.add_paragraph('')
