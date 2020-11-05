import os
import io
import zipfile
from tempfile import mkdtemp
import shutil
import datetime as dt
from decimal import Decimal

class XlsxZadania:
    def __init__(self, nowy_xlsx, trybtablicaStr,clvl):
        self.tempdir = mkdtemp()  # folder tymczasowy
        self.__nowy_xlsx = nowy_xlsx
        self._arkuszlista = []
        self._arkusz_numer = -1
        self.trybtablicaStr = trybtablicaStr
        self.listaArkuszyUkrytych = set()
        self.maksszerokosc = 80
        self.kolekcjaDanych = []
        self.liczbanapisow = 0
        self.__inicjujlitery()
        self.__numerstringuuniklany = 0
        self.__tabstrDic = {}
        self.__inicjuj_foldery()
        self.szerokosc_datetime = 15.28515625
        self.szerokosc_date = 10.140625
        self.clvl = clvl

    def dodajArkusz(self,  *listaArkuszy):
        self._arkuszlista.extend([(ark, self.tempdir + r"\xl\worksheets\{0}.xml".format(ark)) for ark in listaArkuszy])

    @classmethod
    def nowy(cls, nowy_xlsx, trybtablicaStr=True, clvl=5):
        return cls(nowy_xlsx, trybtablicaStr, clvl)

    def __inicjujlitery(self):
        self.litery = [chr(p) for p in range(65, 91)]
        temp = []
        for q in self.litery:
            temp += [q + chr(p) for p in range(65, 91)]
        temp2 = []
        for r in filter(lambda x: x <= "XF", temp):
            temp2 += [r + chr(p) for p in range(65, 91)]
        self.litery.extend(temp)
        self.litery.extend(temp2)
        del temp2
        del temp

    def __tworz_dodatkowe_pliki(self):


        f1 = open(self.tempdir + r"\[Content_Types].xml", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types '
                 'xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                 '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                 '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                 '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                 '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                 ''.join(['<Override PartName="/xl/worksheets/{:}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(x[0])
                         for x in self._arkuszlista])
                +'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                 '</Types>')
        f1.close()

        f1 = open(self.tempdir + r"\docProps\app.xml", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties '
                 'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
                 'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application'
                 '>Krzysiek - arkusz</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs'
                 '><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt'+
                 ':variant><vt:i4>{0}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
                 '<TitlesOfParts>'
                 '<vt:vector size="{0}" baseType="lpstr">'.format(len(self._arkuszlista))
                 + ''.join(['<vt:lpstr>{0}</vt:lpstr>'.format(x[0]) for x in self._arkuszlista])
                 +'</vt:vector>'
                 '</TitlesOfParts>'
                 '<Company></Company><LinksUpToDate>false</LinksUpToDate'
                 '><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>12.0000'
                 '</AppVersion></Properties>')
        f1.close()

        f1 = open(self.tempdir + r"\xl\workbook.xml", 'w', encoding="UTF-8")



        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook '
                 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion '
                 'appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr '
                 'defaultThemeVersion="124226"/><bookViews><workbookView xWindow="240" yWindow="15" '
                 'windowWidth="16095" windowHeight="9660"/></bookViews>'
                 '<sheets>'+
                 ''.join(['<sheet name="{0}" sheetId="{1}"{2} r:id="rId{1}"/>'.
                         format(x[1][0], x[0]+1,' state ="hidden"' if x[1][0] in self.listaArkuszyUkrytych and x[0] !=0 else '') for x in enumerate(self._arkuszlista)])
                 +
                 '</sheets>'
                 '<calcPr calcId="124519" fullCalcOnLoad="1"/></workbook>')
        f1.close()


        f1 = open(self.tempdir + r"\xl\_rels\workbook.xml.rels", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + ''.join(['<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/{1}.xml"/>'. format(x[0]+1, x[1]) for x in enumerate(map(lambda a1:a1[0], self._arkuszlista))])
            + '<Relationship Id="rId{:}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
            '<Relationship Id="rId{:}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '<Relationship Id="rId{:}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            .format(len(self._arkuszlista)+1, len(self._arkuszlista)+2, len(self._arkuszlista)+3)
            + '</Relationships>')
        f1.close()

        # to bez żadnych zmian
        f1 = open(self.tempdir + r"\_rels\.rels", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships '
                 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
                 'Target="xl/workbook.xml"/>'
                 '<Relationship Id="rId2" '
                 'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
                 'Target="docProps/core.xml"/>'
                 '<Relationship Id="rId3" '
                 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
                 'Target="docProps/app.xml"/>'
                 '</Relationships>')
        f1.close()

        # tu modufikujacy itp. można potem dopracować
        f1 = open(self.tempdir + r"\docProps\core.xml", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties '
                 'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
                 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator></dc:creator><cp:lastModifiedBy'
                 '></cp:lastModifiedBy><dcterms:created '
                 'xsi:type="dcterms:W3CDTF">2020-07-20T05:51:26Z</dcterms:created><dcterms:modified '
                 'xsi:type="dcterms:W3CDTF">2020-07-20T05:51:26Z</dcterms:modified></cp:coreProperties>')
        f1.close()

        # tu style
        f1 = open(self.tempdir + r"\xl\styles.xml", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet '
                 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz '
                 'val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme '
                 'val="minor"/></font></fonts><fills count="2"><fill><patternFill '
                 'patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders '
                 'count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs '
                 'count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                 '<cellXfs count="3">'
                 '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                 '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'  # ten jest do daty
                 '<xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'  # ten jest do datyczasu
                 '</cellXfs>'
                 '<cellStyles '
                 'count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/>'
                 '</cellStyles>'
                 '<dxfs '
                 'count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" '
                 'defaultPivotStyle="PivotStyleLight16"/></styleSheet>')
        f1.close()

        #motyw - nic tu nie ruszaj
        f1 = open(self.tempdir + r"\xl\theme\theme1.xml", 'w', encoding="UTF-8")
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme '
                 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office '
                 'Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" '
                 'lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr '
                 'val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr '
                 'val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr '
                 'val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr '
                 'val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr '
                 'val="0000FF"/></a:hlink><a:folHlink><a:srgbClr '
                 'val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin '
                 'typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="ＭＳ '
                 'Ｐゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font '
                 'script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font '
                 'script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font '
                 'script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" '
                 'typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" '
                 'typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" '
                 'typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" '
                 'typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font '
                 'script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" '
                 'typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" '
                 'typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" '
                 'typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" '
                 'typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" '
                 'typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft '
                 'Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs '
                 'typeface=""/><a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/><a:font script="Hang" typeface="맑은 '
                 '고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/><a:font '
                 'script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" '
                 'typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" '
                 'typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" '
                 'typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" '
                 'typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" '
                 'typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font '
                 'script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font '
                 'script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" '
                 'typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" '
                 'typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" '
                 'typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" '
                 'typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" '
                 'typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme '
                 'name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill '
                 'rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod '
                 'val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint '
                 'val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr '
                 'val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin '
                 'ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs '
                 'pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod '
                 'val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade '
                 'val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr '
                 'val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin '
                 'ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" '
                 'cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod '
                 'val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" '
                 'cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash '
                 'val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr '
                 'val="phClr"/></a:solidFill><a:prstDash '
                 'val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw '
                 'blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha '
                 'val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a'
                 ':outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr '
                 'val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a'
                 ':effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" '
                 'rotWithShape="0"><a:srgbClr val="000000"><a:alpha '
                 'val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera '
                 'prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" '
                 'dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" '
                 'h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr '
                 'val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr '
                 'val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs '
                 'pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod '
                 'val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade '
                 'val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path '
                 'path="circle"><a:fillToRect l="50000" t="-80000" r="50000" '
                 'b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr '
                 'val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs '
                 'pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod '
                 'val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" '
                 't="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a'
                 ':themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>')
        f1.close()

    def __spakuj(self):
        ziph = zipfile.ZipFile(self.__nowy_xlsx, 'w')
        for root, dirs, files in os.walk(self.tempdir):
            for file in files:
                ziph.write(os.path.join(root, file), os.path.join(root, file).replace(self.tempdir, ""), compress_type=zipfile.ZIP_DEFLATED, compresslevel= self.clvl)
        ziph.close()
        shutil.rmtree(self.tempdir)

    def __spakuj7z(self, tryb="-mx2"):
        path_7zip = r"C:\Program Files\7-Zip\7z.exe"

        if os.path.exists(path_7zip):
            import subprocess
            path_working = self.tempdir
            outfile_name = self.__nowy_xlsx
            os.chdir(path_working)
            p = subprocess.Popen([path_7zip, "a", tryb, "-tzip", "-sdel", outfile_name])
            p.wait()
        else:
            print("nie znaleziono 7zipa",path_7zip)
            self.__spakuj()

    def zamknij(self, pakowanie="std"):
        try:
            self._zapisz_shared_strings()
            self.__tworz_dodatkowe_pliki()
            if pakowanie == "std":
                self.__spakuj()
            elif pakowanie in ("7z","7z-Low"):
                self.__spakuj7z()
            elif pakowanie == "7z-Normal":
                self.__spakuj7z("-mx5")
            elif pakowanie == "7z-maximum":
                self.__spakuj7z("-mx7")
            elif pakowanie == "7z-ultra":
                self.__spakuj7z("-mx9")
            else:
                print("parametr z poza listy pakowanie standardowe")
                self.__spakuj()
        except PermissionError as e:
            print(e)
            try:
                self.__nowy_xlsx = self.__nowy_xlsx+"UWAGA"
                self.__spakuj()
            except Exception as e2:
                shutil.rmtree(self.tempdir)
                print(e2)
                print(self.tempdir, " - skasowane")

    def zapisz(self, daneTabelaryczne, naglowki=None): #args = tupla kolekcji danych
        self._arkusz_numer += 1
        self.kolekcjaDanych.clear()
        if naglowki:
            self.kolekcjaDanych.append([naglowki])

        self.kolekcjaDanych.append(daneTabelaryczne)
        try:
            self.__zapiszZakladke()
        except Exception as e:
            print(e)
            shutil.rmtree(self.tempdir)
            print(self.tempdir, " - skasowane")

    def __kolekcjaX(self):
        for dane in self.kolekcjaDanych:
            for p in dane:
                yield p

    def __zapiszZakladke(self):

        with open(self._arkuszlista[self._arkusz_numer][1], "w", encoding="UTF-8") as f2:
            wierszy, kolumn, szerokosc_kolumn = self.__zapisz_zakladka_glowne_dane_arkusza(f2)

        with os.fdopen(os.open(self._arkuszlista[self._arkusz_numer][1], os.O_RDWR | os.O_CREAT), 'r+',encoding="UTF-8") as f1:
            self.__zapisz_zakladka_dane_wstepne(f1, kolumn, szerokosc_kolumn, wierszy)


        # with open(self._arkuszlista[self._arkusz_numer][1]+"cz2", "w", encoding="UTF-8") as f2:
        #     wierszy, kolumn, szerokosc_kolumn = self.__zapisz_zakladka_glowne_dane_arkusza(f2)
        #
        # with open(self._arkuszlista[self._arkusz_numer][1]+"cz1", "w", encoding="UTF-8") as f1:
        #     self.__zapisz_zakladka_dane_wstepne(f1, kolumn, szerokosc_kolumn, wierszy)
        #
        #     # scalanie dwóch plików
        # with open(self._arkuszlista[self._arkusz_numer][1], "wb") as wfd:
        #     for f in [self._arkuszlista[self._arkusz_numer][1]+"cz1",
        #                 self._arkuszlista[self._arkusz_numer][1]+"cz2"]:
        #         with open(f, 'rb') as fd:
        #             shutil.copyfileobj(fd, wfd)
        # os.remove(self._arkuszlista[self._arkusz_numer][1] + "cz1")
        # os.remove(self._arkuszlista[self._arkusz_numer][1] + "cz2")

    def __zapisz_zakladka_dane_wstepne(self, f1, kolumn, szerokosc_kolumn, wierszy):
        f1.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        f1.write(r'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        f1.write('<dimension ref="A1:{0}{1}"/>'.format(self.litery[kolumn - 1], wierszy))
        if self._arkusz_numer == 0:
            f1.write(
                '<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>')
        else:
            f1.write('<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>')
        szerokosc_kolumn = enumerate(szerokosc_kolumn)
        szerokosc_kolumn = list(filter(lambda l: l[1] != 1.0, szerokosc_kolumn))
        if len(szerokosc_kolumn) > 0:
            f1.write('<cols>')
            for x in szerokosc_kolumn:
                if x[1] == 1.0:
                    continue
                f1.write('<col min="{0}" max="{0}" width="{1}" bestFit = "1" customWidth="1" />'.format(x[0] + 1, min(x[1], self.maksszerokosc)))
            f1.write('</cols>')
        f1.write(r'<sheetData>')

    def __zapisz_zakladka_glowne_dane_arkusza(self, f):
        kolumn = 0
        szerokosc_kolumn = [1.0]
        bufor_pliku = []
        # glowna tabela danych
        for i, wierszDanych in enumerate(self.__kolekcjaX()):
            if i == 0:
                kolumn = len(wierszDanych)
                szerokosc_kolumn = [1.0 for i in range(kolumn)]
                for _ in range(600+100*kolumn):
                    f.write(" ")
                f.write("\n")


            bufor_pliku.append('<row r="{0}" spans="1:{1}">'.format(i + 1, kolumn))
            self.__zapisz_wiersz(bufor_pliku, i, szerokosc_kolumn, wierszDanych)
            bufor_pliku.append('</row>')
            if i % 100 == 0:
                f.write(''.join(bufor_pliku))
                bufor_pliku.clear()

        f.write(''.join(bufor_pliku))
        bufor_pliku.clear()
        f.write(r'</sheetData>')
        f.write(r'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
        return i + 1, kolumn, szerokosc_kolumn

    def __zapisz_wiersz(self, bufor_pliku, i, szerokosc_kolumn, wierszDanych):
        for j, dana in enumerate(wierszDanych):
            if isinstance(dana, (int, float, Decimal)):
                bufor_pliku.append('<c r="{:}{:}"><v>{:}</v></c>'.format(self.litery[j], i + 1, dana))
            elif isinstance(dana, dt.datetime):
                if szerokosc_kolumn[j] < self.szerokosc_datetime:
                    szerokosc_kolumn[j] = self.szerokosc_datetime
                bufor_pliku.append('<c r="{:}{:}" s="2"><v>{:}</v></c>'.format(self.litery[j], i + 1, self.__excel_date(
                    dana)))  # zapis ze wskazaniem stylu
            elif isinstance(dana, dt.date):
                if szerokosc_kolumn[j] < self.szerokosc_date:
                    szerokosc_kolumn[j] = self.szerokosc_date
                bufor_pliku.append('<c r="{:}{:}" s="1"><v>{:}</v></c>'.format(self.litery[j], i + 1, self.__excel_date(
                    dana)))  # zapis ze wskazaniem stylu
            elif isinstance(dana, str):
                if szerokosc_kolumn[j] < len(dana) * 1.25 + 2.0:
                    szerokosc_kolumn[j] = len(dana) * 1.25 + 2.0
                if "&" in dana:
                    dana = dana.replace("&","&amp;")
                if "<" in dana:
                    dana = dana.replace("<", "&lt;")
                if ">" in dana:
                    dana = dana.replace(">", "&gt;")
                if "\"" in dana:
                    dana = dana.replace("\"", "&quot;")
                if "'" in dana:
                    dana = dana.replace("'", "&apos;")


                if self.trybtablicaStr and len(dana) > 0 and (dana[0] == '\t' or dana[0] == ' '):
                    dana = f'<t xml:space="preserve">{dana}</t>'
                elif self.trybtablicaStr:
                    dana = f'<t>{dana}</t>'

                if self.trybtablicaStr and self.__tabstrDic.get(dana, -1) == -1:
                    self.__tabstrDic[dana] = self.__numerstringuuniklany
                    bufor_pliku.append(r'<c r="{:}{:}" t="s"><v>{:}</v></c>'.format(self.litery[j], i + 1, self.__numerstringuuniklany))
                    self.__numerstringuuniklany += 1
                elif self.trybtablicaStr:
                    bufor_pliku.append(r'<c r="{:}{:}" t="s"><v>{:}</v></c>'.format(self.litery[j], i + 1, self.__tabstrDic[dana]))
                else:  # tryb - napisy inline
                    if len(dana) > 0 and (dana[0] == '\t' or dana[0] == ' '):
                        bufor_pliku.append(r'<c r="{:}{:}" t="inlineStr"><is><t xml:space="preserve">{:}</t></is></c>'.format(self.litery[j], i + 1, dana))
                    else:
                        bufor_pliku.append(r'<c r="{:}{:}" t="inlineStr"><is><t>{:}</t></is></c>'.format(self.litery[j], i + 1, dana))
                self.liczbanapisow += 1
            else:
                f'<c r="{self.litery[j]}{i + 1}" t="inlineStr"><is><t>zglos krzyskowi problem z eksportem !{str(type(dana))}</t></is></c>'

    def _zapisz_shared_strings(self):
        with open(self.tempdir + r"\xl\sharedStrings.xml", 'w', encoding="UTF-8") as napisy:
            napisy.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
            napisy.write('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
            napisy.write(' count="{0}" uniqueCount="{1}">'.format(self.liczbanapisow, len(self.__tabstrDic)))
            if self.trybtablicaStr:
                for x in self.__tabstrDic.keys():
                    napisy.write(f'<si>{x}</si>')
            napisy.write('</sst>')

    def __inicjuj_foldery(self):
        os.mkdir(self.tempdir + r"\xl")
        os.mkdir(self.tempdir + r"\xl\worksheets")
        os.mkdir(self.tempdir + r"\_rels")
        os.mkdir(self.tempdir + r"\docProps")
        os.mkdir(self.tempdir + r"\xl\_rels")
        os.mkdir(self.tempdir + r"\xl\theme")

    def __excel_date(self, date1):
        if isinstance(date1, dt.datetime):
            temp = dt.datetime(1899, 12, 30)  # Note, not 31st Dec but 30th!
        else:
            temp = dt.date(1899, 12, 30)
        delta = date1 - temp
        return float(delta.days) + (float(delta.seconds) / 86400)
