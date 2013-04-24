# -*- coding: utf-8 -*-
#TODO В разделе Стандартные изделия ввести подразделы типа "болты ГОСТ 343443"
import win32com.client
import pywintypes,logging,re
class SuperSection:
    def __init__(self):
        self.names=(u'переменные данные для исполения',
            u"устанавливать по ")
class SubSection:
    def __init__(self):
        self.names=(u"Болты")

class Sections:
    def __init__(self):
        self.names=(u'документация',u'сборочные единицы',
          u'детали',u'стандартные изделия',
          u"Материалы:",
          u'прочие изделия',
          u'комплекты',
        )

        constNames=('DOCUMENTATION',"ASSAMBLYS","PARTS","STANDART_PARTS","MATERIALS","OTHER","KITS")
        self.__dict__.update(dict(zip(constNames,self.names)))#записываем константы
        #например для сборочный едениц self.ASSAMBLYS=="сборочные единицы"


logging.basicConfig(level=logging.DEBUG)
class SpecElement:
    def __init__(self,name,section,position=0,description=""):
        self.section=section
        self.name=name
        self.position=position
        self.description=description
    def __str__(self):
        strOut="%s %s"%(self.description,self.name)
        return strOut
    def __repr__(self):
        strOut="<%s:%s %s>"%(self.__class__.__name__,self.description,self.name)
        return strOut


class DocSpecification:
    def __init__(self):
        self.SECTION_NAMES=Sections()
        self.spElement=[]
        self.section={}
        self._wdoc=object()
        for row in parseDoc():
            self.spElement.append(SpecElement(row['Name'],row['Section'],row['Position'],row['Description']))
            lastRow=self.spElement[-1]
            if lastRow.section in self.SECTION_NAMES.names:
                if not self.section.has_key(lastRow.section):
                    self.section[lastRow.section]=[]
                self.section[lastRow.section].append(lastRow)

    def openWord(self):
        self._word=win32com.client.Dispatch("Word.Application")
        self._word.Visible=0
    def __delete__(self, instance):
        if self._wdoc: self._wdoc.Close(False)
        if self._word: self._word.Quit()
    def openDoc(self,DocFile=u'D:\\project\\python\\com\\СКИД.461411.001 АРК.doc'):
        self._wdoc=self._word.Documents.Open(DocFile)
    def closeDoc(self):
        self._wdoc.Close(False)
def getCell(table,row,col):
    return unicode(table.Cell(row,col).Range.Text)[:-2].strip()
#TODO неудачное название функции
def checkDefis(fcell, FirstRow=True):#функция обработки переноса
    fcell['Name']=fcell['Name'].strip()
    if fcell['Name'][-1:]=="-":
        fcell['Name']="".join((fcell['Name'][:-1],"\\-"))
    else:
        afterChar=""
        beforeChar=""
        if not FirstRow:
            if fcell['Name'][:1].isupper():#Если строка начинается с заглавной буквы и не первая то ставим
                                                            # перед ней точку
                beforeChar=". "
            else:
                beforeChar=" "

        fcell['Name']="%s%s%s"%(beforeChar,fcell['Name'],afterChar)
    return fcell
def columnNames(table):####определять колонки по названию.. в разных шаблонах разные номера колонок#####
    Headers={
        "Position":u"поз.",
        "Description":u"обозначение",
        "Name":u"наименование"}
    headersIndexs={}
    currentIndex=1
    while currentIndex:

        try:
            currentHeader=getCell(table,1,currentIndex)
        except pywintypes.com_error:
            break
        for HeaderName,HeaderString in Headers.items():
            if HeaderString==currentHeader.lower():
                headersIndexs.update(dict([(HeaderName,currentIndex)]))
        currentIndex+=1
    return headersIndexs
def loging2file(rows,logfile='doc.log'):
    #лог в файл распарсенных строк
    docLog=open(logfile,'w')

    for row in rows:
        if row['Position']==u"Лист":logging.debug('Позиция \"лист\"')
        line2file=("%s:\t%s\t%s\t%s\n")%(row['Section'],row['Position'],row['Description'],row['Name'])
        docLog.write (line2file.encode('utf-8'))
    docLog.close()
def parseDoc(DocFile=u'D:\\project\\python\\com\\СКИД.461411.001 АРК.doc',logFlag=True):

    word=win32com.client.Dispatch("Word.Application")
    word.Visible=0


    wdoc=word.Documents.Open(DocFile)
    rows=[]
    #TODO: сделать регулярные выражения типа "детали:"
    currentSection=""

    for table in wdoc.Tables:

        #############################################

        for row in range(2,40):
            cell={}
            try:#TODO проверка на наличие в словаре headerIndex ключей и что делать если нет ключа
                for HeaderName,HeaderIndex in columnNames(table).items():
                    cell[HeaderName]=getCell(table,row,HeaderIndex)

            except pywintypes.com_error,er:
                logging.debug("end table: "+er[2][2])
                break
            #проверка на не пустое имя и является ли позиция цифрой(защита от паразитных "лист")

            DescriptionPattern=re.compile(r"^[^\*]|$")#защита от примечаний типа "* из комплекта"

            if cell and cell['Name'] \
                                    and DescriptionPattern.match(cell['Description']) \
                                    and (cell['Position'].strip().isdigit() or not cell['Position']):

                sections=Sections()
                for secItem in sections.names:#определение раздела

                    if cell['Name'].lower()==secItem:
                        currentSection=secItem
                        break

                if cell['Name'].lower()!=currentSection:
                    if cell['Position'] or cell['Description']:#первая строка
                        cell['Section']=currentSection
                        #print "%s|%s"%(cell['Name'],cell['Name'].strip()[-1:])

                        rows.append(checkDefis(cell))

                    else:#последующие строки

                        cell=checkDefis(cell,False)

                        if len(rows):
                            lastRowName="%s%s"%(rows[len(rows)-1]['Name'],cell['Name'])
                            lastRowName=lastRowName.replace(u'\\- ',u'')
                            rows[len(rows)-1]['Name']=lastRowName
    if logFlag: loging2file(rows)
    wdoc.Close(False)
    word.Quit()
    return rows

if __name__=='__main__':
    parseDoc()