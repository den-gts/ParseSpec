# -*-coding: utf-8 -*-
import pywintypes,win32com.client,logging,os
from lxml import etree
from settings import Sections

class WordSpecification:
    def __init__(self,fileName):
        self.__Headers={
            "Position":u"поз.",
            "Description":u"обозначение",
            "Name":u"наименование"}

        self.__word=win32com.client.Dispatch("Word.Application")
        self.__word.Visible=0
        self.__wdoc=self.__word.Documents.Open(fileName)

    def getCell(self,table,row,col):
        result=unicode(table.Cell(row,col).Range.Text)[:-2].strip()
        return result

    def getColumnNames(self,table):
        currentIndex=1
        headersIndexs={}
        while currentIndex:

            try:
                currentHeader=self.getCell(table,1,currentIndex)
            except pywintypes.com_error:

                break
            for HeaderName,HeaderString in self.__Headers.items():
                if HeaderString==currentHeader.lower():
                    headersIndexs.update(dict([(HeaderName,currentIndex)]))
            currentIndex+=1
        headersIndexs['lastColumn']=currentIndex
        return headersIndexs

    def __RawCol(self,row,table,columnNames,col):#функция вывода сырого столбца без обработки
            return self.getCell(table,row,columnNames[col])

    def __del__(self):
        self.__wdoc.Close(False)
        self.__word.Quit()
    def __funcRow(self,funct,**kwargs):#унифицированная функция-обертка для манипуляций над строкой
        result=[]
        for table in self.__wdoc.Tables:
            self.__columnNames=self.getColumnNames(table)

            for row in range(2,40):
                try:
                    result.append(funct(row,table,**kwargs))
                except KeyError,er:
                    pass
                except pywintypes.com_error,er:
                #print "error: "+er[2][2]
                    break
        return result
    def rawCol(self):#обертка для функции вывода сырого столбца
        return self.__funcRow(self.__RawCol,col="Name")
    def __rawRow(self,row,table):#функция вывода сырой строки
        result=[]
        for column in range(self.__columnNames['lastColumn']):
            result.append(self.getCell(table,row,column))
        return result

    def getRawRows(self):#возвращает сырые строки
        return self.__funcRow(self.__rawRow)

    def __rwParceToXML(self,rowNumber,table,**kwarg):#парсит отдельную строку
        #вытаскиваем сырую строку из док файла
        row=self.__rawRow(rowNumber,table)#сырая строка
        ColNamesWithoutLast=self.__columnNames.copy()#TODO костылечек с удалением lastColumn из словаря имен колонок
        ColNamesWithoutLast.pop('lastColumn')
        row={key:row[value] for key,value in ColNamesWithoutLast.items() }#формируем новый словарь для краткости

        #собираем строку. Если многострочный элемент, объеденяем в соответсвии с условиями
        #вытаскиваем значение колонки "Наименование" и кладем в буффер до условия новой строки
        if not row['Name'] and not kwarg['lstBuffer']:return#игнорируем пустые строки
        lstParentLstbuffer=(self.__section,kwarg['lstBuffer'],kwarg['dicColumns'])
        #TODO доделать обработку разделов
        ##если имя совпадает с выражением раздела, то добавляем новый раздел
        sections=Sections()
        sectionName=sections.compareSection(row['Name'])
        if sectionName or (row['Name'] and table.Cell(rowNumber,self.__columnNames['Name']).Range.Font.Underline):
        #обнаружено начало нового раздела. Добавляем буффер прошлого элемента если не пуст
            self.addXMLelement(*lstParentLstbuffer)
            self.addSection(sectionName,row['Name'])
            return

        if not row['Name']:
            #обнаружена пустая строка - записываем буффер в элемент XML,очищаем буффер и выходим из функции
            self.addXMLelement(*lstParentLstbuffer)

            return
        #добавляем элемент в XML если встречаем признак нового эдемента - позицию или обозначение
        #и заносим текющее наименоваине в буфер
        if (row['Position'] or row['Description']) and kwarg['lstBuffer']:
            self.addXMLelement(*lstParentLstbuffer)
        #если новый элемент, то формируем словарь значений столбцов.
        #значения столбцов определяются в первой строке элемента. Обозначение, позиция и.т.п.
        # Пустые значения и наименование нам не нужны
        #TODO а если поле "примечание" многострочное?
        if not kwarg['dicColumns']:
            kwarg['dicColumns'].update({atr:row[atr] for atr in row.keys() if (row[atr] and  atr!='Name')})
        #TODO написать регулярку для добавления точки в тексте в случаях подобных:Руководство по эксплуатации Лист утверждения
        kwarg['lstBuffer'].append("%s%s"%(" ",row['Name']))#условий нового элемента ненайдено. Пополняем стек


    def addXMLelement(self,parent,lstBuffer=None,dicColumns=None):
        element=None
        #и наконец добавляем буффер с атрибутами в дерево XML
        if lstBuffer or dicColumns:
            element=etree.SubElement(parent,"element")
            if dicColumns:
                for columnName,columnValue in dicColumns.items():
                    etree.SubElement(element,columnName).text=columnValue
                    #чистим словарь для нового элемента
                    dicColumns.clear()

            #, попутно удаляя дефисы и склеивая строки
            etree.SubElement(element,"Name").text="".join(lstBuffer).replace("- ","")
            #чистим буфер
            del lstBuffer[:]

    def addSection(self,sectionName,defaultSectionName="раздел без имени"):
        self.__section=self.root
        self.__section=etree.SubElement(self.__section,"section")
        if not sectionName:
            sectionName=defaultSectionName
        etree.SubElement(self.__section,'name').text=sectionName

    def getXML(self):#функция парсинга документа в XML
        root=etree.Element("specification")#корневой элемент XML
        self.root=root
        self.__section=root#текущий раздел корень
        self.__funcRow(self.__rwParceToXML,lstBuffer=list(),dicColumns=dict())#запускаем функцию обработки строк
        return etree.tostring(root,pretty_print=True,encoding='utf-8', xml_declaration=True)



if __name__=='__main__':
    logging.basicConfig(level=logging.DEBUG)#filename='f:\\log',filemode="w")

    Wspec=WordSpecification(u'%s\\СКИД.461411.001 АРК.doc'%os.getcwd())
    xmlfile=open('output.xml','w')
    xmlfile.write(Wspec.getXML())
    xmlfile.close()

