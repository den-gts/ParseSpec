# -*-coding: utf-8 -*-
import pywintypes,win32com.client
from lxml import etree
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
        return unicode(table.Cell(row,col).Range.Text)[:-2].strip()

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
            columnNames=self.getColumnNames(table)

            for row in range(2,40):
                try:
                    result.append(funct(row,table,columnNames,**kwargs))
                except KeyError,er:
                    pass
                except pywintypes.com_error,er:
                #print "error: "+er[2][2]
                    break
        return result
    def rawCol(self,col='Name'):#обертка для функции вывода сырого столбца
        return self.__funcRow(self.__RawCol,col="Name")
    def __rawRow(self,row,table,columnNames):#функция вывода сырой строки
        result=[]
        for column in range(columnNames['lastColumn']):
            result.append(self.getCell(table,row,column))
        return result
    def getRawRows(self):#возвращает сырые строки
        return self.__funcRow(self.__rawRow)
    def __rwParceToXML(self,row,table,columnNames,**kwarg):#парсит отдельную строку
        #вытаскиваем сырую строку из док файла
        row=self.__rawRow(row,table,columnNames)#сырая строка
        ColNamesWithoutLast=columnNames.copy()#TODO костылечек с удалением lastColumn из словаря имен колонок
        ColNamesWithoutLast.pop('lastColumn')
        row={key:row[value] for key,value in ColNamesWithoutLast.items() }#формируем новый словарь для краткости

        elemText=row['Name']#вытаскиваем значение колонки "Наименование"
        etree.SubElement(kwarg['parent'],"element").text=elemText
    def getXML(self):#функция парсинга документа в XML
        root=etree.Element("specification")#корневой элемент XML
        section=root#пока будет один раздел - кореневой
        self.__funcRow(self.__rwParceToXML,parent=section)#запускаем функцию обработки строк
        return etree.tostring(section,pretty_print=True,encoding='utf-8', xml_declaration=True)

if __name__=='__main__':
    Wspec=WordSpecification(u'D:\\project\\python\\com\\СКИД.461411.001 АРК.doc')
    print Wspec.getXML()


