# -*-coding: utf-8 -*-
import re
from lxml import etree
class Sections():
    def __init__(self,root=etree.parse('sections.xml')):
        self.root=root
    def compareSection(self,cmpString):
        for regExp in self.root.xpath('./section[name][regExp]/regExp'):
            if regExp.text:
                regExpPattern=re.compile(regExp.text)
                if re.match(regExpPattern,cmpString):
                    #return regExp.xpath('..')[0]
                    return regExp.xpath('../name')[0].text

if __name__=='__main__':
    sect=Sections()
    print type(sect.compareSection(u"детали"))