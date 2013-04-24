# -*-coding: utf-8 -*-

import re

class Section:
    def __init__(self,name):
        self.name=name
        self.members=[]
    def add(self,item):
        self.members.append(item)
        return self.lastMember()
    def __getitem__(self, item):
        return self.members[item]
    def __repr__(self):
        return "section: '%s'(%i members)"%(self.name,len(self))
    def __len__(self):
        return len(self.members)
    def getOnlyItem(self):
        items=[]
        for item in self.members:
            if not isinstance(item,Section):items.append(item)
        return items
    def getSubsections(self):
        sections=[]
        for section in self.members:
            if isinstance(section,Section):sections.append(section)
        return sections
    def lastMember(self):
        lastIndex=len(self.members)-1
        if lastIndex>0:return self.members[lastIndex]
        else:return 0



def printSection(section,offset=0):
    for member in section:
        if isinstance(member,Section):
            offset+=1
            print "%s===%s==="%("\t"*offset,member)
            printSection(member,offset)
        else:print "%s%s"%("\t"*offset,member)
if __name__=='__main__':
    reSection=re.compile(r"^сборочные единицы")
    #testfile=open("testfile.txt","r")
    #line=testfile.readline()
    noNameSection=Section(".")

    section=Section(r"сборочные единицы")
    section.add('test')
    section.add(u'сборочный чертеж')
    section.add(Section('Болт'))
    subsection=section.getSubsections()[0]
    subsection.add('sub')
    subsection.add('Болт2')
    subsection=subsection.add(Section('Винт'))
    subsection.add('винт1')
    printSection(section)
    #testfile.close()
