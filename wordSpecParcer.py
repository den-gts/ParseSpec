# -*-coding: utf-8 -*-

import re
class Section:
    def __init__(self,regularExp):
        self.__regExp=regularExp
        self.members=[]
    def add(self,item):
        self.members.append(item)
    def __getitem__(self, item):
        return self.members[item]

reSection=re.compile(r"^сборочные единицы")

section=Section(r"сборочные единицы")
section.add('test')
section.add('сборочный чертеж')
section.add(Section('Болт'))
for i in section:
    print i
