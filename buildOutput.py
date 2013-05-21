# -*-coding: utf-8 -*-
import os
from jinja2 import Environment
from jinja2.loaders import FileSystemLoader
from wordSpecification import WordSpecification
from lxml import etree,objectify
env=Environment(loader=FileSystemLoader(os.getcwd()))
template=env.get_template("structure.html")
spec=WordSpecification((u'%s\\СКИД.461411.001 АРК.doc'%os.getcwd()))
XML=spec.getXML()
XML=objectify.XML(XML)
sections=XML.xpath('section')
for el in sections[1].element:
    print unicode(el.Name)

file=open("out.html","w")
file.write(template.render(
    sections=sections,
    ).encode('utf-8'))
file.close()

