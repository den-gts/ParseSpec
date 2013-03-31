# -*-coding: utf-8 -*-
import os
from jinja2 import Environment
from jinja2.loaders import FileSystemLoader
from docSpecParcer import DocSpecification
env=Environment(loader=FileSystemLoader(os.getcwd()))
template=env.get_template("structure.html")
spec=DocSpecification()
file=open("out.html","w")
file.write(template.render(
    sections=spec.section.keys(),
    spec=spec
    ).encode('utf-8'))
file.close()

