#!/usr/bin/env python

import os
import string
import copy
import re
import sys
import codecs
import xlrd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from collections import OrderedDict

# localize_target = set(["ko", "ja", "en"]);
# eventScenePath = "C:\\Work\\SheepFarm\\client\\Resource\\eventScene";

def exportXml(creatureList):
	# output = '<?xml version="1.0" encoding="UTF-8" ?>'

	for creatureInfo in creatureList:
		creatureNode = ET.Element("creature")
		infoNode = ET.SubElement(creatureNode, "stat")

		ET.SubElement(infoNode, "name").text = creatureInfo["name"]
		ET.SubElement(infoNode, "level").text = creatureInfo["level"]
		ET.SubElement(infoNode, "attackType").text = creatureInfo["type"]

		ET.SubElement(infoNode, "horrorProb").text = unicode(creatureInfo["horrorProb"])
		ET.SubElement(infoNode, "horrorMin").text = unicode(creatureInfo["horrorMin"])
		ET.SubElement(infoNode, "horrorMax").text = unicode(creatureInfo["horrorMax"])

		ET.SubElement(infoNode, "physicsProb").text = unicode(creatureInfo["physicsProb"])
		ET.SubElement(infoNode, "physicsMin").text = unicode(creatureInfo["physicsMin"])
		ET.SubElement(infoNode, "physicsMax").text = unicode(creatureInfo["physicsMax"])

		ET.SubElement(infoNode, "mentalProb").text = unicode(creatureInfo["mentalProb"])
		ET.SubElement(infoNode, "mentalMin").text = unicode(creatureInfo["mentalMin"])
		ET.SubElement(infoNode, "mentalMax").text = unicode(creatureInfo["mentalMax"])

		ET.SubElement(infoNode, "feelingMax").text = unicode(creatureInfo["feelingMax"])

		ET.SubElement(infoNode, "feelingDownProb").text = unicode(creatureInfo["feelingDownProb"])
		ET.SubElement(infoNode, "feelingDownValue").text = unicode(creatureInfo["feelingDownValue"])

		ET.SubElement(infoNode, "energyGenGood").text = unicode(creatureInfo["energyGenGood"])
		ET.SubElement(infoNode, "energyGenNorm").text = unicode(creatureInfo["energyGenNorm"])
		ET.SubElement(infoNode, "energyGenBad").text = unicode(creatureInfo["energyGenBad"])

		ET.SubElement(infoNode, "preferSkiilsGood").text = unicode(creatureInfo["preferSkiilsGood"])
		ET.SubElement(infoNode, "preferValuesGood").text = unicode(creatureInfo["preferValuesGood"])
		ET.SubElement(infoNode, "rejectSkiilsGood").text = unicode(creatureInfo["rejectSkiilsGood"])
		ET.SubElement(infoNode, "rejectValuesGood").text = unicode(creatureInfo["rejectValuesGood"])

		ET.SubElement(infoNode, "preferSkiilsNorm").text = unicode(creatureInfo["preferSkiilsNorm"])
		ET.SubElement(infoNode, "preferValuesNorm").text = unicode(creatureInfo["preferValuesNorm"])
		ET.SubElement(infoNode, "rejectSkiilsNorm").text = unicode(creatureInfo["rejectSkiilsNorm"])
		ET.SubElement(infoNode, "rejectValuesNorm").text = unicode(creatureInfo["rejectValuesNorm"])

		ET.SubElement(infoNode, "preferSkiilsBad").text = unicode(creatureInfo["preferSkiilsBad"])
		ET.SubElement(infoNode, "preferValuesBad").text = unicode(creatureInfo["preferValuesBad"])
		ET.SubElement(infoNode, "rejectSkiilsBad").text = unicode(creatureInfo["rejectSkiilsBad"])
		ET.SubElement(infoNode, "rejectValuesBad").text = unicode(creatureInfo["rejectValuesBad"])

		# ET.ElementTree(creatureNode).write('output.xml', encoding="utf-8", xml_declaration=True)

		print(minidom.parseString(ET.tostring(creatureNode,  encoding="utf-8")).toprettyxml())

		fp = codecs.open(creatureInfo["name"] + '.xml', 'w', "utf-8") 
		fp.write(minidom.parseString(ET.tostring(creatureNode,  encoding="utf-8")).toprettyxml())
		fp.close()
	pass

def importTable(book):
	sheet = book.sheet_by_name("Sheet1");

	creatureList = []

	for rowIndex in range(2, sheet.nrows):
		creatureInfo = {}

		creatureInfo['name'] = sheet.cell(rowIndex, 0).value
		creatureInfo['level'] = sheet.cell(rowIndex, 1).value
		creatureInfo['type'] = sheet.cell(rowIndex, 2).value

		if creatureInfo['level'] == '':
			continue


		creatureInfo['horrorProb'] = sheet.cell(rowIndex, 3).value
		creatureInfo['horrorMin'] = sheet.cell(rowIndex, 4).value
		creatureInfo['horrorMax'] = sheet.cell(rowIndex, 5).value

		creatureInfo['physicsProb'] = sheet.cell(rowIndex, 8).value
		creatureInfo['physicsMin'] = sheet.cell(rowIndex, 9).value
		creatureInfo['physicsMax'] = sheet.cell(rowIndex, 10).value

		creatureInfo['mentalProb'] = sheet.cell(rowIndex, 11).value
		creatureInfo['mentalMin'] = sheet.cell(rowIndex, 12).value
		creatureInfo['mentalMax'] = sheet.cell(rowIndex, 13).value

		creatureInfo['feelingMax'] = sheet.cell(rowIndex, 14).value

		creatureInfo['feelingDownProb'] = sheet.cell(rowIndex, 16).value
		creatureInfo['feelingDownValue'] = sheet.cell(rowIndex, 17).value

		creatureInfo['energyGenGood'] = sheet.cell(rowIndex, 19).value
		creatureInfo['energyGenNorm'] = sheet.cell(rowIndex, 20).value
		creatureInfo['energyGenBad'] = sheet.cell(rowIndex, 21).value

		creatureInfo['preferSkiilsGood'] = unicode(sheet.cell(rowIndex, 25).value).replace("\n",",")
		creatureInfo['preferValuesGood'] = unicode(sheet.cell(rowIndex, 26).value).replace("\n",",")
		creatureInfo['rejectSkiilsGood'] = unicode(sheet.cell(rowIndex, 27).value).replace("\n",",")
		creatureInfo['rejectValuesGood'] = unicode(sheet.cell(rowIndex, 28).value).replace("\n",",")

		creatureInfo['preferSkiilsNorm'] = unicode(sheet.cell(rowIndex, 30).value).replace("\n",",")
		creatureInfo['preferValuesNorm'] = unicode(sheet.cell(rowIndex, 31).value).replace("\n",",")
		creatureInfo['rejectSkiilsNorm'] = unicode(sheet.cell(rowIndex, 32).value).replace("\n",",")
		creatureInfo['rejectValuesNorm'] = unicode(sheet.cell(rowIndex, 33).value).replace("\n",",")

		creatureInfo['preferSkiilsBad'] = unicode(sheet.cell(rowIndex, 35).value).replace("\n",",")
		creatureInfo['preferValuesBad'] = unicode(sheet.cell(rowIndex, 36).value).replace("\n",",")
		creatureInfo['rejectSkiilsBad'] = unicode(sheet.cell(rowIndex, 37).value).replace("\n",",")
		creatureInfo['rejectValuesBad'] = unicode(sheet.cell(rowIndex, 38).value).replace("\n",",")

		creatureList.append(creatureInfo)

	return creatureList

	# scriptFile = codecs.open(os.path.dirname(__file__) + os.sep + "output.lua", 'w', encoding='utf8');
	# scriptFile.write(template);
	# scriptFile.close();




if __name__ == '__main__':

	try:
		# book = xlrd.open_workbook(sys.argv[1]);
		book = xlrd.open_workbook(os.path.dirname(__file__) + os.sep + "creature.xlsx")

		creatureList = importTable(book)

		exportXml(creatureList)
		
		print("success");
		os.system("pause");
		
	except Exception, error:
		import traceback
		print(traceback.format_exc())
		os.system("pause");
