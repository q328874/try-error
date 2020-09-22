#!/usr/bin/env python3
# -*- coding: utf-8 -*-
#
#  test.py
#
#  Copyright 2019 Isildur <>
#
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#
import locale
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import csv
locale.setlocale(locale.LC_ALL, 'de_DE.utf8')
#r = requests.get('https://example.org')
#soup = BeautifulSoup(r.text, 'html.parser')
#counter = soup.find(id="counter").text.strip())
counter = 27
d = datetime.now()
datum = d.strftime("%x")
zeit = d.strftime("%X")
with open ('test.csv', 'a') as newFile:
    newFileWriter=csv.writer(newFile)
    newFileWriter.writerow([datum, zeit, counter])
newFile.close()
