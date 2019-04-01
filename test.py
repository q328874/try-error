#!/usr/bin/env python
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
# next Tutorial: 18
# https://www.youtube.com/watch?v=aDlTl06iBHo

from tkinter import *
from math import *

def calculate(event):
	gleichung = t.get()
	t.delete(0,END)
	try:
			t.insert(0, eval(gleichung))
	except:
			t.insert(0, "Invalid syntax")

top = Tk()
t = Entry(top)
t.grid(row=0,columnspan=3)

b1 = Button(top,text="1")
b1.grid(row=1,column=0)
b2 = Button(top,text="2")
b2.grid(row=1,column=1)
b3 = Button(top,text="3")
b3.grid(row=1,column=2)
b4 = Button(top,text="4")
b4.grid(row=2,column=0)
b5 = Button(top,text="5")
b5.grid(row=2,column=1)
b6 = Button(top,text="6")
b6.grid(row=2,column=2)
b7 = Button(top,text="7")
b7.grid(row=3,column=0)
b8 = Button(top,text="8")
b8.grid(row=3,column=1)
b9 = Button(top,text="9")
b9.grid(row=3,column=2)
b0 = Button(top,text="0")
b0.grid(row=4,column=1)
bp = Button(top,text="+")
bp.grid(row=0,column=3)
bm = Button(top,text="-")
bm.grid(row=1,column=3)
bmu = Button(top,text="*")
bmu.grid(row=2,column=3)
bd = Button(top,text="/")
bd.grid(row=3,column=3)
be = Button(top,text="=")
be.grid(row=4,column=3)
bdel = Button(top,text="DEL")
bdel.grid(row=4,column=2)

b1.bind("<Button-1>", lambda x: t.insert(END, "1"))
b2.bind("<Button-1>", lambda x: t.insert(END, "2"))
b3.bind("<Button-1>", lambda x: t.insert(END, "3"))
b4.bind("<Button-1>", lambda x: t.insert(END, "4"))
b5.bind("<Button-1>", lambda x: t.insert(END, "5"))
b6.bind("<Button-1>", lambda x: t.insert(END, "6"))
b7.bind("<Button-1>", lambda x: t.insert(END, "7"))
b8.bind("<Button-1>", lambda x: t.insert(END, "8"))
b9.bind("<Button-1>", lambda x: t.insert(END, "9"))
b0.bind("<Button-1>", lambda x: t.insert(END, "0"))
bp.bind("<Button-1>", lambda x: t.insert(END, "+"))
bm.bind("<Button-1>", lambda x: t.insert(END, "-"))
bmu.bind("<Button-1>", lambda x: t.insert(END, "*"))
bd.bind("<Button-1>", lambda x: t.insert(END, "/"))
be.bind("<Button-1>", calculate)
bdel.bind("<Button-1>", lambda x: t.delete(0,END,))

top.mainloop()
