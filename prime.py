#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  prime.py
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
#
import time
t0 = time.time()
n = 200000
sieve = [True] * (n+1)
for i in range(2, n+1):
	if sieve[i]:
		for mult in range(i + i, n + 1, i):
			sieve[mult] = False
			s = 0
for i in range(2,n+1):
	if sieve[i]:
		s+=i
print(s)
t1 = time.time()
total = t1-t0
print(total)
