# -*- coding: gbk -*-
import re
import sys
from types import *

def str2floatOrInt(s):
	try:
		f = float(s)
		if f % 1.0 == 0:
			return str(int(float(s)))
	except:
		try:
			return str(float(s))
		except:
			return None

class FlexAttr:
	def convert(self, input):
		t = 'default'

		if isinstance(input, UnicodeType):
			input = input.encode('gbk')
		else:
			input = str(input)

		input = input.strip()
		if input[0] == '"' or input[0] == "'":
			input = input.strip(input[0])
			t = 'string'
		elif input[:2]=='[[':
			input = input.strip()
			t = 'string'
		elif re.search(r'^function\s*\(', input):
			t = 'function'

		if t == 'string':
			return str(input)

		if t == 'function':
			return input

		value = str2floatOrInt(input)
		return value and value or str(input)

