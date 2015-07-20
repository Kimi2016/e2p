# -*- $Id -*-  
# -*- coding: utf-8 -*-
# author: dengzhifeng
# dep: xlrd

"""
	通用导表
	说明: 
		!!! 例子是最好的说明	
		请查询 xml/devxls_test.xls
		然后 python devxls.py xml/devxls_test.xls > result.lua
		对照 result.lua 的结果
		
	详细说明如下：

	字段生成规则:
	1. 注释忽略
		以 '//' 开头的字段将忽略 
		ex: //Desc

	2. 多层存储
		以 . 来制定存储的位置
		ex: A.B.C 将生成
			A = {
				B = {
					C = xxx	
				}	
			} 

	3. Id
		第一个列为Id 字段, 用于作为记录的Key
		形如[ 1001 ] = { xxx = xxx, ...}
		类型必须为Uid 或者 AutoId
		Uid 为手动指定Id, 可以为Number 或 String
		AutoId 为自动生成Id


	类型规则:
	1. 类型关键字 以 '|' 切分
		ex: List|String|Defalut 

	2. List 的字段将生成 列表
		ex: List|Int
			1|2|3
		结果 为 {1,2,3}

	3. VarArgs 的字段为多列模式
		ex: VarArgs|String
			AAAAAAAAAAAA		BBBBB
		结果为 {'AAAAAAAAAA' , 'BBBBB' }

		VarArgs 只能在最后一字段使用
		VarArgs 和 List 不能并用

	4. Defalut 制定使用默认值
		ex: Int|Default
		
		结果 为 0

		不使用Default的字段默认皆为nil

	表单生成规则:
	1. main 为主表单 main 的 id 字段将为其他表单的 索引
		字表单嵌套在Content字段里
		ex:
			main 
			1001 xxx xxx
			1002 xxx xxx

		1001 1002

		将生成
		{
			[1001] = {
				Content = {
				}	
			}
			[1002] 
		}
	
	2. desc 表单 用于写描述信息, 生成数据时忽略

	3. 表单中前面的行如果首列字段为空,将被忽略

	扩展类型规则
	指定 -i hookfile 后可以使用扩展类型
	hookfile 是一 python 文件
	必须具有 handler_dict
	如 handler_dict = {
		"Ext1" = Class1, 
	}
	Class1必须具有 convert 或 default(可选) 方法
	返回值为 str, 将作为 dump 时输出的文本
	详细参考 tools/dtype_sample.py
"""
import xlrd
import math
import sys
import traceback
from sys import exit
import os.path
from types import *
import dtype_flexattr
from collections import OrderedDict

KEY_ROW = 1
TYPE_ROW = 2
DATA_ROW = 3

## 关键字
INVALID = "Invalid"
NAME = "Name"
TYPE = "Type"
LIST = "List"
COMPOSITE = 'Composite'
REF = "RefType"
RAWDATA = "RawData"
RAW_TEXT = "RawText"	# 原始数据
COMMENT = "Comment"
VARARGS = "VarArgs"
DEFAULT = "Default"
CONTENT = "Content"
UID = "Uid"
AUTO_ID = "AutoId"
NEED_OUTPUT = "NeedOutput"
OUTPUT_ENCODE = "utf-8"
LUAC_BIN_PATH = "luac"
PLATFORM = "Platform"
PLATFORM_IOS = "_ios"
PLATFORM_AND = "_android"
MERGE_KEY = 'MergeKey'
SKIP_ROW = 'SkipRow'

write = sys.stdout.write
err_write = sys.stderr.write
datemode = None
pre_dump_table = None
post_custom_text = None
handler_dict = {
	'FlexAttr':  dtype_flexattr.FlexAttr,
}
output_lang = None

class ExtentType:
	def __init__(self, value, comment = None):
		self.value = value
		self.comment = comment
	def __str__(self):
		if type(self.value) == UnicodeType:
			return self.value.encode(OUTPUT_ENCODE)
		return str(self.value)


# 读取xls里的类型数据
def xls_format(cell, i, j):
	value = cell.value
	ctype = cell.ctype
	# print(value, ctype, type(value))

	if ctype == xlrd.XL_CELL_EMPTY: #0
		return None 
	elif ctype == xlrd.XL_CELL_TEXT: #1
		return value
	elif ctype == xlrd.XL_CELL_NUMBER: #2
		return value
	elif ctype == xlrd.XL_CELL_DATE: #3
		if datemode is None:
			exit("Error: datemode is None")
		datetuple =	xlrd.xldate_as_tuple(value, datemode)
		# time only	no date	component
		if datetuple[0]	== 0 and datetuple[1] == 0 and datetuple[2] == 0: 
			value =	"%02d:%02d:%02d" % datetuple[3:]
		# date only, no	time
		elif datetuple[3] == 0 and datetuple[4]	== 0 and datetuple[5] == 0:
			value =	"%04d:%02d:%02d" % datetuple[:3]
		else: #	full date
			value =	"%04d:%02d:%02d:%02d:%02d:%02d"	% datetuple

		return value
	elif ctype == xlrd.XL_CELL_BOOLEAN: #4
		# 现在不使用
		# TODO
		exit("Error: invalid xls data type of BOOLEAN")
	else:
		exit("Error: invalid xls data type of No. %d  at row %d, col %d" % (ctype, i, j))

def adjust_type(value):
	if isinstance(value, float) and round(value) == value:
		return int(value)
	return value

# 转化为Python 类型值
def parse_type(value, vtype):
	if handler_dict: #优先处理自定义类型
		# 存在钩子
		if vtype in handler_dict:
			try:
				handler = handler_dict[vtype]()
				value = handler.convert(value)
			except:
				err_write("Warnning: handler_dict[%s].convert(%s) failed\n" % (vtype, value))
				traceback.print_exc() 
				raise
				value = None

			if value is None:
				return None
			else:
				return ExtentType(value)

	if vtype == RAWDATA:
		if value is None:
			return None
		else:
			value = adjust_type(value)
			return ExtentType(value)
	elif vtype == RAW_TEXT: # 没有检查类型了，让luac去检查吧
		if value is None:
			return None
		else:
			value = adjust_type(value)
			value = str(value)
			return ExtentType(value)
	elif vtype == COMMENT:
		if value is None:
			return None
		else:
			return ExtentType(value, True)
	elif vtype == "Int":
		# 默认值
		if value == "":
			value = 0
		try:
			value = int(value)
		except:
			exit("Error: %s can not convert to Int" % value)
	elif vtype == "Number" or vtype == "Float":
		if value == "":
			value = 0.0
		try:
			value = float(value)
		except:
			exit("Error: %s can not convert to Number" % value)
	elif vtype == "Bool":
		try:
			value = bool(value)
		except:
			exit("Error: %s can not convert ot Bool" % value)
	elif vtype == "DateTable":
		parm = value.replace("-",":")
		parm = parm.split(":")
		return {
				"year": int(parm[0]),
				"month": int(parm[1]),
				"day": int(parm[2]),
				}
	elif vtype == "TimeTable":
		parm = value.split(":")
		return {
				"hour": int(parm[0]),
				"min": int(parm[1]),
				"sec": int(parm[2]),
				}
	elif vtype == "DateTime":
		parm = value.replace("-",":")
		parm = parm.replace(" ",":")
		parm = parm.split(":")
		if len(parm) <=3:
			err_write("Warnning: DateTime append hour:min:sec auto value=%s\n" % value)
			parm.append(0)
			parm.append(0)
			parm.append(0)

		return {
				"year": int(parm[0]),
				"month": int(parm[1]),
				"day": int(parm[2]),
				"hour": int(parm[3]),
				"min": int(parm[4]),
				"sec": int(parm[5]),
				}

	elif vtype == "Expr":
		if value is None:
			return None
		else:
			try:
				value = str(int(value))
			except:
				value = "function (Args) return %s end"	% value
			return ExtentType(value)
	elif vtype == "Func":
		global  output_lang
		if output_lang == LANG_LUA:
			value = "function (Args, ...) return %s end" % (value or "nil")
		elif output_lang == LANG_PYTHON:
			value = "lambda : %s" % (value or "None")
		else:
			err_write('only support output .py|.lua')
		return ExtentType(value)	
	elif vtype == "TrackFunc":
		if output_lang == LANG_LUA:
			value = "function (Args, t) return %s end" % (value or "nil")
		elif output_lang == LANG_PYTHON:
			value = "lambda t : %s" % (value or "None")
		else:
			err_write('only support output .py|.lua')
		return ExtentType(value)
	elif vtype == "Run":
		if value is None:
			if output_lang == LANG_LUA:
				value = "function(Args) return Args end"
			elif output_lang == LANG_PYTHON:
				value = "lambda val: val"
			else:
				err_write('only support output .py|.lua')
		else:
			old_value = value
			value = "function (Args) %s; return Args end" % value
			tmp = "echo 'test = %s' | %s -p -" % (value, LUAC_BIN_PATH)
			if os.system(tmp):
				exit("Error: %s lua语法错误" % old_value)
		return ExtentType(value)
	elif vtype == "SeqFunc":
		if value is None:
			value = "function() return {} end"
		else:
			old_value = value
			value = "function() return UTIL.parseSeq(%s) end" % value
			tmp = "echo 'test = %s' | %s -p -" % (value, LUAC_BIN_PATH)
			if os.system(tmp):
				exit("Error: %s lua语法错误" % old_value)
		return ExtentType(value)
	elif vtype.startswith(COMPOSITE):
		template = eval(vtype[len(COMPOSITE):])
		value = convert(value, template)
		return  value

	return value

##########################COMPOSITE type begin#############################
def _convert_template(data, template = []):
	if template is None:
		return data
	elif isinstance(template, list):
		return _convert_template_list(data,template)
	elif isinstance(template, dict):
		return _convert_template_dict(data,template)
	else:
		assert False,'Invalid tempalte type'%type(template)

def _convert_template_list(data, template):
	new_temp = template[0] if len(template)>0 else None
	ret = []
	for k in data:
		ret.append(_convert_template(k, new_temp))
	return ret

def _convert_template_dict(data, template):
	ret = {}
	for i in xrange(len(data)):
		info = template[i]
		key = info[0] if isinstance(info, tuple) else info
		new_temp = info[1] if isinstance(info, tuple) and len(info)>1 else None
		ret[key] = _convert_template(data[i], new_temp)
	return ret

def convert(data_str, template = []):
	import json
	json_str = '['+data_str.replace('(','[').replace(')',']')+']'
	data = json.loads(json_str)

	return _convert_template(data, template)
##########################COMPOSITE type end#############################

def load_hookfile(filename):
	import imp
	global pre_dump_table 
	global post_custom_text 
	global handler_dict
	try:
		module = imp.load_source('d', filename)
		if hasattr(module, 'handler_dict'):
			for k, v in module.handler_dict.iteritems():
				handler_dict[k] = v
		if hasattr(module, 'pre_dump_table'):
			pre_dump_table = module.pre_dump_table
		if hasattr(module, 'post_custom_text'):
			post_custom_text = module.post_custom_text
	except:
		err_write("Warnning: load hookfile %s failed. use default type.\n" % filename)

# 默认值 
def default(type_info):
	vtype = type_info[TYPE]
	if handler_dict: #优先处理自定义类型
		# 存在钩子
		if vtype in handler_dict:
			try:
				handler = handler_dict[vtype]()
				value = handler.default()
			except:
				err_write("Warnning: handler_dict[%s].default() failed\n" % (vtype))
				value = None

			if value is None:
				return None
			else:
				return ExtentType(value)

	if type_info[LIST]:
		return []
	elif vtype == "Int":
		return 0
	elif vtype == "Number":
		return 0
	elif vtype == "Bool":
		return False
	elif vtype == "DateTable":
		return None
	elif vtype == "TimeTable":
		return { "hour":0, "min":0, "sec":0 }
	elif vtype == "DateTime":
		return None
	else:
		return None 

def try_convert_int(str):
	try:
		return int(str)
	except:
		return str

def push_value(value, path, record_table, varargs=False):
	path_list = path.split(".")
	#检测是否需要合并
	merge_lst = path_list[-1].split('|')
	merge_prop = None
	if len(merge_lst) > 1:
		assert len(merge_lst) == 2,'只支持简单数组的属性合并'
		path_list[-1] = merge_lst[0]	
		merge_prop = merge_lst[1]
		assert type(value) == ListType,'属性合并项必须是数组'

	cur_table = record_table
	for path in path_list[:-1]:
		path = try_convert_int(path)

		if not path in cur_table:
			cur_table[path] = {}
		cur_table = cur_table[path]
	
	key = path_list[-1]
	key = try_convert_int(key)

	if varargs:
		if not key in cur_table:
			cur_table[key] = [ value ]
		else:
			cur_table[key].append(value)
	elif merge_prop is not None:
		if not key in cur_table:
			cur_table[key] = OrderedDict()
		if merge_prop == MERGE_KEY:
			for merge_key in value:
				assert merge_key not in cur_table[key], ('属性合并时key值重复',merge_key)
				cur_table[key][merge_key] = {}
		else:
			assert len(value) == len(cur_table[key]),('merge_key和待合并属性数量不吻合',key,merge_prop)
			cur_pos = 0
			for merge_key,merge_dict in cur_table[key].iteritems():
				merge_dict[merge_prop] = value[cur_pos]
				cur_pos += 1
	else:
		try:
			cur_table[key] = value
		except:
			exit("Error: %s 的 %s 已经被赋值 不能展开为子表" % (path, cur_table))



def parse_value(value, type_info):
	# 列表类型 并且 是 文本
	if type_info[LIST]:
		if type(value) == UnicodeType or type(value) == StringType:
			value = value.split("|")
			for i in xrange(len(value)):
				value[i] = parse_type(value[i], type_info[TYPE])
		else:
			value = [ parse_type(value, type_info[TYPE]) ]
	elif type_info[REF]:
		pass
		# load(type_info[TYPE])
		# value = check_ref(value, type_info[TYPE])
	else:
		value = parse_type(value, type_info[TYPE])

	return value

	

def parse_sheet(sheet, keyrow, typerow, datrow_begin, sheet_name, skip_row_num = 0):
	ncol = sheet.ncols
	nrow = sheet.nrows

	#扫描第一行,首列不为空的, 作为表头
	skip_row = 0
	for i in xrange(nrow):
		if xls_format(sheet.row(i)[0], i, 0) is not None:
			skip_row = i
			break
	skip_row += skip_row_num  #略过表头

	key_row = sheet.row(keyrow + skip_row)
	type_row = sheet.row(typerow + skip_row)


	sheet_table = {}
	
	# 处理字段名 和 类型数据
	type_info_list = {} 
	for i in xrange(ncol):
		# 初始化类型信息表
		type_info_list[i] = {}

		key = xls_format(key_row[i], i, 0)
		vtype = xls_format(type_row[i], i, 0)

		# 默认关键字值
		type_info_list[i][LIST] = False
		type_info_list[i][COMPOSITE] = False
		type_info_list[i][VARARGS] = False
		type_info_list[i][DEFAULT] = False
		type_info_list[i][INVALID] = False
		type_info_list[i][REF] = False

		if key is not None: #处理 key 是整数的情况
			if type(key) == FloatType:
				if math.floor(key) == key:
					key = int(key)

		# vtype为None, 或者非第一列的key为None, 或者key以//开头

		if type(key) != StringType and type(key) != UnicodeType:
			key = str(key)

		if vtype is None or (i != 0 and key is None) or (key is not None and key.startswith("//")):
			type_info_list[i][INVALID] = True
			continue

		type_info_list[i][NAME] = key

		type_list = vtype.split('|')
		for ii in xrange(len(type_list)):
			type_key = type_list[ii]
			if type_key == LIST:
				type_info_list[i][LIST] = True
			elif type_key == VARARGS:
				type_info_list[i][VARARGS] = True
			elif type_key == DEFAULT:
				type_info_list[i][DEFAULT] = True
			elif type_key == REF:
				type_info_list[i][REF] = True
			else:
				type_info_list[i][TYPE] = type_key
			#err_write("%s %i %s\n" % (sheet.name, i, type_info_list[i][TYPE]))

		if type_info_list[i][LIST] and type_info_list[i][VARARGS]:
			exit("Error: field %s 同时是List 和 VarArgs" % vtype)
		if type(key) == StringType and key.find('[')>=0 or key.find(']')>=0:
			if type_info_list[i][LIST] is False:
				exit("Error: field %s 属性合并项必须是List" % vtype)
			if key.find('[')<0 or key.find(']')<0:
				exit("Error: field %s 格式错误" % key)


	if type_info_list[0][TYPE] == UID:
		auto_id = False
	elif type_info_list[0][TYPE] == AUTO_ID:
		auto_id = True
	else:
		err_write("Error sheet %s" % sheet_name)
		exit("Error: 第一列非空的列类型必须为Uid 或 AutoId")

	#确认是否需要导表
	if len(type_info_list) > 1 and TYPE in type_info_list[1] and type_info_list[1][TYPE] == NEED_OUTPUT:
		output_all = False
	else:
		output_all = True

	# 确认平台属性
	if len(type_info_list) > 1 and TYPE in type_info_list[1] and type_info_list[1][TYPE] == PLATFORM:
		platform = True
	else:
		platform = False
	
	last_id = 0
	record_id = 0
	id_dict = {}
	
	for i in xrange(datrow_begin + skip_row, sheet.nrows) :
		row_data = sheet.row(i)
		#判断本行是否需要导出
		need_output = xls_format(row_data[1],0,0)
		if not output_all and not bool(need_output):
			continue

		if auto_id:
			last_id = last_id + 1
			record_id = last_id 
		else:
			record_id = xls_format(row_data[0], i, 0)
			try: # 尝试转换为整数Id
				record_id = int(record_id)
			except:
				pass
			if not platform and (record_id in id_dict):
				err_write("%s\n" % row_data)
				exit("Error: reocrd_id %s duplicate, i=%d,sheet=%s" % (record_id , i, sheet.name))
			id_dict[record_id] = True
		# 按照平台区分表
		sheet_row = {}
		if platform:
			platform_type = xls_format(row_data[1], i, 0)
			if platform_type == None:
				sheet_table[record_id] = sheet_row
			else:
				if not record_id in sheet_table:
					sheet_table[record_id] = {}
				sheet_table[record_id][platform_type] = sheet_row
		else:
			sheet_table[record_id] = sheet_row


		varargs = False
		varargs_pos = 0
		for ii in xrange(1, ncol) :
			value = xls_format(row_data[ii], i, ii)

			if varargs:
				type_info = type_info_list[varargs_pos]
			else:
				type_info = type_info_list[ii]

			# 如果设置了多列模式, 下一列开始使用
			if not varargs and type_info[VARARGS]:
				varargs = True
				varargs_pos = ii

			
			if type_info[INVALID]:
				continue

			if value is None and type_info[DEFAULT]:
				# 赋予默认值
				# 最小惊讶原则: 显式控制默认值
				value = default(type_info)
			elif not value is None:
				value = parse_value(value, type_info)

			if value is None:
				continue

			# 多列数据格式, 只能出现在最后一列 
			if varargs:
				push_value(value, type_info[NAME], sheet_row, True)
				continue

			push_value(value, type_info[NAME], sheet_row)


	return sheet_table

def parse_table(xls_fileobj) :

	main_table = {}
	for sheet in xls_fileobj.sheets():
		try:
			name = int(sheet.name)
		except:
			name = sheet.name

		if name == "main":
			main_table = parse_sheet(sheet, KEY_ROW, TYPE_ROW, DATA_ROW, name)
		elif name == "desc":
			pass
		else:
			if not name in main_table:
				#err_write("Warning: %s sheet exist but not in main sheet Uid\n" % name)
				continue
			sheet_table = parse_sheet(sheet, KEY_ROW, TYPE_ROW, DATA_ROW, name,main_table[name][SKIP_ROW])
			main_table[name] = sheet_table
	
	return main_table

def convert2Dict(xls_file, xls_sheet= None,skip_row_num=0):
	xls_fileobj = xlrd.open_workbook(xls_file, logfile=sys.stderr)
	global datemode
	datemode = xls_fileobj.datemode

	if xls_sheet is None:#导出整表
		return parse_table(xls_fileobj)
	else:#导出指定sheet
		for sheet in xls_fileobj.sheets():
			try:
				name = int(sheet.name)
			except:
				name = sheet.name

			if name != xls_sheet:
				continue
			return parse_sheet(sheet,KEY_ROW,TYPE_ROW,DATA_ROW,name,skip_row_num)
		else:
			exit('No sheet named %s'%sheet_name)

base_type_dict = { 
		IntType: True,
		FloatType: True,
		BooleanType: True,
		StringType: True,
		UnicodeType: True,
		NoneType: True,
		ListType: True,
		InstanceType: True,
}

############################################dump python begin###################################################
def base_dump_python(value):
	if type(value) == IntType:
		write("%d" % value)
	elif type(value) == FloatType:
		write("%f" % value)
	elif type(value) == BooleanType:
		if value:
			write("True")
		else:
			write("False")
	elif type(value) == StringType:
		write("\"%s\"" % value)
	elif type(value) == UnicodeType:
		write("\"%s\"" % value.encode(OUTPUT_ENCODE))

	elif type(value) == NoneType:
		write("None")
	elif type(value) == ListType:
		write("( ")
		for x in value:
			base_dump_python(x)
			write(", ")
		write(")")
	elif type(value) == DictType:
		# 不向下展开的Dict,只用于List中的最后一层数据
		# 如 List|DateTable 等
		write("{ ")
		for k, v in value.iteritems():
			if type(k) == StringType:
				write("\"%s\"  : " % k)
			elif type(value) == UnicodeType:
				write("\"%s\"  :  " % k.encode(OUTPUT_ENCODE))
			else:
				write("%s : " % k)
			base_dump_python(v)
			write(", ")
		write("}")
	elif type(value) == InstanceType:
		write(str(value))

def dump_value_python(data, level=1):
	ret_dict = {}
	if isinstance(data,OrderedDict):
		for k,v in data.iteritems():
			ret_dict[k] = v
		data = ret_dict
	type_value = type(data)

	if type_value in base_type_dict :
		base_dump_python(data)
	elif type_value == DictType:
		write("{\n")
		##
		# 为了每次读表diff好看。。。sort一下
		sortitems = data.items()
		sortitems.sort()
		for k, v in sortitems:
			write("\t"*level)

			if isinstance(v, ExtentType) and v.comment == True:
				write("'''\n")
				write("\t"*level)
				write(k + ":\n")
				write("\t"*level)
			elif type(k) == IntType:
				write("%d : " % k)
			elif k.find('@') >= 0:
				write("\"%s\" : " % k.encode('ascii'))
			else:
				tag = '\"' if type(k) == StringType or type(k) == UnicodeType else ''
				try:
					write("%s%s%s : " % (tag, k.encode('ascii'), tag))
				except:
					write("%s%s%s : " % (tag,k.encode(OUTPUT_ENCODE),tag))
			dump_value_python(v, level + 1)
			if isinstance(v, ExtentType) and v.comment == True:
				write("\n")
				write("\t"*level)
				write("'''\n")
			else:
				write(" ,\n")
		##
		write("\t"*(level - 1))
		write("}")
	else:
		exit("Error: Unkonwn Type %s in dump " % type_value)
############################################dump python end###################################################


############################################dump lua begin####################################################
def base_dump_lua(value):
	if type(value) == IntType:
		write("%d" % value)
	elif type(value) == FloatType:
		write("%f" % value)
	elif type(value) == BooleanType:
		if value:
			write("true")
		else:
			write("false")
	elif type(value) == StringType:
		if value.find("]]") >= 0:
			exit("Error: %s cantain ']]'" % value)
		write("[[%s]]" % value)
	elif type(value) == UnicodeType:
		if value.find("]]") >= 0:
			exit("Error: %s cantain ']]'" % value)
		write("[[%s]]" % value.encode(OUTPUT_ENCODE))

	elif type(value) == NoneType:
		write("nil")
	elif type(value) == ListType:
		write("{ ")
		for x in value:
			base_dump_lua(x)
			write(", ")
		write("}")
	elif type(value) == DictType:
		# 不向下展开的Dict,只用于List中的最后一层数据
		# 如 List|DateTable 等
		write("{ ")
		for k, v in value.iteritems():
			if type(k) == IntType:
				write("[ %d ] = " % k)
			else:
				write("%s = " % k)
			base_dump_lua(v)
			write(", ")
		write("}")
	elif type(value) == InstanceType:
		write(str(value))

def dump_value_lua(data, level=1):
	type_value = type(data)

	if type_value in base_type_dict:
		base_dump_lua(data)
	elif type_value == DictType:
		write("{\n")
		##
		# 为了每次读表diff好看。。。sort一下
		sortitems = data.items()
		sortitems.sort()
		for k, v in sortitems:
			for i in xrange(level):
				write("\t")

			if isinstance(v, ExtentType) and v.comment == True:
				write(" --[[\n")
				for i in xrange(level):
					write("\t")
				write(k + ":\n")
				for i in xrange(level):
					write("\t")
			elif type(k) == IntType:
				write("[ %d ] = " % k)
			elif k.find('@') >= 0:
				write("['%s'] = " % k.encode('ascii'))
			else:
				try:
					write("%s = " % k.encode('ascii'))
				except:
					write("['%s'] = " % k.encode(OUTPUT_ENCODE))
			dump_value_lua(v, level + 1)
			if isinstance(v, ExtentType) and v.comment == True:
				write("\n")
				for i in xrange(level):
					write("\t")
				write("]]\n")
			else:
				write(" ,\n")
		##
		for i in xrange(level - 1):
			write("\t")
		write("}")
	else:
		exit("Error: Unkonwn Type %s in dump " % type_value)
############################################dump lua end#####################################################

def dump_value(data, lang):
	if lang == LANG_PYTHON:
		return dump_value_python(data)
	elif lang == LANG_LUA:
		return dump_value_lua(data)
	else:
		err_write('Only support output to lua|python')


def MakeQuickLink(Name, lang):
	if lang == LANG_PYTHON:
		return '''
local __%(Name)s__ = data.%(Name)s
function Get%(Name)s() return __%(Name)s__ end
''' % { 'Name' : Name }
	elif lang == LANG_LUA:
		return '''
__%(Name)s__ = data[%(Name)s]
def Get%(Name)s(): return __%(Name)s__
''' % { 'Name' : Name }

def usage():
	exit('''usage: filename [-i hookfile] [-h]
			hookfile: 扩展类型钩子文件
		''')

def merge(src, tar):
	for k, v in tar.items():
		if k in src:
			for name, content in v.items():
				if type(v[name]) == ListType:
					src[k][name].extrend(v[name])
				else:
					for idx, val in v[name].items():
						if idx in src[k][name]:
							exit("Error: %s %s conflict" % (k, str(idx)))
						else:
							src[k][name][idx] = val
		else:
			src[k] = v

#输出语言选择
LANG_PYTHON = 0x1111
LANG_LUA = 0x2222

def convert2File(filename,output_filename,xls_sheet=None,skip_row_num=0,hookfile=None,update_file_list=None):
	if not os.path.isfile(filename) and not os.path.isdir(filename):
		exit("Error: %s is not a valid filename or pathname" % filename)

	#输出文件
	if output_filename.find('.') <0:
		exit('need extension in .py|.lua')

	ext = output_filename.split('.')[-1]
	global  output_lang
	if ext == 'py':
		output_lang = LANG_PYTHON
		output_comment = '#'
	elif ext == 'lua':
		output_lang = LANG_LUA
		output_comment = '--'
	else:
		output_comment = None
		exit('invalid extension .%s, must in .py|.lua'%ext)

	output_file = open(output_filename,'w')
	global write
	write = output_file.write

	# 参数处理
	if hookfile is not None:
		if not os.path.isfile(hookfile):
			exit("Error: %s is not a valid filename" % hookfile)
		# 载入类型扩展钩子
		load_hookfile(hookfile)

	if os.path.isfile(filename):
		data_table = convert2Dict(filename,xls_sheet,skip_row_num)
	else:
		data_table = {}
		for subfilename in os.listdir(filename):
			subfilepath = filename + subfilename
			if os.path.isfile(subfilepath):
				merge(data_table, convert2Dict(subfilepath))
	#write("--autogen-begin\n")
	if output_lang == LANG_PYTHON:
		write('#-*-coding:utf-8 -*-\n')
		write("\ndata =\\\n")
	elif output_lang == LANG_LUA:
		write("\nlocal data = \n")

	# 尝试调用扩展钩子模块里的'pre_dump_table'，将数据表转化为dump用的表。
	if pre_dump_table:
		data_table = pre_dump_table(data_table)

	dump_value(data_table,output_lang)

	if output_lang == LANG_PYTHON:
		write("\ndef GetTable(): return data\n")
		write("\ndef GetContent(SheetName): return data[SheetName]\n")

		if update_file_list and len(update_file_list) > 0:
			write("\ndef __update__():\n")
			for filepath in update_file_list:
				write("\tUpdate('" + filepath +"')\n")
	if output_lang == LANG_LUA:
		write("\nfunction GetTable() return data end\n")
		write("\nfunction GetContent(SheetName) return data[SheetName] end\n")

		if update_file_list and len(update_file_list) > 0:
			write("\nfunction __update__()\n")
			for filepath in update_file_list:
				write("\tUpdate('" + filepath +"')\n")
			write("end\n")

	# 尝试生成QuickLink函数
	for Name, Data in data_table.iteritems():
		if 'QuickLink' in Data:
			write(MakeQuickLink(Name,output_lang))

	#write("--autogen-end\n")

	if post_custom_text:
		write("\n%spost_custom_text-begin\n"%output_comment)
		write(post_custom_text())
		write("\n%spost_custom_text-end\n"%output_comment)

	output_file.close()

def main():
	if len(sys.argv) < 3:
		usage()

	#输入文件
	filename = sys.argv[1]
	#输出文件
	output_filename = sys.argv[2]
	# 参数处理
	update_file_list = None
	hookfile = None
	for i in xrange(len(sys.argv)):
		arg = sys.argv[i]
		if arg == '-i':
			hookfile = sys.argv[i + 1]
			if not os.path.isfile(hookfile):
				exit("Error: %s is not a valid filename" % hookfile)
			# 载入类型扩展钩子
			load_hookfile(hookfile)

		if arg == "-u":
			update_file_list = sys.argv[i+1:]

		if arg == '-h':
			usage()

	convert2File(filename,output_filename,hookfile,update_file_list)

if __name__=="__main__":
	# main()
	# print convert2Dict('autogen_config.xls','autogen_config')
	convert2File('devxls_test.xls','result2.py',)
	# convert2File('autogen_config.xls','result.py','autogen_config')
	# convert2Dict('autogen_config.xls','autogen_config')

