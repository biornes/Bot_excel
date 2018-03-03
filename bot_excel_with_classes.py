import telebot
import datetime
from keyboard import *
import misk
from datetime import datetime

from telebot import types
from copy import deepcopy
from openpyxl.utils import get_column_letter
#from function import *
#from openpyxl import Workbook
from openpyxl import load_workbook
month_table = {'1': "Январь",
		'2': "Февраль",
		'3': "Март",
		'4': "Апрель",
		'5': "Май",
		'6': "Июнь",
		'7': "Июль",
		'8': "Август",
		'9': "Сентябрь",
		'10': "Октябрь",
		'11': "Ноябрь",
		'12': "Декабрь"}

weekday = { '1': "ПН",
		    '2': "ВТ",
		    '3': "СР",
		    '4': "ЧТ",
		    '5': "ПТ",
		    '6': "СБ",
		    '7': "ВС"}




ROW = 0 # Переменная для записи количества групп
DATES = []
# def initialization(flag):
# 	#Получает количество колонок групп при запуске бота
# 	#После обновления информации
# 	#Флаг != 0 - открываем на чтение
# 	#Флаг = 0 - открываем на запись
# 	if flag:
# 		with open("init_student_row.txt", 'r') as file:
# 			ROW = file.read()
# 	else:
# 		col=0
# 		wb =load_workbook('journal.xlsx', read_only = True)
# 		sheet = wb.get_sheet_by_name('Students')
# 		for row in sheet.iter_rows(min_row=1, max_col=100, max_row=1):
# 				for cell in row:
# 					if cell.value == None:
# 						with open ("init_student_row.txt", 'w') as file:
# 							file.write(col)
# 							break
# 					col+=1
# 		ROW =col
				
# 	return row

def get_row(id):
		#
		# считывает строку с последним преподавателем
		#
		wb = load_workbook('journal.xlsx', read_only=True)
		sheet = wb.get_sheet_by_name('Teachers')
		with open('last_row.txt', 'r') as file:
			last_row = file.read()
		cell_range = sheet["A2:A"+last_row]
		for cellObj in cell_range:
			for cell in cellObj:				
				if str(cell.value) == str(id): # id имеет тип int, необходимо преобразование к строке
					return cell.row
def get_mark(name):
	if name in set(who_is_absent):
		return "нет"
	else:
		return "x"

class Teacher():

	name = ''		
	id = ''
	subjects = []


	def __init__(self,id=id, name=name, subjects=subjects):
		self.name = ""#'Ne robit'
		self.id = id
		self.subjects = []


		#
	

	def name_teacher(self, name=None, id=None, subjects=None):
		#
		# Устанавливает значения атрибутов в соответствие с данными в таблице "Teachers" в файле "journal.xlsx"
		#
	
		attributes = ['name', 'subjects']#Teacher.__dict__.keys()

		row = get_row(self.id)
		wb = load_workbook('journal.xlsx')
		col = 2
		sheet = wb.get_sheet_by_name('Teachers')
		for attribute in attributes:
			setattr(teacher, attribute, sheet.cell(column=col, row=row).value) # УСТАНАВЛИВАЕТ ЗНАЧЕНИЕ АТРИБУТОВ (объект, свойства класса, значение)
			col+=1 # переход к следующему значению
		# разбиение строки из третьей колонки на отдельные предметы
		string_subj = self.subjects
		self.subjects = []

		for subject in string_subj.split(', '):
			self.subjects.append(subject)

class Lesson():
	teacher = ''
	subject = ''
	students = []


	def get_list(self, name_group):
		wb =load_workbook('journal.xlsx')
		sheet = wb.get_sheet_by_name('Students')
		col = 1
		row = 1
		students = []
		for row in sheet.iter_rows(min_row=1, max_col=5, max_row=1):
			for cell in row:
				if cell.value == name_group:
					cell_range = sheet[str(cell.column)+str(cell.row+1)+':'+str(cell.column)+"6"]
					for cellObj in cell_range:
						for cell_ in cellObj:
							if cell_.value!=None:
								students.append(cell_.value)
							else:
								break
					break
		return students

	# def keyboard_for_mark():	
	# 	keyboard = types.InlineKeyboardMarkup()
	# 	butns = []

	# 	for name in who_was:
	# 		butns.append(types.InlineKeyboardButton(text = name, callback_data = name))
	# 	butns.append(types.InlineKeyboardButton(text = 'Всё', callback_data = "Всё"))
	# 	butns.append(types.InlineKeyboardButton(text = 'Отмена', callback_data = 'Отмена'))
	# 	keyboard.add(*butns)
	# 	return keyboard

	def write_lesson(self, teacher=None, students=None):
		global current_subject
		global writing_date
		lesson_list = []
		for name in students:
			lesson_list.append(writing_date)
			lesson_list.append(get_weekday(writing_date))
			lesson_list.append(get_month(writing_date))
			lesson_list.append(name)
			lesson_list.append(current_subject)
			lesson_list.append(get_mark(name))
			lesson_list.append(self.teacher)
			write_excel(lesson_list)
			lesson_list = []
		return "Success"
	def __init__(self, teacher = teacher, subject = subject):
		self.teacher = teacher
		self.subject = subject



#### конец класса teacher
####

students = []



teacher = Teacher()
lesson = Lesson()
current_subject = ''
token = misk.token
bot = telebot.TeleBot(token)
who_is_absent = []
who_was = []
sl_u_dates = {}
writing_date = ''

def get_date():
	# date_str =''
	# date_str += str(datetime.today().date.day)+'.'
	# date_str += str(datetime.today().date.month)+'.'
	# date_str += str(datetime.today().date.year)
	# #date_str = str(datetime.date.isoformat(datetime.date.today()))

	# #date_str= date_str.replace("-", ".")
	# print (date_str)
	# return date_str
	today = datetime.today()
	return  today#.strftime("%m.%d.%y") # '04/05/2017'

def get_week_day():
	return weekday[str(datetime.date.isoweekday(datetime.today()))]
def get_weekday(i):
	return weekday[str(datetime.isoweekday(datetime.strptime(i, '%d.%m.%y')))]
def get_month(i):
	return month_table[str(datetime.strptime(i, '%d.%m.%y').month)]
#def check_who_is()


def write_excel(lesson_list):
		with open ("row.txt", "r") as count:
			row = count.read()
		row = int(row)
		wb = load_workbook('my_journal.xlsx')

		sheet = wb.get_sheet_by_name("Journal")

		for col in range(7):			
			sheet.cell(column =col+1, row=row, value=lesson_list[col])
		row+=1
		cell_ = sheet.cell(column = 1, row = int(row))
		cell_.number_format = 'DD\.MM\.YY;@'
		wb.save('my_journal.xlsx')
		with open ("row.txt", 'w') as count:
				count.write(str(row))
		# wb = load_workbook('journal.xlsx')

		# sheet = wb.get_sheet_by_name("Журнал")
		# for row in lesson_list:
		# 	sheet.append(lesson_list)

		# wb.save('journal.xlsx')



@bot.message_handler(commands = ['start'])
def menu(text):
	#запуск бота, проверка по списку пользователей
	if check_who_is(text.from_user.id):

		bot.send_message(chat_id = text.chat.id, text = 'Йоу')



@bot.message_handler(commands = ['file'])
def give_file(text):
	if text.from_user.id in (325726476 , 223103214):
		doc = open('journal.xlsx', 'rb')
		bot.send_document(chat_id= text.chat.id, data = doc)

@bot.message_handler(commands = ['lesson'])
def msg(text):
	global teacher
	global sl_u_dates
	global DATES
	slov = {"ВТ": ["Математика"],
			"ПТ": ["Информатика", "Математика"]}
	teacher.id = text.from_user.id
	teacher.name_teacher()
	lesson.teacher = teacher.name
	count = 1
	date = ''
	temp = []
	message = ''
	today = get_date()
	# with open (str("list_with_dates")+".txt", "r") as file:
	with open (str(text.from_user.id)+".txt", "r") as file:
		for line in file:
			date = line.split(" ")[0]
			subj = line.split(" ")[1][0:-1]
			if  datetime.strptime(date, '%d.%m.%y')<=today:
				DATES.append(date)
			#print (line)
				#for i in slov[get_weekday(line)]:
				#print (i)
				message += str(count)+". "+ date+ "-"+subj+"\n"
				temp.append(count)
				temp.append(date+ "-"+subj)
				sl_u_dates.update([temp])
				count+=1
				temp=[]
			else:
				break
	#bot.send_message(chat_id= text.chat.id, text = str_dates, reply_markup = keyboard_subjects(range(1, size+1)))
	if count !=1:
		bot.send_message(chat_id=text.chat.id, text = message, reply_markup = keyboard_subjects(range(1, count)))
	else:
		bot.send_message(chat_id=text.chat.id, text = "Нет доступных для записи занятий. Попробуйте позже.")


@bot.callback_query_handler(func=lambda text: True)
def func(text):
		global current_subject
		global who_was
		global who_is_absent
		global students
		global DATES
		global sl_u_dates
		global writing_date

		try:
			if int(text.data) in sl_u_dates:
				teacher.subject = sl_u_dates[int(text.data)].split("-")[1]
				current_subject = sl_u_dates[int(text.data)].split("-")[1]
				writing_date = sl_u_dates[int(text.data)].split("-")[0]
				lesson.subject = current_subject
				name_group = current_subject +'-'+ str(teacher.id)
				students = lesson.get_list(name_group)
				who_was = deepcopy(students)
				bot.edit_message_text(chat_id = text.message.chat.id, message_id = text.message.message_id,
				 						text = 'На занятии были все?', reply_markup = keyboard_yes_no(2))
		except:
			if text.data == 'Да':
				bot.edit_message_text(chat_id = text.message.chat.id, message_id = text.message.message_id,
									text = 'По какому предмету?', reply_markup = keyboard_subjects(teacher.subjects))
			elif text.data == 'Отмена':

				bot.edit_message_text(chat_id = text.message.chat.id,
					 			message_id = text.message.message_id,
								text = "Отмена")

			elif text.data == 'Нет':
				for name in teacher.students:
					who_is_absent.append(name)
				lesson.write_lesson()

			elif text.data == 'Нет, не все':
				bot.edit_message_text(chat_id = text.message.chat.id,message_id = text.message.message_id,
									 text = "Кого не было?", reply_markup = keyboard_subjects(students))


			elif text.data == 'Да, все':
				bot.edit_message_text(chat_id = text.message.chat.id, message_id = text.message.message_id,
											text = lesson.write_lesson(teacher.name, students))
				f = open(str(text.from_user.id)+'.txt').read()
				
				
				f = f.replace(writing_date+" "+current_subject+"\n",'')
				
				with open (str(text.from_user.id)+'.txt', "w") as file:
					file.write(f)
			elif text.data == "Всё":
				bot.edit_message_text(chat_id = text.message.chat.id, message_id = text.message.message_id,
											text = lesson.write_lesson(teacher.name, students))
				f = open(str(text.from_user.id)+'.txt').read()
				
				
				f = f.replace(writing_date+" "+current_subject+"\n",'')
				
				with open (str(text.from_user.id)+'.txt', "w") as file:
					file.write(f)
			elif text.data == "Занятия не было":
				for name in students:
					who_was.remove(name)
					who_is_absent.append(name)
				bot.edit_message_text(chat_id = text.message.chat.id, message_id = text.message.message_id,
											text = lesson.write_lesson(teacher.name, students))
				f = open(str(text.from_user.id)+'.txt').read()
				
				
				f = f.replace(writing_date+" "+current_subject+"\n",'')
				
				with open (str(text.from_user.id)+'.txt', "w") as file:
					file.write(f)
			# elif int(text.data) in sl_u_dates:
			# 	teacher.subject = sl_u_dates[int(text.data)].split("-")[1]
			# 	current_subject = sl_u_dates[int(text.data)].split("-")[1]
			# 	writing_date = sl_u_dates[int(text.data)].split("-")[0]
			# 	lesson.subject = current_subject
			# 	name_group = current_subject +'-'+ str(teacher.id)
			# 	students = lesson.get_list(name_group)
			# 	who_was = deepcopy(students)
			# 	bot.edit_message_text(chat_id = text.message.chat.id, message_id = text.message.message_id,
			# 	 						text = 'На занятии были все?', reply_markup = keyboard_yes_no(2))

			#elif text.data in teacher.subjects:

			else:
				if text.data in who_was:
					who_was.remove(text.data)
					who_is_absent.append(text.data)
					bot.edit_message_text(chat_id=text.message.chat.id, message_id = text.message.message_id,
											 text = 'Кого не было?', reply_markup = keyboard_after_delete_name(who_was))

@bot.message_handler(content_types= 'text')
def create_timetable(text):
	# if text.data == 'Файл' and text.from_user.id == #id #Алексея Евгеньевича
	pass
	# if check_who_is(text.from_user.id):
	# 	if text.from_user.id == 223103214:
	# 		bot.send_message(chat_id = text.chat.id, text = 'Йоу', reply_markup= keyboard_for_me())
	# 	a=datetime.datetime.today()
	# 	date= ''
	# 	date=str(a.day)+'.'+str(a.month)+'.'+str(a.year)



bot.polling()