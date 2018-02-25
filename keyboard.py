from telebot import types

yes_no = ["Да", "Нет", "Отмена"]
yes_no_2 = ["Да, все", "Нет, не все", "Отмена"]

# def keyboard_for_me(flag):
# 	keyboard = types.InlineKeyboardMarkup()
# 	butns = []
# 	global prisutstvuiushie
# 	students = names()
# 	if flag == 1:
# 		prisutstvuiushie = deepcopy(inform)
	
# 		for i in students:
# 			butns.append(types.InlineKeyboardButton(text = i, callback_data = i))
# 	if flag == 2:

# 		for i in students:
# 			butns.append(types.InlineKeyboardButton(text = i, callback_data = i))
# 		butns.append(types.InlineKeyboardButton(text = 'Всё', callback_data = "Всё"))
# 		butns.append(types.InlineKeyboardButton(text = 'Отмена', callback_data = 'Отмена'))
# 	keyboard.add(*butns)
# 	return keyboard

def keyboard_after_delete_name(list_of_students):
	#global students
	keyboard = types.InlineKeyboardMarkup()
	butns = []
	for name in list_of_students:
		butns.append(types.InlineKeyboardButton(text = name, callback_data = name))
	butns.append(types.InlineKeyboardButton(text = 'Всё', callback_data = "Всё"))
	butns.append(types.InlineKeyboardButton(text = 'Отмена', callback_data = 'Отмена'))
	keyboard.add(*butns)
	return keyboard



def keyboard_yes_no(flag):
	keyboard = types.InlineKeyboardMarkup()
	butns = []
	if flag == 1:
		for i in yes_no:
			butns.append(types.InlineKeyboardButton(text = i, callback_data = i))
	elif flag == 2:
		for i in yes_no_2:
			butns.append(types.InlineKeyboardButton(text = i, callback_data = i))
	keyboard.add(*butns)
	return keyboard


def keyboard_subjects(subjects):
	keyboard = types.InlineKeyboardMarkup()
	butns = []
	for subject in subjects:
		butns.append(types.InlineKeyboardButton(text = subject, callback_data = subject))
	butns.append(types.InlineKeyboardButton(text = 'Отмена', callback_data = 'Отмена'))
	keyboard.add(*butns)
	return keyboard
