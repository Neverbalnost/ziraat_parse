import pandas as pd

fileName = input('Имя файла, будьте любезны ')

isUsd = input('Это USD счет? (y/n) ') == 'y'

file = pd.ExcelFile(fileName)

tabnames = file.sheet_names

ziraatData = file.parse(sheetname=tabnames[0], skiprows=11)
ziraatData.drop(ziraatData.tail(8).index, inplace = True)
print(str(len(ziraatData.index)) + ' lines parsed')

ziraatData.drop('Invoice No.', inplace=True, axis='columns')
ziraatData.drop('Balance', inplace=True, axis='columns')
ziraatData.rename({
		  'Transaction Amount':'Outcome'
		}, inplace=True, axis='columns')

incomes = []
categories = []
account = []
recieverAccount = []

for ind in ziraatData.index:
	comment = ziraatData['Explanation'][ind]

	if (ziraatData['Outcome'][ind] > 0):
		incomes.append(ziraatData['Outcome'][ind])
		ziraatData['Outcome'][ind] = ''
	else:
		incomes.append(''),
		ziraatData['Outcome'][ind] = abs(ziraatData['Outcome'][ind])


	if 'GETIR' in comment:
		categories.append('Продукты / Getir')
	elif 'TUTARI' in comment or 'Tahsilatı' in comment:
		categories.append('Комиссии')
	elif 'Petravichyus' in comment:
		categories.append('Перевод Саше')
	elif 'AVEA' in comment or 'TURKCELL' in comment or 'LODOSNET' in comment:
		categories.append('Телефон и интернет')
	elif 'TRENDYOL' in comment:
		categories.append('Trendyol')
	elif 'IZSU' in comment or 'Gediz' in comment or 'Gaz' in comment:
		categories.append('КУ')
	elif 'MIGROS' in comment or 'UNLU' in comment or 'PEHLIVANOGL' in comment or 'MMM' in comment or 'METRO' in comment or 'BIM' in comment:
		categories.append('Продукты')
	elif 'ESOS' in comment:
		categories.append('Бухлишко')
	elif 'GOOGLE' in comment:
		categories.append('Подписки')
	elif 'Fatma Postalcilar' in comment:
		categories.append('Жильё')
	elif 'IZMIRIM KART' in comment:
		categories.append('Проезд')
	elif 'ECZANESI' in comment or 'MEDIKAL' in comment:
		categories.append('Лекарства и здоровье')
	elif 'OZCE OYUNCAK' in comment:
		categories.append('Для дома')
	else:
		categories.append('')


	if (isUsd):
		account.append('Ziraat USD')
		if ('Döviz' in comment):
			recieverAccount.append('Ziraat TRY')
		else:
			recieverAccount.append('')
	else:
		if ('Döviz' in comment):
			ziraatData.drop(ind)

ziraatData.insert(3, "Income", incomes)
ziraatData.insert(1, "Categories", categories)
ziraatData.insert(0, "Account", account)
ziraatData.insert(1, "Reciever Account", recieverAccount)
ziraatData = ziraatData[['Date', 'Account', 'Reciever Account', 'Categories', 'Outcome', 'Income', 'Explanation']]

saveDestination = input('Куда сохранять изволите? ')

ziraatData.to_csv(saveDestination + '.csv', encoding='UTF-8', header=False)

print('Готово!')


