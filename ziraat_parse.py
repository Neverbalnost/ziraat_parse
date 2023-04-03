import pandas as pd

tryRaw = pd.ExcelFile('try.xlsx')
usdRaw = pd.ExcelFile('usd.xlsx')


def getDataTable(file):
	tabnames = file.sheet_names

	ziraatData = file.parse(sheetname=tabnames[0], skiprows=11)
	ziraatData.drop(ziraatData.tail(8).index, inplace = True)
	print(str(len(ziraatData.index)) + ' lines parsed')

	ziraatData.drop('Invoice No.', inplace=True, axis='columns')
	ziraatData.drop('Balance', inplace=True, axis='columns')
	ziraatData.rename({
			  'Transaction Amount':'Outcome'
			}, inplace=True, axis='columns')
	return ziraatData

tryData = getDataTable(tryRaw)
usdData = getDataTable(usdRaw)

incomes = []
categories = []
account = []
recieverAccount = []



for ind in tryData.index:
	comment = tryData['Explanation'][ind]

	if (tryData['Outcome'][ind] > 0):
		incomes.append(tryData['Outcome'][ind])
		tryData['Outcome'][ind] = ''
	else:
		incomes.append(''),
		tryData['Outcome'][ind] = abs(tryData['Outcome'][ind])

	if 'Döviz' in comment:
		usdRow = usdData.loc[usdData['Date'] == tryData['Date'][ind]]
		account.append('Ziraat USD')
		recieverAccount.append('Ziraat TRY')
		tryData['Outcome'][ind] = abs(usdRow.iloc[0]['Outcome'])
	else:
		account.append('Ziraat TRY')
		recieverAccount.append('')


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


tryData.insert(3, 'Income', incomes)
tryData.insert(1, 'Categories', categories)
tryData.insert(0, 'Reciever Account', recieverAccount)
tryData.insert(0, 'Account', account)

tryData = tryData[['Account', 'Reciever Account', 'Date', 'Categories', 'Outcome', 'Income', 'Explanation']]

tryData.to_csv('try_output.csv', encoding='UTF-8', header=False, index=False)

print('Done!')


