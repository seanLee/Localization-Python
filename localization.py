from openpyxl import load_workbook

workbook = load_workbook(filename = './source.xlsx')

fileName = "./NNLocalization.strings"
fileToWrite = "./To/"

#数字键对应的文件名
countryList = {
        0: "CN_",
        1: "CN",
        2: "English",
        3: "Spanish",
        4: "Portuguese",
        5: "French",
        6: "German",
        7: "Italian",
        8: "Russian",
        9: "Danish",
        10:  "Norwegian",
        11:  "Swedish",
        12:  "Romanian",
        13:  "Bulgarian",
        14:  "Greek",
        15:  "Czech",
        16:  "Slovak",
        17:  "Dutch",
        18:  "Hungarian",
        19:  "Test",
        20:  "Polish",
        21:  "CH_Tra",
        22:  "",
        23:  "",
    }
#英文文案的键值
pairs = {}
#储存文案
toDict = {}

with open(fileName, "r") as file:
	for line in file.readlines():
		if "=" in line:
			keyValue = line.split("=", 1)
			#key
			key = keyValue[0]
			key = key[1:len(key)-2]
			#value
			value = keyValue[1]
			value = value[2:len(value)-3]
			pairs[key] = value

print("all ---> ", str(len(pairs.keys())))

for sheet in workbook:
	temp = {}
	for index in range(5,len(tuple(sheet.rows))):
		row = tuple(sheet.values)[index]
		col_two = row[1]
		if col_two is None:
			continue
		#get tht english value
		col_english = row[2]
		#the key exist in english map
		local_key = ""
		keyList = []
		for key in pairs.keys():
			if pairs[key] == col_english:
				local_key = key
				keyList.append(key)

		if (len(keyList) == 0):
			continue

		# for value in row:
		for row_index in range(1,22):
			value = row[row_index]
			_toDict = {}
			if row_index in toDict.keys():
				_toDict = toDict[row_index]
			else:
				_toDict = {}

			if value is None:
				value = col_english

			for key in keyList:
				_toDict[key] = value

			toDict[row_index] = _toDict


for key in toDict.keys():
	curDict = toDict[key]
	toFileName = countryList[key]
	sortedKey = list(curDict.keys())
	#排序
	sortedKey.sort()
	toPath = fileToWrite+toFileName+".txt"
	with open(toPath, "w") as file_to_write:
		for toKey in sortedKey:
			file_to_write.write('"' + toKey + '" = "' + curDict[toKey] + '";' + "\n")
		file_to_write.close()


# b中有而a中没有的
print (list(set(pairs.keys()).difference(set(curDict.keys())))) 




