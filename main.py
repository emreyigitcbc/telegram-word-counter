import json
import xlsxwriter
import difflib
# Export your telegram chat file as "json"
with open('result.json', 'r', encoding="utf8") as f:
    veriMain = json.load(f)
collectedData = {}
rawData = []
userList = {}
# It scans for most close word. (Example: "xxx" will be counted as "xxy")
checkForDifference = True
collectList = ["some","words","here"]
collectListCopy = {collect: 0 for collect in collectList} | {"total_message": 0}
# Calculating raw collected data
for message in veriMain["messages"]:
    if message["type"] == "message":
        if message["from"] not in collectedData:
            collectedData[message["from"]] = dict(collectListCopy)
            userList[message["from"]] = len(rawData)
            rawData.append([0 for _ in range(len(collectList)+1)])
        collectedData[message["from"]]["total_message"] += 1
        rawData[userList[message["from"]]][len(collectList)] += 1
        contents = message["text_entities"]
        for content in contents:
            content = content["text"].lower()
            if content != "":
                for collected in collectList:
                    if collected in content:
                        collectedData[message["from"]][collected] += 1
                        rawData[userList[message["from"]]][collectList.index(collected)] += 1
                    elif checkForDifference:
                        for altWord in (content + " ").split(" "):
                            ratio = difflib.SequenceMatcher(None, collected, altWord).ratio()
                            if ratio >= 0.67:
                                print(collectedData)
                                collectedData[message["from"]][collected] += 1
                                rawData[userList[message["from"]]][collectList.index(collected)] += 1
                                print(collectedData)

# Creating excel file
workbook = xlsxwriter.Workbook('collected.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
# Writing titles
for i in range(len(rawData)):
    worksheet.write(1+i, 0, list(userList.keys())[i], bold)
for i in range(len(collectList)):
    worksheet.write(0, i+1, collectList[i], bold)
worksheet.write(0, len(collectList)+1, "TOTAL MESSAGE",bold)
worksheet.write(0, len(collectList)+2, "TOTAL COLLECTED",bold)
worksheet.write(0, len(collectList)+3, "% COLLECTED WORD",bold)
worksheet.write(0, len(collectList)+4, "MOST USED WORD",bold)
# Kişilerin küfür verilerini yazma
collectedCompByRatio = []
collectedCompByAmmount = []
for user in userList:
    for index, data in enumerate(rawData[userList[user]]):
        worksheet.write(1+userList[user], 1+index, data)
    totalcollected = sum(rawData[userList[user]][0:len(rawData[userList[user]])-1])
    collectedCompByAmmount.append((user, totalcollected))
    worksheet.write(1+userList[user], 1+len(collectList)+1, totalcollected)
    collectedRate = (100 * totalcollected) / rawData[userList[user]][-1]
    collectedCompByRatio.append((user, float("{:.2f}".format(collectedRate))))
    worksheet.write(1+userList[user], 1+len(collectList)+2, float("{:.2f}".format(collectedRate)))
    (maxcollected,i) = max((v,i) for i,v in enumerate(rawData[userList[user]][0:len(rawData[userList[user]])-1]))
    worksheet.write(1+userList[user], 1+len(collectList)+3, collectList[i] + " - " + str(maxcollected))
collectedCompByRatio.sort(key=lambda tup: tup[1], reverse=True)
collectedCompByAmmount.sort(key=lambda tup: tup[1], reverse=True)
worksheet.write(1+len(rawData)+2, 1, "Ranking: (By Ratio)", bold)
for index, user in enumerate(collectedCompByRatio):
    worksheet.write(1+len(rawData)+3+index, 1, str(index+1) + ". " + user[0])
worksheet.write(1+len(rawData)+2, 5, "Ranking: (By Count)", bold)
for index, user in enumerate(collectedCompByAmmount):
    worksheet.write(1+len(rawData)+3+index, 5, str(index+1) + ". " + user[0])
workbook.close()
with open('collected.json', 'w', encoding='utf8') as f:
    try:
        json.dump(collectedData, f, ensure_ascii=False, indent=4)
    finally:
        f.close()
print("Done. Files saved as 'collected.json' and 'collected.xlsx'")