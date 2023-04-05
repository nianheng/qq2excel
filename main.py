import xlwt  # 用来进行excel写入的库

filename = 'data.txt'

with open(filename, 'r', encoding='utf-8') as f:
    text = f.readlines()  # 读入聊天记录
    text = text[8:]  # 去掉前几行

file = []  # 存储处理之后的聊天记录

num = 0  # 当前处理到第几行

while num < len(text):  # 逐行读入进行处理
    title = text[num]  # 时间戳以及用户名
    num += 1

    content = text[num]

    # content是这条聊天记录的内容，可能有多行，因此需要使用while
    while num + 1 < len(text) and (text[num + 1][:3] != '20' or text[num + 1][4] != '-' or text[num + 1][7] != '-'):
        num += 1
        content += text[num]
    num += 1

    title = title.split()  # 将时间戳与用户名分离
    day = title[0]
    time = title[1]
    name = ''.join(title[2:])

    # 存储处理后的单条聊天记录以及时间戳，姓名
    file.append({
        'day': day,
        'time': time,
        'name': name,
        'content': content
    })

# 将处理后的聊天记录存入excel

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet("history0")

worksheet.write(0, 0, "DAY")
worksheet.write(0, 1, "TIME")
worksheet.write(0, 2, "NAME")
worksheet.write(0, 3, "CONTENT")

cnt = 0
count = 0
for message in file:
    cnt += 1
    worksheet.write(cnt, 0, message['day'])
    worksheet.write(cnt, 1, message['time'])
    worksheet.write(cnt, 2, message['name'])
    worksheet.write(cnt, 3, message['content'])

    # 因为xls格式的excel最多65535行，因此如果超过这个行数就要建一个新的sheet
    if cnt == 60000:
        cnt = 0
        count += 1
        worksheet = workbook.add_sheet("history" + str(count))

workbook.save("history.xls")
