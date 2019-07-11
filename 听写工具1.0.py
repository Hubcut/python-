import pyttsx3
import random
import easygui
import xlsxwriter

speaker = pyttsx3.init()
correct =0
wrong =0
all_number = easygui.integerbox('数量')
number = 0
word = []
number_sheet = 1

#写入excel的原始数据
workbook = xlsxwriter.Workbook('重听.xlsx')
worksheet = workbook.add_worksheet('数据')
headings = ['听写单词','单词答案','是否正确']
worksheet.write('A1',headings[0])
worksheet.write('B1',headings[1])
worksheet.write('C1',headings[2])

#获取听写的单词
# noinspection PyInterpreter
def listing ( ):
        for x in range (all_number):
                word_now = easygui.enterbox('听写的单词')
                word.append(word_now)
#开始听写
        for x in range (all_number):
        number_sheet += 1
        #number_sheet是我建立计量行数的变量
        number = random.randint(0,len(word) - 1)
        #number 是用于抽取单词随机数的变量
        word_answer = word[number]
        speaker.say('ready to listen word')
        #speaker.say（） 括号里的内容会被扬声器发音出来
        for y in range(2):
                speaker.say(word_answer)
        speaker.runAndWait()
        #runAndWait()等待的意思
        answer =easygui.enterbox('刚刚听写的单词')
        #开始要求你开始听写了
        worksheet.write('A'+str(number_sheet),word_answer)
        worksheet.write('B'+str(number_sheet),answer)
        
        #判断正确或错误，并作出对应的行为
        if answer == word_answer:
                #如果输入等于答案那么就写入 C列 第 number_sheet（for 循环第一句时候出现过这个变量，是用于计量该把数据写入第几行的值)行，写入数据为正确。
                worksheet.write("C"+str(number_sheet),'正确')
                easygui.msgbox('正确')
                correct += 1
        else:
                #如果输入等于答案那么就写入 C列 第 number_sheet（for 循环第一句时候出现过这个变量，是用于计量该把数据写入第几行的值)行，写入数据为错误。 
                worksheet.write('C'+str(number_sheet),'错误')
                easygui.msgbox('错误')
                wrong += 1
        #列表后面跟着的pop () 括号里的是删除的项
        word.pop(number)
listing_now = listing()
#最后的提示标语
number_sheet += 1 
#写入总量
worksheet.write('A'+str(number_sheet),'听写总量：'+str(all_number))
workbook.close()
easygui.msgbox('你答对了'+str(correct),'个','数据已经保存')

        
