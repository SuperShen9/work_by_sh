#coding=utf-8
# 猜拳游戏
# 游戏规则
# 石头剪刀布 
#引入模块

import random
import sys #系统调用模块


winlist = [['石头','剪刀'],['剪刀','布'],['布','石头']]
#选择列表
choicelist = ['石头','剪刀','布']
#用户提示  ''' ''' 换行输出
prompt ='''可选项如下
(0)石头
(1)剪刀
(2)布
(3)退出
请输入你的选择(输入数字即可)'''
while True:
	try:
		#用户选择一个数字
		choicenum = int(raw_input(prompt))
		if choicenum == 3:
			break #跳出循环 结束
		userchoice = choicelist[choicenum]
		#电脑选择
		comchoice = random.choice(choicelist)
		#把用户选择和电脑选择的放在列表中
		bothchoice = [userchoice,comchoice]
		print '你选择了%s,电脑选择了%s'%(userchoice,comchoice)
		#判断输赢
		if userchoice == comchoice:
			print '平手'
		elif bothchoice in winlist:
			print '你赢了'
		else:
			print '你输了,要不要再来一局'
	except(KeyboardInterrupt,EOFError,ValueError,IndexError):
		sys.exit()

