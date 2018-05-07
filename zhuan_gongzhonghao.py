# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import os,shutil
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')

df = pd.read_excel('sheet1.xlsx')
df_PC = pd.read_excel('sheet_PC.xlsx')
df_MV = pd.read_excel('sheet_MV.xlsx')

if os.path.exists('RUN'):
    shutil.rmtree('RUN')
os.makedirs('C:\\Users\Administrator\Desktop\\RUN')

os.chdir('C:\\Users\Administrator\Desktop\\RUN')

for i in range(df.shape[0]):
    fl = open('%s-%s-%s.txt' % (df['name'].loc[i], df['organization'].loc[i], df['webName'].loc[i]), 'a')
    for x in df.columns[:-2]:
        val = df[x].loc[i]
        if isinstance(val, float):
            val = ''
        else:
            val = val.encode('gbk')

        if x == 'webName':
            fl.write('{\n')
            fl.write('\t\r"webUrl": ""\n')
            fl.write('\t\r"{}\r"'.format(x) + ':' + '\r"' + str(val) + '\r"' + ',')
            fl.write("\n")
        else:
            fl.write('\t\r"{}\r"'.format(x) + ':' + '\r"' + str(val) + '\r"' + ',')
            fl.write("\n")
# 合并PC表
    uuid=df['uuid'].loc[i]
    fl.write('\t\r"PCinfo": [{')
    for ii in range(df_PC[df_PC['uuid']==uuid].shape[0]):
        fl.write("\n")
        for xx in df_PC.columns:
            val_pc = df_PC[xx].loc[ii]
            if  xx == 'time' :
                val_pc = val_pc
            else:
                val_pc = val_pc.encode('gbk')
            if xx=='uuid':
                pass
            elif xx=='WeChatSubName':
                fl.write('\t\t\r"{}\r"'.format(xx) + ':' + '\r"' + str(val_pc) + '\r"' + '\n')
            else:
                fl.write('\t\t\r"{}\r"'.format(xx) + ':' + '\r"' + str(val_pc) + '\r"' + ',\n')
        if ii<df_PC[df_PC['uuid']==uuid].shape[0]-1:
            fl.write("\t}, {")
        else:
            pass
    fl.write("\t}],")

# 合并MV表
    fl.write('\n\t\r"MVinfo": [{')
    for iii in range(df_MV[df_MV['uuid']==uuid].shape[0]):
        fl.write("\n")
        for xxx in df_MV.columns:
            val_mv = df_MV[xxx].loc[iii]
            if xxx == 'time':
                val_mv = val_mv
            else:
                val_mv = val_mv.encode('gbk')
            if xxx == 'uuid':
                pass
            elif xxx=='WeChatSubName':
                fl.write('\t\t\r"{}\r"'.format(xx) + ':' + '\r"' + str(val_mv) + '\r"' + '\n')
            else:
                fl.write('\t\t\r"{}\r"'.format(xxx) + ':' + '\r"' + str(val_mv) + '\r"' + ',\n')
        if iii < df_MV[df_MV['uuid'] == uuid].shape[0] - 1:
            fl.write("\t}, {")
        else:
            pass
    fl.write("\t}],")

# 补充最后2个字段
    fl.write('\n\t\r"MDinfo": {')
    for x in df.columns[-2:]:
        val = df[x].loc[i]
        if isinstance(val, float):
            val = ''
        else:
            val = val.encode('gbk')
        fl.write('\n')
        fl.write('\t\t\r"{}\r"'.format(x) + ':' + '\r"' + str(val) + '\r"' + ',')
    fl.write("\n\t}")
    fl.write("\n}")





