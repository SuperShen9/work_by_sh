# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import os,shutil
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')

df = pd.read_excel('sheet1.xlsx')


if os.path.exists('RUN'):
    shutil.rmtree('RUN')
os.makedirs('C:\\Users\Administrator\Desktop\\RUN')

os.chdir('C:\\Users\Administrator\Desktop\\RUN')

for i in range(df.shape[0]):
    os.chdir('C:\\Users\Administrator\Desktop')
    df_PC = pd.read_excel('sheet_PC.xlsx')
    df_MV = pd.read_excel('sheet_MV.xlsx')
    os.chdir('C:\\Users\Administrator\Desktop\\RUN')
    fl = open('%s-%s-%s.txt' % (df['name'].loc[i], df['organization'].loc[i], df['webName'].loc[i]), 'a')
    for x in df.columns[:-2]:
        val = df[x].loc[i]
        if isinstance(val, float):
            val = ''
        else:
            val = val.encode('gbk')

        if x == 'webName':
            fl.write('{\n')
            fl.write('\t\r"webUrl": "",\n')
            fl.write('\t\r"{}": "{}",\n'.format(x, str(val)))

        else:
            fl.write('\t\r"{}": "{}",\n'.format(x, str(val)))
# 合并PC表
    uuid = df['uuid'].loc[i]

    fl.write('\t\r"PCinfo": [{')
    # print df_PC[df_PC['uuid'] == uuid].shape[0], df_MV[df_MV['uuid'] == uuid].shape[0]
    # print df_PC
    # exit()
    df_PC = df_PC[df_PC['uuid'] == uuid]
    df_PC = df_PC.reset_index(drop=True)
    if df_PC[df_PC['uuid'] == uuid].shape[0]==0:
        fl.write('"time": "","medicalPlatform": "","subjectContent": "","isMeeting": "","WeChatSubName": ""')
        fl.write("}],")
    else:
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
                    fl.write('\t\t\r"{}": "{}"\n'.format(xx, str(val_pc).replace('\n','')))

                else:
                    fl.write('\t\t\r"{}": "{}",\n'.format(xx, str(val_pc).replace('\n','')))
            if ii<df_PC[df_PC['uuid']==uuid].shape[0]-1:
                fl.write("\t}, {")
            else:
                pass
        fl.write("\t}],")

# 合并MV表
    fl.write('\n\t\r"MVinfo": [{')
    df_MV = df_MV[df_MV['uuid'] == uuid]
    df_MV = df_MV.reset_index(drop=True)
    if df_MV[df_MV['uuid'] == uuid].shape[0] == 0:
        fl.write('"time": "","weChatViewpoint": "","WeChatSubName": "","isIndependentPub": ""')
        fl.write("}],")
    else:

        for iii in range(df_MV[df_MV['uuid']==uuid].shape[0]):
            fl.write("\n")
            for xxx in df_MV.columns:
                val_mv = df_MV[xxx].loc[iii]
                if xxx == 'time'or xxx=='isIndependentPub':
                    val_mv = val_mv
                else:
                    val_mv = val_mv.encode('gbk')
                if xxx == 'uuid':
                    pass
                elif xxx=='isIndependentPub':
                    fl.write('\t\t\r"{}": "{}"\n'.format(xxx, str(val_mv).replace('\n','')))
                else:
                    fl.write('\t\t\r"{}": "{}",\n'.format(xxx, str(val_mv).replace('\n','')))

            if iii < df_MV[df_MV['uuid'] == uuid].shape[0] - 1:
                fl.write("\t}, {")
            else:
                pass
        fl.write("\t}],")

# 补充最后2个字段
    fl.write('\n\t\r"MDinfo": {')
    fl.write('\n')
    for x in df.columns[-2:]:
        val = df[x].loc[i]
        if isinstance(val, float):
            val = ''
        else:
            val = val.encode('gbk')
        if x=='WeChatSubName':
            fl.write('\t\t\r"{}": "{}"\n'.format(x, str(val)))
        else:
            fl.write('\t\t\r"{}": "{}",\n'.format(x, str(val)))
    fl.write("\t}")
    fl.write("\n}")





