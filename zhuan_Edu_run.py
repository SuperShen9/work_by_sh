# -*- coding: utf-8 -*-
# author:Super
import pandas as pd
import os,shutil
pd.set_option('expand_frame_repr',False)
os.chdir('C:\\Users\Administrator\Desktop')

df = pd.read_excel('sheet1.xlsx')
df_edu = pd.read_excel('sheet_edu.xlsx')
df_edu = df_edu.sort_values(by=['uuid','year','unitName'],ascending=[1,1,1])


if os.path.exists('RUN'):
    shutil.rmtree('RUN')
os.makedirs('C:\\Users\Administrator\Desktop\\RUN')

os.chdir('C:\\Users\Administrator\Desktop\\RUN')


for i in range(df.shape[0]):
    fl = open('%s-%s-%s.txt' % (df['name'].loc[i], df['organization'].loc[i], df['webName'].loc[i]), 'a')
    for x in df.columns:
        val = df[x].loc[i]
        if isinstance(val, float):
            val = ''
        else:
            val = val.encode('gbk')

        if x == 'webName':
            fl.write('{\n')
            fl.write('\t\r"webUrl": "http://xhgb.cma.org.cn/xuehui_project/listProjectGongbu.jsp?projectLevel=2&orgId=200100"\n')
            fl.write('\t\r"{}": "{}",\n'.format(x, str(val)))

        else:
            fl.write('\t\r"{}": "{}",\n'.format(x, str(val)))

    uuid=df['uuid'].loc[i]
    fl.write('\t\r"years": [{\n')
    df_ren = df_edu[df_edu['uuid'] == uuid]
    df_ren=df_ren.reset_index(drop=True)

    # print df_ren
    # exit()
    # print df_ren.groupby('year').first().reset_index()
    # exit()
    # print len(df_ren)
    count_all = 0
    for y in df_ren.groupby('year').first().reset_index()['year']:
        df_year=df_ren[df_ren['year']==y]
        fl.write('\t\t\t\r"year": "{}",'.format(y) + '\n')
        fl.write('\t\t\t\r"info": [{' + '\n')

        count=0

        all=len(df_year.groupby('unitName').first().reset_index()['unitName'])
        for n in df_year.groupby('unitName').first().reset_index()['unitName']:
            count+=1

            # print count_all
            df_name = df_year[df_ren['unitName'] == n]
            df_name = df_name.reset_index(drop=True)

            fl.write('\t\t\t\t\r"unitName": "{}",'.format(n.encode('gbk')) + ',\n')
            fl.write('\t\t\t\t\r"unitInfo": [{\n')
            for ii in range(df_name.shape[0]):
                #原来底层循环在这
                count_all += 1

                for xx in df_name.columns:
                    val_edu = df_name[xx].loc[ii]
                    if xx == 'holdingPeriod' or xx == 'creditHour' or xx == 'year' :
                        val_edu = val_edu
                    else:
                        val_edu = val_edu.encode('gbk')
                    if ii!=df_name.shape[0]-1:
                        if xx=='uuid' or xx=='year'or xx=='unitName' :
                            pass
                        elif xx=='creditHour':
                            fl.write('\t\t\t\t\t\t\r"{}": "{}"\n'.format(xx, str(val_edu)))
                            fl.write('\t\t\t\t\t\r},\n\t\t\t\t\t\r{\n')
                        else:
                            fl.write('\t\t\t\t\t\t\r"{}": "{}",\n'.format(xx, str(val_edu)))
                    else:
                        if xx=='uuid' or xx=='year'or xx=='unitName' :
                            pass
                        elif xx == 'creditHour':
                            fl.write('\t\t\t\t\t\t\r"{}": "{}"\n'.format(xx, str(val_edu)))
                        else:
                            fl.write('\t\t\t\t\t\t\r"{}": "{}",\n'.format(xx, str(val_edu)))

            if count_all==len(df_ren):
                fl.write("\t\t\t\t\t}\n\t\t\t\t]\n\n\t\t\t}]\n\t\t\r}\n\t\r]\n}")
            else:
                fl.write("\t\t\t\t\t}\n\t\t\t\t]\n\n\t\t\t}]\n\t\t\r},\n\t\t\r{\n")













