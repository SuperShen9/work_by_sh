import shutil,os,zipfile
#coding:utf-8
#copy super.txt to zz
# shutil.copy('C:\z\super.txt','C:\zz')
#copy super.txt to zz and build a new name
# shutil.copy('C:\z\super.txt','C:\zz\super_new.txt')
#copy z to zz
# shutil.copytree('C:\z','C:\zz\z')
#move tang.txt to zz
# shutil.move('C:\z\\tang','C:\zz')
# move tang.txt to zz and build a new name
# shutil.move('C:\zz\\tang.txt','C:\z\\tang_new.txt')
#dele file os method
# os.unlink('C:\z\\tang.txt')
# os.rmdir('C:\zzz')
# shutil.rmtree('C:\zz\z\z1')
#print file in C:\z endwith .txt
# for filename in os.listdir('C:\z\\'):
#     if filename.endswith('.txt'):
#         print (filename)
#send2trash send file to the rubbish
# import send2trash
# superfile=open('super.txt','a')
# superfile.write('super is very clever')
# superfile.close()
# send2trash.send2trash('super.txt')
#List Folder Contents
os.chdir('C:\\')
for foldername,subfolders,filenames in os.walk('C:\z1'):
    print 'folder is' + ' --'+foldername
    for subfolder in subfolders:
        print 'subfolder is' + ' --'+subfolder
    for filename in filenames:
        print 'filename is' +' --'+ filename
#read ZipFile
# os.chdir('C:\\')
# examplezip=zipfile.ZipFile('z.zip')
# print examplezip.namelist()
# superinfo=examplezip.getinfo('super.txt')
# print superinfo.file_size
# print superinfo.compress_size
# examplezip.close()
#jieya ZipFile
# os.chdir('C:\\zzzz')
# examplezip=zipfile.ZipFile('z.zip')
# examplezip.extractall()
# examplezip.close()



