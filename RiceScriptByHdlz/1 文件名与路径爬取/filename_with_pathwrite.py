#!/usr/bin/python 
# -*- coding: utf-8 -*-
import os

filePath = os.getcwd()
filename = r'D:\FTP同步\小包装米\路径.txt'
with open(filename,'w') as f:
    for root, dirs, files in os.walk(filePath):
        for name in files:
    #‘’‘发现文件名很多作废了的，为了加快速度加入一个筛选机制
            if "废" in name:
                cache_name =" "
            elif".DS_Store" in name:
                cache_name =" \n"
            else:
        #这一行打开为输出文件路径和文件名
                cache_name = os.path.join(root, name)
                print(cache_name)
                f.write(str(cache_name)+"\n")
                
    #这一行打开为仅输出文件名
    #   cache_name = os.path.join(name)

#这个代码块不知道干嘛用的
#for name in dirs:
#        print(os.path.join(root, name))
#
print ("文件爬取完毕，按任意键退出~~")
os.system("pause")
