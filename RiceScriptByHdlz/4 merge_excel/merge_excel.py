#!/usr/bin/python 
# -*- coding: utf-8 -*-

'''
    Function decribe：
        To merge a buck of sheets in different tables, format is xlxs. The max colunm is known，Max sheet too.
        Then we need a function to merge the ，e.g. sheet1：
            1.xlsx      |2.xlsx     |   ... => merge.xlsx
            A1 B1 C1    |A2 B2 C2   |   ... => A1A2B1B2C1C2...
    Steps：
        0 use xlrd xlwt lib
        1 read and walk all files, get a list named "socop":
            {
                {"filename1",Max_sheet,Max_cols},
                {"filename2",Max_sheet,Max_cols},
                {"filename3",Max_sheet,Max_cols},
                ...
            }
        2 New a file named "merge.xlsx",
        3 write first sheet:
            3.1 walk the socop(Max_sheet) in a new list
            3.2 
'''