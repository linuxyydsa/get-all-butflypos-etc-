import os
import xlwt
import xlrd
import re
import time 

filepath = r'H:\Proto'
#获取日期
#name_name = '210505'
name_name = str(input('which day would you wannna see？\neg:210505\n'))

answer = int(input('do you want to see all tubes or just one tube? \n 1:one tube \n 2:all tubes（It may take 60min） \n'))
if answer == 1:
    CVD = input('which tube would you wanna see?\n eg:PR8-1 is CVD5-1,please input 5-1 \n')


#保存到目标文件
pwd = os.getcwd()
os.chdir(filepath)
pwd = os.getcwd()
pwd_a = os.listdir(filepath)
main_w = xlwt.Workbook(encoding = 'utf-8')
dir_names = []

dir_name = []
dir_dir = []
#拿到文件路径，获取所有文件名
def main():
    

    for i in pwd_a:
        f = re.match(r'CVD\d{1,2}-\d',i)
        
        fj = re.match(r'CVD\d{1,2}-\d.zip',i)
        if fj != None:
            continue
        if f != None:
            fg = f.group()
            dir_dir.append(fg)
            get_pwd = filepath +"\\" +str(i)
            pwd_aa = os.listdir(get_pwd)
            for j in pwd_aa:
                ff = re.match(r'P'+name_name,j)
                if ff != None:
                    get_pwd_aa = get_pwd + "\\" +str(j)
                    dir_name.append(get_pwd_aa)

                    


def main2():


    for i in dir_name:
        get_sheet_name_all = re.search(r'CVD\d{1,2}-\d',i)
        get_sheet_name_all_str = get_sheet_name_all.group()
        count_x = 1
        count_y = 1
        main_s = main_w.add_sheet(get_sheet_name_all_str,cell_overwrite_ok = True)
        getline = open(i,"r",encoding = 'utf-8')  
        for line in getline:

            get_start = re.search('Recipe started',line)
            get_MT1  = re.search(r'\d\d\.\d\d\.\d{1,5} \d\d:\d\d:\d\d L Memory Text1 .*',line)
            
            if get_start != None:
                count_x = 1
                count_y += 5
                if count_y >= 15:
                    break  
            if get_MT1 != None:
            
                get_MT1_str = get_MT1.group()
                get_MT1_str_ti = re.search(r'\d\d:\d\d:\d\d',get_MT1_str)
                get_MT11 = re.search(r'Memory Text1 .*',get_MT1_str)
                get_MT1_str_ti_str = get_MT1_str_ti.group()   #绝对时间
                get_MT1_str_na_str = get_MT11.group()
                
                
                
                main_s.col(count_y+1).width = 13000

                
                main_s.write(count_x,count_y,get_MT1_str_ti_str)
                main_s.write(count_x,count_y+1,get_MT1_str_na_str)
                main_w.save(r'C:\Users\operator\Desktop\严旭（勿动）'+'\\'+'每一步工艺时间对比'+name_name+'.xlsx')
                count_x += 1
        print(get_sheet_name_all_str,' is ok')
    print('loading down！')
    input('press enter to exit！')

def main3():

    
    get_sheet_name_all_str = 'CVD'+str(CVD)
    filename_1 = 'H:\\Proto\\' + get_sheet_name_all_str +'\\P'+name_name+'.txt'
    count_x = 1
    count_y = 1
    main_s = main_w.add_sheet(get_sheet_name_all_str,cell_overwrite_ok = True)
    getline = open(filename_1,"r",encoding = 'utf-8')  
    for line in getline:
        
        get_start = re.search('Recipe started',line)
        get_MT1  = re.search(r'\d\d\.\d\d\.\d{1,5} \d\d:\d\d:\d\d L Memory Text1 .*',line)
            
        if get_start != None:
            count_x = 1
            count_y += 5

            
        if get_MT1 != None:
            
            get_MT1_str = get_MT1.group()
            get_MT1_str_ti = re.search(r'\d\d:\d\d:\d\d',get_MT1_str)
            get_MT11 = re.search(r'Memory Text1 .*',get_MT1_str)
            get_MT1_str_ti_str = get_MT1_str_ti.group()   #绝对时间
            get_MT1_str_na_str = get_MT11.group()
                
                
                
            main_s.col(count_y+1).width = 13000

                
            main_s.write(count_x,count_y,get_MT1_str_ti_str)
            main_s.write(count_x,count_y+1,get_MT1_str_na_str)
            
            count_x += 1
    print(get_sheet_name_all_str,' is ok')
    main_w.save(r'C:\Users\operator\Desktop\严旭（勿动）'+'\\'+'CVD'+str(CVD)+'每一步工艺时间对比'+name_name+'.xlsx')
    print('loading down！')
    
    input('press enter to exit！')

main()
if answer == 1:
    main3()
    
if answer == 2:
    main2()
    


            
