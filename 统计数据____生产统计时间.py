import os
import xlwt
import xlrd
import re
import time 

filepath = r'H:\Proto'
#获取日期


name_name = str(input('which day would you wannna see？\n20220105  —>  220105\n'))


    
    

#保存到目标文件
pwd = os.getcwd()
os.chdir(filepath)
pwd = os.getcwd()
pwd_a = os.listdir(filepath)
main_w = xlwt.Workbook(encoding = 'utf-8')
dir_names = []

dir_name = []
dir_dir = []


count_x = 1
count_y = 2
count = 0

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
                    
                    
            
            #print(get_pwd)#打印文件夹名称
           
    #        pwd_b = os.open(get_pwd)
    #        print(pwd_b)
        else:
            continue
    return dir_name



def main2():

    

    title = ['工艺结束时间','工艺舟号','舟使用次数','工艺时间','二次侧漏','测漏','三次测漏','四次侧漏']
    
    for i in dir_name:
        get_sheet_name_all = re.search(r'CVD\d{1,2}-\d',i)
        get_sheet_name_all_str = get_sheet_name_all.group()

        
        count_x = 1
        title_count = 0
        control = 0

        
        main_s = main_w.add_sheet(get_sheet_name_all_str,cell_overwrite_ok = True)
        main_s.col(count_y-2).width = 5000
        main_s.col(count_y).width = 2500

        for w_title in title:
            main_s.write(0,int(title_count),w_title)
            title_count += 1
        
        getline = open(i,"r",encoding = 'utf-8')  
        for line in getline:

            get_start = re.search('Recipe started',line)
            get_leak  = re.search(r'O leak, pressure rise \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
            get_leak1 = re.search(r'O pressure rise test \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
            get_leak2 = re.search(r'Try no.:  2 . Leak, pressure rise \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
            get_leak3 = re.search(r'Try no.:  3 . Leak, pressure rise \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
            pd_data_run = re.search(r'O  Boat runs:  \d{1,3}',line)
            pd_data = re.search(r'\d\d.\d\d.\d{1,4} \d\d:\d\d:\d\d.*Boat:.*[A-Za-z][A-Za-z]\d\d.Runtime:.*\d{1,2}:\d{1,2}:\d{1,2}',line)

            
            if get_start != None:
                control = 1
            if control ==1:
                if get_leak != None :
                    get_leak_f = get_leak.group()
                    get__leak = re.search(r'\d{1,5}.\d{1,5}',get_leak_f)
                    adc = get__leak.group()
                    
                    
                    main_s.write(count_x,count_y+2,adc)
                    
                if get_leak1 != None:
                    get_leak1_f = get_leak1.group()
                    get__leak1 = re.search(r'\d{1,5}.\d{1,5}',get_leak1_f)
                    aec = get__leak1.group()
                    main_s.write(count_x,count_y+3,aec)
                    
                if get_leak2 != None :
                    get_leak2_f = get_leak2.group()
                    get__leak2 = re.search(r'\d{1,5}.\d{1,5}',get_leak2_f)
                    agc = get__leak2.group()
                    main_s.write(count_x,count_y+4,agc)
                    
                if get_leak3 != None :
                    get_leak3_f = get_leak3.group()
                    get__leak3 = re.search(r'\d{1,5}.\d{1,5}',get_leak3_f)
                    afc = get__leak3.group()

                    
                    main_s.write(count_x,count_y+5)
                if pd_data_run != None :
                    pd_data_run_f = pd_data_run.group()
                    pd_data_run__f = re.search(r'\d{1,3}',pd_data_run_f)
                    akc = pd_data_run__f.group()
                    
                    
                    main_s.write(count_x,count_y,akc)
 
        
                
                
                if pd_data != None:
                    pd_datag = pd_data.group()
                    
                    get__time = re.search(r'Runtime:.*\d{1,2}:\d{1,2}:\d{1,2}',pd_datag)
                    get__boat = re.search(r'Boat:.*[A-Za-z][A-Za-z]\d\d',pd_datag)
                    get__boat_time = re.search(r'\d\d.\d\d.\d{1,4} \d\d:\d\d:\d\d',pd_datag)
                    get__time_runtime = get__time.group()
                    get__boat_time_boat = get__boat_time.group()
                    get__boat_boat = get__boat.group()
                    get___time = re.search(r'\d{1,2}:\d{1,2}:\d{1,2}',get__time_runtime)
                    get___boat = re.search(r'[A-Za-z][A-Za-z]\d\d',get__boat_boat)
                    
                    if get__time != None:
                        
                        abc = get___time.group()#工艺时间
                        acc = get___boat.group()  #舟号
                        alc = get__boat_time.group()#结束时间
                        
                        main_s.write(count_x,count_y+1,abc)
                        main_s.write(count_x,count_y-1,acc)
                        main_s.write(count_x,count_y-2,alc)
                        count_x += 1

        main_w.save(r'C:\Users\operator\Desktop\严旭（勿动）'+'\\'+'生产统计'+name_name+'.xlsx')  
        print(get_sheet_name_all_str,' is ok')
            

    print('loading down！')
    input('press enter to exit！')


main()

main2()





