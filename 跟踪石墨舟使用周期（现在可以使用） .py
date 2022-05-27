import os
import xlwt
import xlrd
import re
import time


filepath = r'H:\Proto'
#获取日期；


name_name = []


#获取机台号；


#保存到目标文件
pwd = os.getcwd()
os.chdir(filepath)
pwd = os.getcwd()


main_w = xlwt.Workbook(encoding = 'utf-8')
dir_names = []

dir_name = []
dir_dir = []




def get_tube():


    title = ['工艺结束时间','工艺舟号','舟使用次数','工艺时间','二次侧漏','测漏','三次测漏','四次测漏','开始时间','结束时间','工艺名称','底压','中断','机台号']

    name_jitai=[]
    tube_name = []
    goto = {8:5 , 9:2 , 10:7 , 11:12 , 12:13 , 13:8 , 14:10 , 15:11 , 16:9 , 17:6 , 18:14 , 19:3 }
    main_w = xlwt.Workbook(encoding = 'utf-8')

    while True:
    
        name_name_1 = int(input('which days would you wannna see？\n20220105  —>  220105 enter 0 to stop\n'))
        if name_name_1 == 0:
            break
        name_name.append(name_name_1)
        #print(name_name)


    while True:
        name_jitai_1 = int(input('plz enter your machine number \n eg:PR08 please enter 8 \n enter 0 to stop\n'))
        if name_jitai_1 == 0:
            break
        name_jitai_1 = goto[name_jitai_1]
        name_jitai.append(name_jitai_1)
        #print(name_jitai)
    print('please wait......')
    tube = [-1,-2,-3,-4]
    for i in name_jitai:
        #文本名

        count_x = 1
        count_y = 2
        count = 0
        
        sheetname = 'CVD' + str(i)
        main_s = main_w.add_sheet(sheetname,cell_overwrite_ok = True)



        title_count = 0
        control = 0

        
        main_s.col(count_y-2).width = 5000
        main_s.col(count_y).width = 2500
        main_s.col(count_y+8).width = 8500


        for w_title in title:
            main_s.write(0,int(title_count),w_title)
            title_count += 1









        
        #坐标值
        tjitai_name_count = str(i)
        jitai = 'jitai' + str(i)
        jitai1 = []
        for j in tube:
        

            for k in name_name:
                tube_name_j = filepath+ '\\' + 'CVD' + str(i) + str(j) + '\\' + "P" + str(k) +'.txt'


                getline = open(tube_name_j,"r",encoding = 'utf-8')
                #print(getline)
                for line in getline:

                    get_start = re.search('Recipe started',line)
                    get_start_time = re.search('\d\d.\d\d.\d{1,4} \d\d:\d\d:\d\d R Recipe started',line)
                    get_leak  = re.search(r'O leak, pressure rise \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
                    get_leak1 = re.search(r'O pressure rise test \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
                    get_leak2 = re.search(r'Try no.:  2 . Leak, pressure rise \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
                    get_leak3 = re.search(r'Try no.:  3 . Leak, pressure rise \[mTorr/min\]: \d{1,5}.\d{1,5}',line)
                    pd_data_run = re.search(r'O  Boat runs:  \d{1,3}',line)
                    pd_data = re.search(r'\d\d.\d\d.\d{1,4} \d\d:\d\d:\d\d.*Boat:.*[A-Za-z][A-Za-z]\d\d.Runtime:.*\d{1,2}:\d{1,2}:\d{1,2}',line)
                    get_end_time = re.search('\d\d.\d\d.\d{1,4} \d\d:\d\d:\d\d R Recipe End Recipe',line)
                    pd_data_name_gongyimingcheng = re.search(r'\d\d.\d\d.\d{1,4} \d\d:\d\d:\d\d R Recipe Start Recipe:      /PROCESS/.*;\d{1,2}.*',line)
                    diya = re.search(r'base pressure:.*',line)
                    zhongduan = re.search(r'.*process abort.*',line)

                    
                    
                    if get_start != None:
                        control = 1
                    
                    if control ==1:
                        if get_leak != None :
                            get_leak_f = get_leak.group()
                            get__leak = re.search(r'\d{1,5}.\d{1,5}',get_leak_f)
                            adc = get__leak.group()
                            adc = float(adc)
                            
                            
                            main_s.write(count_x,count_y+2,adc)
                            #print(adc)
                            
                        if get_leak1 != None:
                            get_leak1_f = get_leak1.group()
                            get__leak1 = re.search(r'\d{1,5}.\d{1,5}',get_leak1_f)
                            aec = get__leak1.group()
                            aec = float(aec)
                            main_s.write(count_x,count_y+3,aec)
                            
                        if get_leak2 != None :
                            get_leak2_f = get_leak2.group()
                            get__leak2 = re.search(r'\d{1,5}.\d{1,5}',get_leak2_f)
                            agc = get__leak2.group()
                            agc = float(agc)
                            main_s.write(count_x,count_y+4,agc)
                            
                        if get_leak3 != None :
                            get_leak3_f = get_leak3.group()
                            get__leak3 = re.search(r'\d{1,5}.\d{1,5}',get_leak3_f)
                            afc = get__leak3.group()
                            afc = float(afc)

                            
                            main_s.write(count_x,count_y+5)
                        if pd_data_run != None :
                            pd_data_run_f = pd_data_run.group()
                            pd_data_run__f = re.search(r'\d{1,3}',pd_data_run_f)
                            akc = pd_data_run__f.group()
                            akc = float(akc)
                        
                            
                            
                            main_s.write(count_x,count_y,akc)
         
                
                        if get_start_time != None :
                            get_start_time_f = get_start_time.group()
                            get__start = re.search(r'\d\d:\d\d:\d\d',get_start_time_f)
                            ayc = get__start.group()
                    
                            #print(ayc)
                            main_s.write(count_x,count_y+6,ayc)




                        if get_end_time != None :
                            get_end_time_f = get_end_time.group()
                            get__end = re.search(r'\d\d:\d\d:\d\d',get_end_time_f)
                            azc = get__end.group()
                            #print('1',azc)
                            
                            main_s.write(count_x,count_y+7,azc)


                
                        if pd_data_name_gongyimingcheng != None :
                            pd_data_name_gongyimingcheng_f = pd_data_name_gongyimingcheng.group()
                            #print(pd_data_name_gongyimingcheng_f)
                            pd_data_name_gongyimingcheng__f = re.search(r'/PROCESS/[^;]*;\d{0,2}',pd_data_name_gongyimingcheng_f)
                            anmc = pd_data_name_gongyimingcheng__f.group()
                            #print(anmc)
                            
                            main_s.write(count_x,count_y+8,anmc)

                
                        if diya != None :
                            diya_f = diya.group()
                            get__diya = re.search(r'[-| ]\d{1,10}.\d{1,10}',diya_f)
                            
                            
                            if get__diya != None:
                                annc = get__diya.group()
                                annc = float(annc)
                                #print(annc)
                                main_s.write(count_x,count_y+9,annc)

                        
                        if zhongduan != None :
                            zhongduan_f = zhongduan.group()
                            get__zhongduan = re.search(r'\w{1,100} failure',zhongduan_f)
                            
                            if get__zhongduan != None:
                                ancc = get__zhongduan.group()
                                #print(ancc)
                                main_s.write(count_x,count_y+10,ancc)

                                


                        
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
                                main_s.write(count_x,count_y+11,str(i)+str(j))
                                #print(abc,acc,alc)
                                count_x += 1

                
                jitai1.append(tube_name_j)









        main_w.save(r'C:\Users\operator\Desktop\严旭（勿动）\数据在此'+'\\'+"石墨舟次数统计"+str(name_name[0])+'.xlsx')  
        print(jitai,' is ok')
    print('print enter to exit!')    
                

        
            
        #print(jitai1)
















get_tube()

