# with generator tripping
def events_recorder(input, name, ShowFullMessages):
    import csv
    import re
    import Excel
    import Matching
    import xlrd
    import numpy as np
    import Matlab

    p1 = re.compile(r'[=](.*?)[)]', re.S)
    p2 = re.compile(r"Grid split into (.*)", re.S)
    p3 = re.compile(r"Grid\\(.*).ElmLne", re.S)
    p4 = re.compile(r"Grid\\Line(.*)[-](.*).ElmLne", re.S)
    p5 = re.compile(r"Grid\\Bus (.*)\\Cub", re.S)
    p6 = re.compile(r"Element (.*) is local reference", re.S)
    p7 = re.compile(r"  1 (.*)", re.S)
    p8 = re.compile(r"Grid\\(.*).ElmSym", re.S)
    p9 = re.compile(r"Step: (.*)[)]", re.S)
    p10 = re.compile(r"\\Cub_(.*)\\", re.S)
    p11 = re.compile(r"G (.*)", re.S)


    # Global variables
    count = 0
    flag = 0
    flag2 = 4
    Cub1 = ""
    Comp1 = ""
    Name_of_Local_Reference2 = ""
    Switch_Event1 = ""
    Unsupplied_Areas1 = ""
    State_of_Logic1 = ""
    Targeted_Load2 = ""
    Trip_of_generator = ""
    Percent_of_Shedding1 = ""
    Amount_of_Load_Shedding = 0.0
    Amount_of_Load_Shedding1 = 0.0
    Amount_of_Load_Shedding2 = 0.0
    faulty_line = []
    faulty_node_overfreq = []
    faulty_node_underfreq = []
    tripped_line = []
    tripped_node_overfreq = []
    tripped_node_underfreq = []
    tripped_load = []
    tripped_LS_balance = []
    tripped_LS_tripped_gen = []
    tripped_LS_tripped_load = []
    tripped_LS_unsupplied = []
    failed_matrix = []
    faulty_matrix = []
    OutOfStep = []
    State_of_generator = ""
    amount_of_shedding= [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    j = 0
    i = 0

    # Active Power of Loads
    Loads = Excel.read_excel('IEEE39_Parameters.xlsx', 'Load Parameters', 4, 22, 1, 3, 1.0)
    target_load = Loads[1,:]

    # create empty matrix of 512*1
    data = [[''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], [''], ['']]

    with open(name, 'w', newline='') as csvfile:
        writer = csv.writer(csvfile)
        f = open(input, 'r')
        for line in f.readlines():
            temp = line.strip()
            print("Read:",temp)
            if temp.find('.ElmSym') >= 0:
                Comp = re.findall(p8, temp)
                Comp1 = "".join(Comp)
            elif temp.find('Generator out of step (pole slip)') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                data[i].extend([t1, Comp1, 'Out of step', '', ''])
                print(t1, Comp1, 'Out of step')
                i = i + 1
                Comp1 = ""
            elif temp.find('Outage Event: Element set to out of service') >= 0:
                Name_of_Local_Reference2 = ""
                Unsupplied_Areas1 = ""
                t = re.findall(p1, temp)
                t1 = "".join(t)
                Trip_of_generator = ""
                if Comp1 != "" :
                    if Comp1.find("(") == -1 :
                        gen = re.findall(p11, Comp1)
                        gen1 = "".join(gen)
                        name = Comp1 
                        State_of_generator = '>360 tripped'
                        Trip_of_generator = Comp1 + ' All tripped'
                        faulty_node_overfreq.append(Matching.MatchingGen_IEEE39(gen1))
                    else :
                        name = Comp1 
                        State_of_generator = '>360 tripped'
                        Trip_of_generator = ""
                    data[i].extend([t1, name, State_of_generator, Trip_of_generator, ''])
                    print(t1, name, State_of_generator, Trip_of_generator)
                    i = i + 1
                    Comp1 = ""
            elif temp.find('.ElmLne') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                Comp1 = ""
                if temp.find('evt  -') >= 0:
                    Line_Outage = re.findall(p3, temp)
                    Line_Outage1 = "".join(Line_Outage)
                    # print('Line Outage:',Line_Outage1)
                    Switch_Event = re.findall(p4, temp)
                    Switch_Event1 = " ".join(Switch_Event[0])
                    Switch_Event2 = Switch_Event1.split(' ')
                    # print('Line switch event:', Switch_Event1, Switch_Event2)
                    Line_name1 = Switch_Event1 + ' - ' + Switch_Event2
                    Line_name2 = Switch_Event2 + ' - ' + Switch_Event1
                    type, value, row = Matching.MatchingElement_IEEE39(name = Line_name1)
                    if value == '0':
                        type, value, row = Matching.MatchingElement_IEEE39(name = Line_name2)
                    faulty_line.append(row)
                    data[i].extend([t1, Line_Outage1, row, '', ''])
                    print(t1, Line_Outage1, row)
                    i = i + 1
            
            elif temp.find('.StaSwitch') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                Switch_Event = re.findall(p5, temp)
                Switch_Event1 = "".join(Switch_Event)
                Cub = re.findall(p10, temp)
                Cub1 = "".join(Cub)
                # print('Switch_Event & Cub:', Switch_Event1, Cub1)
                
                type, value, row = Matching.MatchingElement_IEEE39(bus = Switch_Event1, cub = Cub1)
                if type == 'Gen':
                    if flag2 == 1:
                        faulty_node_overfreq.append(row)
                    elif flag2 == 3:
                        faulty_node_underfreq.append(row)
                    Trip_of_generator = 'Gen' + value + ' tripped'
                    if ShowFullMessages:
                        data[i].extend([t1, 'Gen ' + value, Trip_of_generator, 'Open the breaker'])
                        print('Generator Outage:', t1, 'Gen ' + value, Trip_of_generator,'Open the breaker')
                    i = i + 1
                    Trip_of_generator = ''
                elif type == 'Load':
                    # faulty_node_underfreq.append(row)
                    Targeted_Load = 'Load ' + value
                    for j in range(Loads.shape[1]):
                        if Targeted_Load == Loads[1, j]:
                            Amount_of_Load_Shedding = float(Loads[2, j])
                            break
                    if ShowFullMessages:
                            data[i].extend([t1, Targeted_Load,  Percent_of_Shedding1, Amount_of_Load_Shedding, 'Open the breaker'])
                    print("Load Outage", t1, Targeted_Load,  Percent_of_Shedding1, Amount_of_Load_Shedding,'Open the breaker')
                    Targeted_Load = ""
                    Amount_of_Load_Shedding = 0.0
                    j = 100
                    i = i + 1
                elif type == 'Line' or 'Trf':
                    # faulty_line.append(value)
                    Line_Outage = "Line" + value
                    if ShowFullMessages:
                        data[i].extend([t1, Line_Outage, row, 'Open the breaker', ''])
                    print('Line Outage:',t1, Line_Outage, row, 'Open the breaker')
                    i = i + 1
            elif temp.find('local reference') >= 0:
                Name_of_Local_Reference = re.findall(p6, temp)
                Name_of_Local_Reference1 = "".join(Name_of_Local_Reference)
                Name_of_Local_Reference3 = Name_of_Local_Reference1 + Name_of_Local_Reference2
                Name_of_Local_Reference2 = Name_of_Local_Reference3
            elif temp.find('area(s) are unsupplied') >= 0:
                Unsupplied_Areas = re.findall(p7, temp)
                Unsupplied_Areas1 = "".join(Unsupplied_Areas)
            elif temp.find('Grid split') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                No_of_Islands = re.findall(p2, temp)
                No_of_Islands1 = "".join(No_of_Islands)
                if ShowFullMessages:
                    data[i].extend([t1, No_of_Islands1, Name_of_Local_Reference2, Unsupplied_Areas1, ''])
                    print(t1, No_of_Islands1, Name_of_Local_Reference2, Unsupplied_Areas1)
                    i = i + 1
            elif temp.find('Circuit-Breaker Action') >= 0 :
                Name_of_Local_Reference2 = ""
                Unsupplied_Areas1 = ""
            elif temp.find('.ElmRelay') >= 0:
                t = re.findall(p1, temp)
                t1 = "".join(t)
                Targeted_Load = re.findall(p5, temp)
                Targeted_Load1 = "".join(Targeted_Load)
                Targeted_Load2 = 'Load ' + Targeted_Load1
                for j in range(Loads.shape[1]):
                    if Targeted_Load2 == Loads[0, j]:
                        Amount_of_Load_Shedding = float(Loads[2, j])
                        break
                if temp.find('UFLS Relay.ElmRelay') >= 0:
                    flag2 = 0 # under frequency load shedding
                    Percent_of_Shedding = '-100%'
                    Percent_of_Shedding1 = "".join(Percent_of_Shedding)
                    amount_of_shedding[j] = Amount_of_Load_Shedding * 1.0
                    Amount_of_Load_Shedding3 = Amount_of_Load_Shedding2 + Amount_of_Load_Shedding1
                    Amount_of_Load_Shedding2 = Amount_of_Load_Shedding3
                    # print('Detect Load Outage',Targeted_Load2)
                elif temp.find('OFGT Relay.ElmRelay') >= 0:
                    flag2 = 1 #overfrequency generator tripping
                    Switch_Event = re.findall(p5, temp)
                    Switch_Event1 = "".join(Switch_Event)
                    Cub = re.findall(p10, temp)
                    Cub1 = "".join(Cub)
                    type, value, row = Matching.MatchingElement_IEEE39(bus = Switch_Event1,cub = Cub1)
                    Trip_of_generator = 'Gen ' + value + ' tripped'
                    # print('Switch_Event & Cub:', Switch_Event1, Cub1)
                            
                elif temp.find('OC') >= 0:
                    flag2 = 2 # Overcurrent Relay tripping
                    Switch_Event = re.findall(p5, temp)
                    Switch_Event2 = "".join(Switch_Event)
                    Cub = re.findall(p10, temp)
                    Cub2 = "".join(Cub)
                    type, value, row = Matching.MatchingElement_IEEE39(bus = Switch_Event2,cub = Cub2)
                    # print('Switch_Event & Cub:', Switch_Event2, Cub2)
                
                elif temp.find('UFGT Relay.ElmRelay') >= 0:
                    flag2 = 3 #overfrequency generator tripping
                    Switch_Event = re.findall(p5, temp)
                    Switch_Event1 = "".join(Switch_Event)
                    Cub = re.findall(p10, temp)
                    Cub1 = "".join(Cub)
                    type, value, row = Matching.MatchingElement_IEEE39(bus = Switch_Event1,cub = Cub1)
                    Trip_of_generator = 'Gen ' + value + ' tripped'
                    # print('Switch_Event & Cub:', Switch_Event1, Cub1)


            elif temp.find('Relay is tripping') >= 0:
                if flag2 == 0:
                    # faulty_node_underfreq.append(row)
                    if ShowFullMessages:
                            data[i].extend([t1, Targeted_Load2, 'UFLS Relay tripped', Percent_of_Shedding1, Amount_of_Load_Shedding1, ''])
                            i = i + 1
                    print("Load Outage", t1, Targeted_Load2, 'UFLS Relay tripped', Percent_of_Shedding1, Amount_of_Load_Shedding1)
                    Targeted_Load2 = ""
                    Amount_of_Load_Shedding = 0.0
                    j = 100
                    
                elif flag2 == 1:
                    # faulty_node_overfreq.append(row)
                    if ShowFullMessages:
                        data[i].extend([t1, 'Gen ' + value, 'OFGT Relay '+ Trip_of_generator, ''])
                        print('Generator Outage:', t1, 'Gen ' + value, 'OFGT Relay '+ Trip_of_generator)
                        i = i + 1
                        Trip_of_generator = ""

                elif flag2 == 2:
                    faulty_line.append(row)
                    data[i].extend([t1, "Line " + value,'Overcurrent Relay tripped', '', ''])
                    print('Line Outage:', t1, "Line " + value,'Overcurrent Relay tripped')
                    i = i + 1

                elif flag2 == 3:
                    # faulty_node_underfreq.append(row)
                    if ShowFullMessages:
                        data[i].extend([t1, 'Gen ' + value, 'UFGT Relay '+ Trip_of_generator, ''])
                        print('Generator Outage:', t1, 'Gen ' + value, 'UFGT Relay '+ Trip_of_generator)
                        i = i + 1
                        Trip_of_generator = ""


            elif temp.find('Simulation successfully executed.') >= 0 or temp.find('System-Matrix Inversion failed') >= 0 :
                print(faulty_line, faulty_node_overfreq, faulty_node_underfreq, amount_of_shedding)
                
                [a, b, c, LS, LS_balance, LS_tripped_gen, LS_tripped_load, LS_unsupplied] = Matlab.Dynamic_Load_Shedding_Calculator_IEEE39(faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding)
                data[i].extend([a, b, c, Amount_of_Load_Shedding2, LS])
                i = i + 1
                print("LS = %.10f" % LS)
                print("LS_balance = %.10f" % LS_balance)
                print("LS_tripped_gen = %.10f" % LS_tripped_gen)
                print("LS_tripped_load = %.10f" % LS_tripped_load)
                print("LS_unsupplied = %.10f" % LS_unsupplied)
                tripped_line.append(a)
                tripped_node_overfreq.append(b)
                tripped_node_underfreq.append(c)
                tripped_load.append(LS)
                tripped_LS_balance.append(LS_balance)
                tripped_LS_tripped_gen.append(LS_tripped_gen)
                tripped_LS_tripped_load.append(LS_tripped_load)
                tripped_LS_unsupplied.append(LS_unsupplied)
                if temp.find('System-Matrix Inversion failed') >= 0 :
                    data[i].extend([t1, 'System-Matrix Inversion failed', '', '', ''])
                    i = i + 1
                    faulty_matrix = 'System-Matrix Inversion failed'
                    failed_matrix.append(faulty_matrix)
                else:
                    faulty_matrix = ''
                    failed_matrix.append(faulty_matrix)
                faulty_line = []
                faulty_node_overfreq = []
                faulty_node_underfreq = []
                faulty_matrix = []
                LS = 0
                amount_of_shedding= [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
                Amount_of_Load_Shedding2 = 0
                if i <= 512:
                    for i in range(i, 512):
                        data[i].extend(['', '', '', '', ''])
                i = 0       
        
        #print(linematrix, nodeoverfreqmatrix, nodeunderfreqmatrix, loadmatrix)
        for i in range(512):
            d = data[i]
            writer.writerow(d)
   

#events_recorder('N-2_100loading_N-1_Unsecure.txt', 'N-2_100loading_N-1_Unsecure_Shortcut.csv', 0)
# events_recorder('1.txt', "1.csv", 1)
events_recorder('report.txt', "table.csv", 1)
# events_recorder('N-2_General.txt', "N-2_General.csv", 1)
