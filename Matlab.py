def Dynamic_Load_Shedding_Calculator_Base(faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding) :
    import matlab
    import matlab.engine
    eng = matlab.engine.start_matlab()
    a = list(map(int,faulty_line))
    b = list(map(int,faulty_node_overfreq))
    c = list(map(int,faulty_node_underfreq))
    d = list(map(int,target_load))
    e = amount_of_shedding
    # Need to finish a m file utilizing MATPOWER - finish when move to 39-bus system
    LS = eng.Dynamic_Load_Shedding_Calculator1_Base(matlab.int32(a),matlab.int32(b),matlab.int32(c),matlab.int32(d),matlab.double(e),nargout = 5) 
    return a, b, LS[0], LS[1], LS[2], LS[3], LS[4]

# # test
# print(Dynamic_Load_Shedding_Calculator_Base(['3', '0', '3'] ,['1'],[],['450', '80'],[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]))