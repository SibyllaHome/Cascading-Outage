def Dynamic_Load_Shedding_Calculator_IEEE39(faulty_line, faulty_node_overfreq, faulty_node_underfreq, target_load, amount_of_shedding) :
    import matlab
    import matlab.engine
    eng = matlab.engine.start_matlab()
    a = list(map(int,faulty_line))
    b = list(map(int,faulty_node_overfreq))
    c = list(map(int,faulty_node_underfreq))
    d = list(map(int,target_load))
    e = amount_of_shedding
    # Need to finish a m file utilizing MATPOWER - finish when move to 39-bus system
    LS = eng.Dynamic_Load_Shedding_Calculator(matlab.int32(a),matlab.int32(b + c),matlab.int32(d),matlab.double(e)) 
    # print(LS)
    return a, b, c, LS[0][0], LS[0][1], LS[0][2], LS[0][3], LS[0][4]

# # test
# print(Dynamic_Load_Shedding_Calculator_IEEE39(['21', '20'], ['31'], ['32'], [3, 4, 7, 8, 12, 15, 16, 18, 20, 21, 23, 24, 25, 26, 27, 28, 29, 31, 39],[0, 500.0, 0, 522.0, 0, 0, 0, 0, 0, 0, 0, 0, 224.0, 139.0, 281.0, 206.0, 283.5, 0, 1104.0]))