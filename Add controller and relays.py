import os
os.environ["PATH"] = r'C:\\Program Files\\DIgSILENT\\PowerFactory 2021 SP5'+ os.environ["PATH"]
import sys
sys.path.append(r'C:\\Program Files\\DIgSILENT\\PowerFactory 2021 SP5\\Python\\3.7')
import powerfactory as pf

app=pf.GetApplication()
if app is None:
    raise Exception('getting Powerfactory application failed')
else:
    print('------ Start Simulation -------')
# app.Show()
app.ActivateProject('Base')



'''
ChangeSymType: change the RMS/EMS model type of TypSym(Type of SYM)
Input:
type - the RMS/EMS model type
'''
def ChangeSymType(type):
    # SelectedGens = []
    sMachineTypes = app.GetCalcRelevantObjects("*.TypSym")
    for item in sMachineTypes:
        if type == 'standard' : 
            item.model_inp='det'
        if type == 'classical' : 
            item.model_inp='cls'

# # test
# ChangeSymType(type = 'standard')


'''
ChangeLoadType: change the Voltage depedence coefficient in 'Load Flow' 
Input: 
ap - coefficient aP
bp - coefficient bP
bq - coefficient bQ
'''
def ChangeLoadType(ap, bp, bq) :
    LoadTypes = app.GetCalcRelevantObjects("*.TypLod")
    for item in LoadTypes:
        # app.PrintPlain(item)
        item.aP = ap
        item.bP = bp
        item.bQ = bq

# # test
# ChangeLoadType(ap = 0, bp = 0, bq = 0)


'''
DisableAVR: disable the corresponding controller
Input:
name - the name of controller that user want to disable
'''
def Disable_Controller(name) :
    CTL = app.GetCalcRelevantObjects("*.ElmDsl")
    for item in CTL:
        if item.loc_name.find(name) >= 0:
            # app.PrintPlain(item)
            item.outserv = 1

# # test
# Disable_Controller(name = 'IEEET1')


'''
AddOvercurrentRelay: add overcurrent relay to selected lines and transformer
Input:
items - the list of items where the user want to add overcurrent relay
relay_type_name - the name of overcurrent relay type
'''
def AddOvercurrentRelay( items = None):
    SelectedLines = []
    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    Transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    if items == None:
        SelectedLines = Lines
    else:
        for item in items:
            for line in Lines:
                if line.loc_name == item:
                    SelectedLines.append(line)
                    break
    # app.PrintPlain(SelectedLines)
    # app.PrintPlain(len(SelectedLines))

    # Get the folder of relay types
    RelayFolder = app.GetLocalLibrary("TypRelay")
    OvercurrentRelay = RelayFolder.GetContents('F50_F51 Phase overcurrent.TypRelay')[0]
    CTFolder = app.GetLocalLibrary("TypCT")
    CT = CTFolder.GetContents('*.TypCt')
    app.PrintPlain(OvercurrentRelay)

    import Excel
    data = Excel.read_excel('Base_Parameters.xlsx', 'Line Parameters', 4, 6, 1, 15, 1.0)

    for line in SelectedLines:
        cub1 = line.bus1
        cub2 = line.bus2
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        # app.PrintPlain(ExistRelay)
        for relay in ExistRelay:
            if relay.loc_name.find('Overcurrent Relay') >= 0:
                app.PrintPlain([line, r'Overcurrent Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', line.loc_name + ' Overcurrent Relay')
        relay.typ_id = OvercurrentRelay
        # Create current trasformer and core current transformer
        ct = relay.CreateObject('StaCt', 'CT')
        corect = relay.CreateObject('StaCt', 'CoreCT')
        ct.typ_id = CT[0]
        # app.PrintPlain(ct)
        ct.ptapset = 1000
        corect.typ_id = CT[0]
        corect.ptapset = 1000
        relay.slotupd()
        # app.PrintPlain(data)

        I1 = relay.GetSlot('I>')
        I2 = relay.GetSlot('I>>')
        I3 = relay.GetSlot('I>>>')
        I4 = relay.GetSlot('I>>>>')
        for i in range(data.shape[1]):
            if line.loc_name == data[0, i]:
                # app.PrintPlain('Find Line')
                I1.Ipset = float(data[7, i]) 
                I1.Tpset = float(data[8, i])
                I2.Ipset = float(data[9, i])
                I2.Tset = float(data[10, i])
                I3.Ipset = float(data[11, i])
                I3.Tset = float(data[12, i])
                I4.outserv = 1
                break
        
        
        Logics = app.GetCalcRelevantObjects("*.RelLogdip")
        for item in Logics:
            if item.fold_id.loc_name.find('Overcurrent Relay') >= 0:
                if item.pSwitch == [None]:
                    item.pSwitch=[cub1, cub2]
        app.PrintPlain([line, r'Overcurrent Relay Installed'])

    for Trf in Transformers:
        cub1 = Trf.bushv
        cub2 = Trf.buslv
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name.find('Overcurrent TF Relay') >= 0:
                app.PrintPlain([Trf, r'Overcurrent TF Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', Trf.loc_name +' Overcurrent Relay')
        relay.typ_id = OvercurrentRelay
        ct = relay.CreateObject('StaCt', 'CT')
        corect = relay.CreateObject('StaCt', 'CoreCT')
        ct.typ_id = CT[0]
        ct.ptapset = 1000
        corect.typ_id = CT[0]
        corect.ptapset = 1000
        relay.slotupd()


        I1 = relay.GetSlot('I>')
        I2 = relay.GetSlot('I>>')
        I3 = relay.GetSlot('I>>>')
        I4 = relay.GetSlot('I>>>>')
        for i in range(data.shape[1]):
            if Trf.loc_name == data[0, i]:
                # app.PrintPlain('Find Line')
                I1.Ipset = float(data[7, i]) 
                I1.Tpset = float(data[8, i])
                I2.Ipset = float(data[9, i])
                I2.Tset = float(data[10, i])
                I3.Ipset = float(data[11, i])
                I3.Tset = float(data[12, i])
                I4.outserv = 1
                break
        
        
        Logics = app.GetCalcRelevantObjects("*.RelLogdip")
        for item in Logics:
            if item.fold_id.loc_name.find('Overcurrent TF Relay') >= 0:
                if item.pSwitch == [None]:
                    item.pSwitch=[cub1, cub2]
        app.PrintPlain([Trf, r'Overcurrent TF Relay Installed'])

# # test
# AddOvercurrentRelay()


'''
RemoveOvercurrentRelay: Remove overcurrent relay from selected lines and transformer
Input:
item - list of selected item(line)
'''
def RemoveOvercurrentRelay(items = None):
    SelectedLines = []
    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    Transformers = app.GetCalcRelevantObjects("*.ElmTr2")
    if items == None:
        SelectedLines = Lines
    else:
        for item in items:
            for line in Lines:
                if line.loc_name == item:
                    SelectedLines.append(line)
                    break
    # app.PrintPlain(SelectedLines)
    # app.PrintPlain(len(SelectedLines))
    for line in SelectedLines:
        cub1 = line.bus1
        cub2 = line.bus2
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name.find('Overcurrent Relay') >= 0:
                app.PrintPlain([line, r'Overcurrent Relay Removed'])
                item.Delete()


    for Trf in Transformers:
        cub1 = Trf.bushv
        cub2 = Trf.buslv
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name.find('Overcurrent TF Relay') >= 0:
                app.PrintPlain([Trf, r'Overcurrent TF Relay Removed'])
                item.Delete()

# # test
# RemoveOvercurrentRelay()
# AddOvercurrentRelay()


'''
AddUnderFrequencyLoadShedding: Add Underfrequency Load Shedding relay to items
Input:
items - the list of items that need to add relay
'''
def AddUnderFrequencyLoadShedding(items = None):
    SelectedLoads = []
    Loads = app.GetCalcRelevantObjects("*.ElmLod")
    
    if items == None:
        SelectedLoads = Loads
    else:
        for item in items:
            for load in Loads:
                if load.loc_name == item:
                    SelectedLoads.append(load)
                    break
    # app.PrintPlain(SelectedLoads)
    # app.PrintPlain(len(SelectedLoads))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    UnderFrequencyLoadShedding = RelayFolder.GetContents('UFLS Relay.TypRelay')[0]
    VTFolder = app.GetLocalLibrary('TypVt')
    VT = VTFolder.GetContents('*.TypVt')

    for load in SelectedLoads:
        cub1 = load.bus1
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name.find('UFLS Relay') >= 0:
                app.PrintPlain([load, r'UFLS Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', load.loc_name + ' UFLS Relay')
        relay.typ_id = UnderFrequencyLoadShedding
        vt = relay.CreateObject('StaCombi','VT')
        vt.typ_vt = VT[0]
        vt.ptapser = 20000
        # app.PrintPlain(vt)
        relay.slotupd()

        F1 = relay.GetSlot('F<')
        F1.Ipsetr = 59.1
        F1.Tpset = 0.2333

        Logics = app.GetCalcRelevantObjects("*.RelLslogic")
        # app.PrintPlain(item.fold_id)
        for item in Logics:
            if item.fold_id.loc_name.find('UFLS Relay') >= 0 and item.pLoad == [None]:
                    # app.PrintPlain('asd')
                    item.pLoad = [load]

        # PLLs = app.GetCalcRelevantObjects("*.ElmPhi__pll")
        # app.PrintPlain(PLLs)
        # for item in PLLs:
        #     if item.fold_id.loc_name == 'UFLS Relay' and item.pbusbar == None:
        #         item.pbusbar = load.bus1
        #         item.mversion = 2

# # test
# AddUnderFrequencyLoadShedding(['Load2 Shedding','Load3 Shedding'])


'''
RemoveUnderFrequencyLoadShedding: Remove Underfrequency Load Shedding relay to items
Input:
items - the list of items that need to add relay
'''
def RemoveUnderFrequencyLoadShedding(items = None):
    SelectedLoads = []
    Loads = app.GetCalcRelevantObjects("*.ElmLod")
    if items == None:
        SelectedLoads = Loads
    else:
        for item in items:
            for load in Loads:
                if load.loc_name == item:
                    SelectedLoads.append(load)
                    break
    # app.PrintPlain(SelectedLoads)
    # app.PrintPlain(len(SelectedLoads))
    for load in SelectedLoads:
        cub1 = load.bus1
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name.find('UFLS Relay') >= 0:
                app.PrintPlain([load, r'UFLS Relay Removed'])
                item.Delete()

# # test
# RemoveUnderFrequencyLoadShedding(['Load2 Shedding','Load3 Shedding'])
# AddUnderFrequencyLoadShedding(['Load2 Shedding','Load3 Shedding'])


'''
AddOverFrequencyGeneratortripng: Add Overfrequency Generaor tripping relay to items
Input:
items - the list of items that need to add relay
'''
def AddOverFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    # app.PrintPlain(SelectedGens)
    # app.PrintPlain(len(SelectedGens))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    OverFrequencyGeneratorTripping = RelayFolder.GetContents('OFGT Relay.TypRelay')[0]
    VTFolder = app.GetLocalLibrary('TypVt')
    VT = VTFolder.GetContents('*.TypVt')
    # app.PrintPlain(TypRelays)

    for Gen in SelectedGens:
        cub1 = Gen.bus1
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name.find('OFGT Relay') >= 0:
                app.PrintPlain([Gen, r'OFGT Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', Gen.loc_name + ' OFGT Relay')
        relay.typ_id = OverFrequencyGeneratorTripping
        vt = relay.CreateObject('StaVt','VT')
        vt.typ_id = VT[0]
        vt.ptapser = 20000
        # app.PrintPlain(vt)
        relay.slotupd()

        F1 = relay.GetSlot('F>1')
        F2 = relay.GetSlot('F>2')
        F3 = relay.GetSlot('F>3')
        F4 = relay.GetSlot('F>4')
        F5 = relay.GetSlot('F>5')
        
        # app.PrintPlain('Find Line')
        F1.Ipsetr = 60.6
        F1.Tpset = 180
        F2.Ipsetr = 61.6
        F2.Tpset = 30
        F3.Ipsetr = 61.7
        F3.Tpset = 0.1
        F4.outserv = 1
        F5.outserv = 1

        Logics = app.GetCalcRelevantObjects("*.RelLogdip")
        # app.PrintPlain(item.fold_id)
        for item in Logics:
            if item.fold_id.loc_name.find('OFGT Relay') >= 0:                
                if item.pSwitch == [None]:
                    item.pSwitch = [cub1]

# # test
# AddOverFrequencyGeneratorTripping()


'''
RemoveOverFrequencyGeneratortripng: Remove Overfrequency Generaor tripping relay to items
Input:
items - the list of items that need to remove relay
'''
def RemoveOverFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    for Gen in SelectedGens:
        cub1 = Gen.bus1
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name.find('OFGT Relay') >= 0:
                app.PrintPlain([Gen, r'OFGT Relay Removed'])
                item.Delete()

# # test
# RemoveOverFrequencyGeneratorTripping()
# AddOverFrequencyGeneratorTripping()


'''
AddUnderFrequencyGeneratortripng: Add Underfrequency Generaor tripping relay to items
Input:
items - the list of items that need to add relay
'''
def AddUnderFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    # app.PrintPlain(SelectedGens)
    # app.PrintPlain(len(SelectedGens))
    RelayFolder = app.GetLocalLibrary("TypRelay")
    UnderFrequencyGeneratorTripping = RelayFolder.GetContents('UFGT Relay.TypRelay')[0]
    VTFolder = app.GetLocalLibrary('TypVt')
    VT = VTFolder.GetContents('*.TypVt')
    # app.PrintPlain(TypRelays)

    for Gen in SelectedGens:
        cub1 = Gen.bus1
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for relay in ExistRelay:
            if relay.loc_name.find('UFGT Relay') >= 0:
                app.PrintPlain([Gen, r'UFGT Relay Exists'])
                relay.Delete()
        relay = cub1.CreateObject('ElmRelay', Gen.loc_name + ' UFGT Relay')
        relay.typ_id = UnderFrequencyGeneratorTripping
        vt = relay.CreateObject('StaVt','VT')
        vt.typ_id = VT[0]
        vt.ptapser = 20000
        # app.PrintPlain(vt)
        relay.slotupd()

        F1 = relay.GetSlot('F<1')
        F2 = relay.GetSlot('F<2')
        F3 = relay.GetSlot('F<3')
        F4 = relay.GetSlot('F<4')
        F5 = relay.GetSlot('F<5')
        
        # app.PrintPlain('Find Line')
        F1.Ipsetr = 59.4
        F1.Tpset = 180
        F2.Ipsetr = 58.4
        F2.Tpset = 30
        F3.Ipsetr = 57.8
        F3.Tpset = 7.5
        F4.Ipsetr = 57.3
        F4.Tpset = 0.75
        F5.Ipsetr = 57
        F5.Tpset = 0

        Logics = app.GetCalcRelevantObjects("*.RelLogdip")
        # app.PrintPlain(item.fold_id)
        for item in Logics:
            if item.fold_id.loc_name.find('UFGT Relay') >= 0:                
                if item.pSwitch == [None]:
                    item.pSwitch = [cub1]

# # test
# AddOverFrequencyGeneratorTripping()


'''
RemoveUnderFrequencyGeneratortripng: Remove Overfrequency Generaor tripping relay to items
Input:
items - the list of items that need to remove relay
'''
def RemoveUnderFrequencyGeneratorTripping(items = None):
    SelectedGens = []
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    if items == None:
        SelectedGens = SymMachines
    else:
        for item in items:
            for SymMachine in SymMachines:
                if SymMachine.loc_name == item:
                    SelectedGens.append(SymMachine)
                    break
    for Gen in SelectedGens:
        cub1 = Gen.bus1
        # switch1 = cub1.GetContents('*.StaSwitch')
        ExistRelay = cub1.GetContents('*.ElmRelay')
        for item in ExistRelay: 
            if item.loc_name.find('UFGT Relay') >= 0:
                app.PrintPlain([Gen, r'UFGT Relay Removed'])
                item.Delete()

# # test
# RemoveUnderFrequencyGeneratorTripping()
# AddUnderFrequencyGeneratorTripping()


'''
SetUpGeneratorControl: Setup the controller for generator
'''
def SetUpGeneratorControl():
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    UserDefinedModels = app.GetProjectFolder('blk')
    GenControl = UserDefinedModels.GetContents('SYM Frame_no droop.BlkDef')[0]
    AVR = UserDefinedModels.GetContents('avr_IEEET1.BlkDef')[0]
    GOV = UserDefinedModels.GetContents('gov_TGOV1.BlkDef')[0]
    # app.PrintPlain(UserDefinedModels)
    # app.PrintPlain(GenControl)

    for SymMachine in SymMachines:
        app.PrintPlain(SymMachine)
        CompModel = SymMachine.c_pmod # plant model of sym
        app.PrintPlain(CompModel)
        CompModel.typ_id = GenControl
        ExistDsl = CompModel.GetContents('*.ElmDsl')
        for item in ExistDsl:
            if item.loc_name.find('IEEET1') >= 0 :
                app.PrintPlain([item, r'AVR Exists'])
                item.Delete()
            if item.loc_name.find('TGOV1') >= 0 :
                app.PrintPlain([item, r'AVR Exists'])
                item.Delete()

        # ExistVmea = CompModel.GetContents('*.StaVmea')
        # for item in ExistVmea:
        #     if item.loc_name.find('Vmea') >= 0 :
        #         app.PrintPlain([item, r'Vmea Exists'])
        #         item.Delete()

        AVRDsl = CompModel.CreateObject('ElmDsl', SymMachine.loc_name + ' IEEET1')
        AVRDsl.typ_id = AVR
        AVRDsl.params = [0.06, 40., 0.1, -0.06, 0.67, 0.12, 1., 3., 0.09, 4., 0.368, -1., 1.]
        GOVDsl = CompModel.CreateObject('ElmDsl', SymMachine.loc_name +' TGOV1')
        GOVDsl.typ_id = GOV
        GOVDsl.params = [10., 0.2, 1., 0., 0.05, 0.1, 0., 0., 1.]
        CompModel.SlotUpdate()
        # Vmea1 = CompModel.CreateObject('StaVmea', 'Vmea1')
        # Vmea1.pbusbar = SymMachine.bus1
        # Vmea2 = CompModel.CreateObject('StaVmea', 'Vmea2')
        # Vmea2.pbusbar = SymMachine.bus1

        CompModel.SlotUpdate()

        pelm = CompModel.pelm
        pelm[0] = SymMachine

        CompModel.SetAttribute('pelm', pelm)
        app.PrintPlain(CompModel.pelm)

# # test
# SetUpGeneratorControl()


'''
ChangeLoadingLevel: Change the level of load
Input: 
loading - the factor of load setting
'''
def ChangeLoadingLevel(loading = 1.0):
    import Excel
    
    Loads = app.GetCalcRelevantObjects("*.ElmLod")
    # Ldf = app.GetFromStudyCase('ComLdf')
    data = Excel.read_excel('Base_Parameters.xlsx', 'Load Parameters', 3, 7, 1, 3, loading)
    app.PrintPlain(loading)
    for Load in Loads:
        for i in range(data.shape[1]):
            if Load.loc_name == data[0, i]:
                Load.plini = float(data[1, i])
                Load.qlini = float(data[2, i])
                break

# # test
# ChangeLoadingLevel()


'''
ChangeInitialGenerationLevel: Change the level of initial generation
Input: 
loading - the factor of initial generation setting
'''
def ChangeInitialGenerationLevel(loading = 1.0):
    import Excel
    SymMachines = app.GetCalcRelevantObjects("*.ElmSym")
    data = Excel.read_excel('Base_Parameters.xlsx', 'Pmin=0 reserve=60', 4, 5, 1, 7, 1.0)
    # Ldf = app.GetFromStudyCase('ComLdf') # Get load flow calculation case
    
    for item in SymMachines:
        for i in range(data.shape[1]):
            if item.loc_name == data[0, i]:
                # Check generator's work status
                # app.PrintPlain(item.loc_name)
                if float(data[3, i]) == 1.0 :
                    item.outserv = 0
                    # app.PrintPlain(item.loc_name)
                else:
                    item.outserv = 1

                item.pgini = float(data[4, i])
                item.qgini = float(data[5, i])
                # app.PrintPlain(float(data[6, i]))
                item.usetp = float(data[6, i])
                item.Pmin_uc = float(data[2, i]) * loading
                # item.pgini = float(data[2, i]) * 1.0
                # item.qgini = float(data[3, i]) * 1.0
                # Load.qlini = 0.0
                break
    # Ldf.Execute()

# # test
# ChangeInitialGenerationLevel()


'''
ChangeLineRating: Change the rating of lines
'''
def ChangeLineRating():
    
    import Excel
    Lines = app.GetCalcRelevantObjects("*.ElmLne")
    name = Excel.read_excel('Base_Parameters.xlsx', 'Line Parameters', 4, 6, 1, 1, 1.0)
    rating = Excel.read_excel('Base_Parameters.xlsx', 'Line Parameters', 4, 6, 9, 9, 1.0)

    for line in Lines:
        for i in range(name.shape[1]):
            if line.loc_name == name[0, i]:
                line.typ_id.sline = float(rating[0, i]) 
                break

# # test 
# ChangeLineRating()