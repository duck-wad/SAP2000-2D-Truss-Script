import os
import sys
import comtypes.client

''' INITIATING THE MODEL '''

# flag to True to attach to existing instance of program
# false, new instance is started
AttachToInstance = False

# flag to True to specify path to specific version of sap
# false, newest installation
SpecifyPath = False
ProgramPath = ''

# path to model
APIPath = 'C:\\Users\\Nick\\source\\repos\\SAP2000 API\\Introductory Example'

if not os.path.exists(APIPath):
    try:
        os.makedirs(APIPath)
    except OSError:
        pass

ModelPath = APIPath + os.sep + 'API_1-001.sdb'

# create API helper object
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

if AttachToInstance:
    # attach to running instance of SAP2000
    try:
        mySapObject = helper.GetObject("CSI.SAP2000.API.SapObject")
    except (OSError, comtypes.COMError):
        print("No running instance of program found or failed to attach")
        sys.exit(-1)
else:
    if SpecifyPath:
        try:
            mySapObject = helper.CreateObject(ProgramPath)
        except (OSError, comtypes.COMError):
            print("Cannot start new instance of program")
            sys.exit(-1)
    else:
        try:
            mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")
        except (OSError, comtypes.COMError):
            print("Cannot start new instance of program")
            sys.exit(-1)
    # start SAP2000 application
    mySapObject.ApplicationStart()

#create sapmodel object
SapModel = mySapObject.SapModel

# initialize model
SapModel.InitializeNewModel()

# create new blank model
ret = SapModel.File.NewBlank()

''' DEFINING MODEL PROPERTIES '''

# define concrete material
MATERIAL_CONCRETE = 2
# (name, mattype)
# 1=steel, 2=concrete, 3=nodesign, 4=aluminum, 5=coldform, 6=rebar, 7=tendon
# returns 0 if model is successful, otherwise nonzero
# ret should remain 0 if everything runs properly throughout model creation
ret = SapModel.PropMaterial.SetMaterial('CONC', MATERIAL_CONCRETE)

# assign isotropic mechanical properties
# (name, elastic modulus, poisson ratio, thermal coefficient)
ret = SapModel.PropMaterial.SetMPIsotropic('CONC', 3600, 0.2, 0.0000055)

# define rectangular frame section property
# (name, material, section depth, section width)
ret = SapModel.PropFrame.SetRectangle('R1', 'CONC', 12, 12)

# define frame section property modifiers
ModValue = [1000, 0, 0, 1, 1, 1, 1, 1]
# array of eight unitless modifiers
# [cross sectional area, shear area in local 2 direction, shear area in local 3 direction, 
# torsional constant, MOI about local 2, MOI about local 3, mass, weight]
ret = SapModel.PropFrame.SetModifiers('R1', ModValue)

# default units are kip, inches, and degrees F
# change units using SetPresentUnits method on the SapModel object
# switch to k-ft units
# for kN, m, C, do 6
UNITS = 4
ret = SapModel.SetPresentUnits(UNITS)

''' DEFINE FRAME COORDINATES '''

# add frame object by coordinates
FrameName1 = ''
FrameName2 = 'peepee'
FrameName3 = 'pooopoo'
# (xyz of I-end, xyz of J-end, name, property name, optional user specified name, coordinate system)
FrameName1, ret = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, FrameName1, 'R1', '1', 'Global')
FrameName2, ret = SapModel.FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, FrameName2, 'R1', '2', 'Global')
FrameName3, ret = SapModel.FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, FrameName3, 'R1', '3', 'Global')
print(FrameName1, FrameName2, FrameName3)

# assign point object restraint at base
PointName1 = ''
PointName2 = ''
Restraint = [True, True, True, True, False, False]
# retrieve name of point objects at each end of frame
# will return an integer ID
PointName1, PointName2, ret = SapModel.FrameObj.GetPoints(FrameName1, PointName1, PointName2)
print(PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName1, Restraint)
# assign point object restraint at top
Restraint = [True, True, False, False, False, False]
PointName1, PointName2, ret = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)
print(PointName1, PointName2)
ret = SapModel.PointObj.SetRestraint(PointName2, Restraint)

# refresh view
ret = SapModel.View.RefreshView(0, False)

''' ADD LOAD '''

# LTYPE = 8 is for other load
# LTYPE = 1 is dead
LTYPE = 8
# (load case name, type, self weight multiplier, True to add linear static load case)
ret = SapModel.LoadPatterns.Add('1', LTYPE, 1, True)
ret = SapModel.LoadPatterns.Add('2', LTYPE, 0, True)
ret = SapModel.LoadPatterns.Add('3', LTYPE, 0, True)

# assign load for pattern 2
PointName1, PointName2, ret = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
PointLoadValue = [0,0,-10,0,0,0]
# apply point loads to point objects, distributed loads to the frame object
ret = SapModel.PointObj.SetLoadForce(PointName1, '2', PointLoadValue)
# (name, load case, type (1 is force per unit length, 2 is moment per unit length), 
# integer indicating direction (10 is gravity dir), dist1, dist2, val1, val2)
ret = SapModel.FrameObj.SetLoadDistributed(FrameName3, '2', 1, 10, 0, 1, 1.8, 1.8)

#assign loading for load pattern 3
PointName1, PointName2, ret = SapModel.FrameObj.GetPoints(FrameName3, PointName1, PointName2)
PointLoadValue = [0,0,-17.2,0,-54.4,0]
ret = SapModel.PointObj.SetLoadForce(PointName2, '3', PointLoadValue)

# switch to k-in units
kip_in_F = 3
ret = SapModel.SetPresentUnits(kip_in_F)

''' RUN AND ANALYZE RESULTS '''
# save the model
ret = SapModel.File.Save(ModelPath)

# run model
ret = SapModel.Analyze.RunAnalysis()

# initialize for results
SapResult = [0,0,0,0,0,0,0]
PointName1, PointName2, ret = SapModel.FrameObj.GetPoints(FrameName2, PointName1, PointName2)

# get results for all load cases
for i in range(0, 7):
    NumberResults = 0
    Obj = []
    Elm = []
    ACase = []
    StepType = []
    StepNum = []
    U1 = []
    U2 = []
    U3 = []
    R1 = []
    R2 = []
    R3 = []
    ObjectElm = 0
    ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(str(i+1))
    if i <= 3:
        NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret = SapModel.Results.JointDispl(PointName2,
                                        ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
        SapResult[i] = U3[0]
    else:
        NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret = SapModel.Results.JointDispl(PointName1,
                                        ObjectElm, NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3)
        SapResult[i] = U1[0]

# close sap2000
ret = mySapObject.ApplicationExit(False)
SapModel = None
mySapObject = None
# fill independent results
IndResult = [0,0,0,0,0,0,0]
IndResult[0] = -0.02639
IndResult[1] = 0.06296
IndResult[2] = 0.06296
IndResult[3] = -0.2963
IndResult[4] = 0.3125
IndResult[5] = 0.11556
IndResult[6] = 0.00651

# calculate percentage diff
PercentDiff = [0,0,0,0,0,0,0]
for i in range (0,7):
    PercentDiff[i] = (SapResult[i] / IndResult[i]) - 1

# display results
for i in range(0,7):
    print()
    print(SapResult[i])
    print(IndResult[i])
    print(PercentDiff[i])