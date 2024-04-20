import os
import sys
import comtypes as ctypes
import math
import numpy as np
import pandas as pd
import comtypes.client
import random
import tabulate
import clr
from enum import Enum

clr.AddReference("System.Runtime.InteropServices")
from System.Runtime.InteropServices import Marshal
 
#set the following path to the installed SAP2000 program directory

clr.AddReference(R'C:\Program Files\Computers and Structures\SAP2000 25\SAP2000v1.dll')

from SAP2000v1 import *

#input parameter model
print ('Mulai SAP2000 API')

N_batang_bawah = 6

jarak_kuda_kuda = 1000

banyak_kuda_kuda = 10

jarak_gording_max = 750

for segment in range (1100, 1300, 100): 
    panjang = segment * N_batang_bawah
    for kec_angin_awal in range (1,50,1):
        for sudut in range(15,45,5):
            print('kec = ' + str(kec_angin_awal) + ', sudut= ' + str(sudut), ', panjang= ' + str(panjang))
            tinggi = float((math.tan(math.radians(sudut))*(0.5*panjang)))
            numberof_divided_area = float(math.sqrt(tinggi**2+(0.5*panjang)**2)/jarak_gording_max)
            numberof_divided_area = round(numberof_divided_area)
            kec_angin_mps = int(str(kec_angin_awal))
            kec_angin_mph = kec_angin_mps*2.237

            #set the following flag to True to execute on a remote computer
            
            Remote = False            

            #if the above flag is True, set the following variable to the hostname of the remote computer

            #remember that the remote computer must have SAP2000 installed and be running the CSiAPIService.exe
            
            RemoteComputer = "SpareComputer-DT"

            #set the following flag to True to attach to an existing instance of the program

            #otherwise a new instance of the program will be started

            AttachToInstance = True

            #set the following flag to True to manually specify the path to SAP2000.exe

            #this allows for a connection to a version of SAP2000 other than the latest installation

            #otherwise the latest installed version of SAP2000 will be launched

            SpecifyPath = False

            #if the above flag is set to True, specify the path to SAP2000 below

            ProgramPath = R"C:\Program Files\Computers and Structures\SAP2000 25\SAP2000.exe"

            #full path to the model

            #set it to the desired path of your model

            #if executing remotely, ensure that this folder already exists on the remote computer

            #the below command will only create the folder locally

            APIPath = R'C:/Users/febyf/Documents/TGA FEBY/SAP HASIL RUNNING/DOUBLE HOWE/API/'+str(kec_angin_mps)+'/'+str(sudut)+'/'+str(panjang)
            csvPath = R'C:/Users/febyf/Documents/TGA FEBY/SAP HASIL RUNNING/DOUBLE HOWE/CSV/'

            if not os.path.exists(APIPath):
                try:
                    os.makedirs(APIPath)
                except OSError:
                    pass
            ModelPath = APIPath + os.sep + 'APIFile atap sudut' + str(sudut) +'kecepatan angin' + str(kec_angin_mps) + 'm per sekon' +  str(panjang)

            #create API helper object
            helper = cHelper(Helper())

            if AttachToInstance:
                #attach to a running instance of SAP2000
                try:
                    #get the active SAP2000 object       
                    if Remote:
                        mySAPObject = cOAPI(helper.GetObjectHost(RemoteComputer, "CSI.SAP2000.API.SAPObject"))
                    else:
                        mySAPObject = cOAPI(helper.GetObject("CSI.SAP2000.API.SAPObject"))
                except:
                    print("No running instance of the program found or failed to attach.")
                    sys.exit(-1)              

            else:
                if SpecifyPath:
                    try:
                        #'create an instance of the SAP2000 object from the specified path
                        if Remote:
                            mySAPObject = cOAPI(helper.CreateObjectHost(RemoteComputer, ProgramPath))
                        else:
                            mySAPObject = cOAPI(helper.CreateObject(ProgramPath))
                    except :
                        print("Cannot start a new instance of the program from " + ProgramPath)
                        sys.exit(-1)
                else:
                    try:
                        #create an instance of the SAP2000 object from the latest installed SAP2000
                        if Remote:
                            mySAPObject = cOAPI(helper.CreateObjectProgIDHost(RemoteComputer, "CSI.SAP2000.API.SAPObject"))
                        else:
                            mySAPObject = cOAPI(helper.CreateObjectProgID("CSI.SAP2000.API.SAPObject"))
                    except:
                        print("Cannot start a new instance of the program.")
                        sys.exit(-1)

                #start SAP2000 application
                mySAPObject.ApplicationStart()

            #create SapModel object
            SapModel = cSapModel(mySAPObject.SapModel)

            #initialize model
            SapModel.InitializeNewModel()

            #create new blank model
            File = cFile(SapModel.File)
            ret = File.NewBlank()

            #set unit id
            lb_in_F = 1
            lb_ft_F = 2
            kip_in_F = 3
            kip_ft_F = 4
            kN_mm_C = 5
            kN_m_C = 6
            kgf_mm_C = 7
            kgf_m_C = 8
            N_mm_C = 9
            N_m_C = 10
            Ton_mm_C = 11
            Ton_m_C = 12
            kN_cm_C = 13
            kgf_cm_C = 14
            N_cm_C = 15
            Ton_cm_C = 16

            #set material id
            MATERIAL_STEEL = 1
            MATERIAL_CONCRETE = 2
            MATERIAL_NODESIGN = 3
            MATERIAL_ALUMINUM = 4
            MATERIAL_COLDFORMED = 5
            MATERIAL_REBAR = 6
            MATERIAL_TENDON = 7

            #set present unit
            ret=SapModel.SetPresentUnits(eUnits(kgf_m_C))

            #define material property
            #define material rangka atap baja ringan
            #set material name
            PropMaterial = cPropMaterial(SapModel.PropMaterial)
            ret=PropMaterial.SetMaterial('Baja Ringan', eMatType(MATERIAL_COLDFORMED))
            #set weight dan mass (berat jenis, massa jenis)
            ret = PropMaterial.SetWeightAndMass('Baja Ringan', 7850, 800)
            ret = SapModel.SetPresentUnits(eUnits(N_mm_C))
            #set material isotropic property data (modulus elastisitas, angka poisson, koefisien ekspansi termal)
            ret = PropMaterial.SetMPIsotropic('Baja Ringan', 200000, 0.3, 0.00001170)
            #set other properties for cold formed materials (Fy, Fu)
            ret = PropMaterial.SetOColdFormed("Baja Ringan", 550, 550, 1 )

            #define meterial property
            #define material penutup atap
            ret = SapModel.SetPresentUnits(eUnits(kgf_m_C))
            #set material name
            ret = PropMaterial.SetMaterial('Material penutup atap', eMatType(MATERIAL_NODESIGN))
            #set weight dan mass (berat jenis, massa jenis)
            ret = PropMaterial.SetWeightAndMass("Material penutup atap", 7850, 800)
            ret=SapModel.SetPresentUnits(eUnits(N_mm_C))
            #set material isotropic property data (modulus elastisitas, angka poisson, koefisien ekspansi termal)
            ret=PropMaterial.SetMPIsotropic('Material penutup atap', 200000, 0.2, 0)

            #define frame section properties
            PropFrame = cPropFrame(SapModel.PropFrame)
            PropArea = cPropArea(SapModel.PropArea)
            #set material
            ret = PropMaterial.AddQuick('Baja ringan', eMatType(MATERIAL_COLDFORMED))
            #set nama, material, dimensi
            ret = PropFrame.SetColdC("C75", 'Baja Ringan', 75, 34, 1.0, 1.5, 10)

            #define area section properties
            #set material
            ret = PropMaterial.AddQuick('Material penutup atap', eMatType(MATERIAL_NODESIGN))
            #set nama, material
            ret = PropArea.SetShell("Penutup atap angin tekan", 1, "Material penutup atap", 0, 0.1, 0.1)
            #set nama, material
            ret = PropArea.SetShell("Penutup atap angin hisap", 1, "Material penutup atap", 0, 0.1, 0.1)
            #set stiffness modification
        
            #modifikasi area section property/stiffness 
            ModValue = [0, 0, 0, 0, 0, 0, 1, 1, 1, 1]
            ret = PropArea.SetModifiers('Penutup atap angin tekan', ModValue)
            ret = PropArea.SetShell("Penutup atap angin hisap", 1, "Material penutup atap", 0, 0.1, 0.1)
            
            #set stiffness modification
            #modifikasi area section property/stiffness 
            ModValue = [0, 0, 0, 0, 0, 0, 1, 1, 1, 1]
            ret = PropArea.SetModifiers('Penutup atap angin hisap', ModValue)
                
            #add frame object by coordinates
            #add coordinates
            nodes=[]
            nodes.append([0, 0, 0])
            nodes.append([(panjang/N_batang_bawah), 0, 0])
            nodes.append([(panjang/N_batang_bawah), 0, (2*tinggi/N_batang_bawah)])
            nodes.append([(2*panjang/N_batang_bawah), 0, 0])
            nodes.append([(2*panjang/N_batang_bawah), 0, (4*tinggi/N_batang_bawah)])
            nodes.append([(3*panjang/N_batang_bawah), 0, 0])
            nodes.append([(3*panjang/N_batang_bawah), 0, (6*tinggi/N_batang_bawah)])
            nodes.append([(4*panjang/N_batang_bawah), 0, 0])
            nodes.append([(4*panjang/N_batang_bawah), 0, (4*tinggi/N_batang_bawah)])
            nodes.append([(5*panjang/N_batang_bawah), 0, 0])
            nodes.append([(5*panjang/N_batang_bawah), 0, (2*tinggi/N_batang_bawah)])
            nodes.append([(6*panjang/N_batang_bawah), 0, 0])
            nodes=np.array(nodes)

            PointObj = cPointObj(SapModel.PointObj)
            for i in range(len(nodes)):
                ret = PointObj.AddCartesian(nodes[i,0],nodes[i,1],nodes[i,2],str(i+1))

            #add frame
            bars=[]
            #Batang Bawah
            bars.append([1,2])
            bars.append([2,4])
            bars.append([4,6])
            bars.append([6,8])
            bars.append([8,10])
            bars.append([10,12])

            #Batang Atas
            bars.append([1,3])
            bars.append([3,5])
            bars.append([5,7])
            bars.append([7,9])
            bars.append([9,11])
            bars.append([11,12])

            #Batang Tengah
            bars.append([2,3])
            bars.append([3,4])
            bars.append([4,5])
            bars.append([5,6])
            bars.append([6,7])
            bars.append([6,9])
            bars.append([8,9])
            bars.append([8,11])
            bars.append([11,10])
            bars=np.array(bars)

            FrameObj = cFrameObj(SapModel.FrameObj)
            for i in range(len(bars)):
                ret = FrameObj.AddByPoint(str(bars[i,0]),str(bars[i,1]),'FrameName'+str(i+1),'C75')

            #tumpuan
            Restraint=[True, True, True, False, False, False]
            ret = PointObj.SetRestraint('1', Restraint)
            ret = PointObj.SetRestraint('12', Restraint)

            ii=[False, False, False, True, True, True]
            jj=[False, False, False, True, True, True]
            StartValue=[0,0,0,0,0,0]
            EndValue=[0,0,0,0,0,0]
            for i in range(len(bars)):
                ret = FrameObj.SetReleases(str(i+1), ii, jj, StartValue, EndValue)

            #refresh view, update (initialize) zoom
            View = cView(SapModel.View)
            ret = View.RefreshView(0, False)

            #replicate kuda-kuda
            ret = SapModel.SelectObj.All(False)

            ObjectType = [1,2]
            EditGeneral = cEditGeneral(SapModel.EditGeneral)
            ret = EditGeneral.ReplicateLinear(0, jarak_kuda_kuda, 0, banyak_kuda_kuda-1, 1, '', ObjectType, False)

            #add area object
            x = [0,(3*panjang/N_batang_bawah),(3*panjang/N_batang_bawah),0]
            y = [0,0,jarak_kuda_kuda,jarak_kuda_kuda]
            z = [0,(tinggi),(tinggi),0]
            Name = '1'
            UserName = ''
            PropName = "Penutup atap angin tekan"
            AreaObj = cAreaObj(SapModel.AreaObj)
            area = AreaObj.AddByCoord(4, x, y, z, Name, PropName, UserName)

            x = [panjang,(3*panjang/N_batang_bawah),(3*panjang/N_batang_bawah),panjang]
            y = [0,0,jarak_kuda_kuda,jarak_kuda_kuda]
            z = [0,tinggi,tinggi,0]
            Name = '2'
            UserName = ''
            PropName = "Penutup atap angin hisap"
            area = AreaObj.AddByCoord(4, x, y, z, Name, PropName, UserName)

            #divide area
            ret = AreaObj.SetSelected("1", True, eItemType(0))

            EditArea = cEditArea(SapModel.EditArea)
            ret = EditArea.Divide("1", 1, 1, '', round(float(N_batang_bawah/2)), 1)

            ret = AreaObj.SetSelected("2", True, eItemType(0))
            ret = EditArea.Divide("2", 1, 1, '', round(float(N_batang_bawah/2)), 1)

            #replicate penutup atap
            ret = SapModel.SelectObj.PropertyArea('Penutup atap angin tekan', False)
            ObjectType=[5]
            ret = EditGeneral.ReplicateLinear(0, jarak_kuda_kuda, 0, banyak_kuda_kuda-2, 1, '', ObjectType, False)

            ret = SapModel.SelectObj.PropertyArea('Penutup atap angin hisap', False)
            ObjectType=[5]
            ret = EditGeneral.ReplicateLinear(0, jarak_kuda_kuda, 0, banyak_kuda_kuda-2, 1, '', ObjectType, False)

            #beban
            kN_m_C = 6
            ret = SapModel.SetPresentUnits(eUnits(kN_m_C))
            LTYPE_DEAD = 1
            LTYPE_SUPERDEAD = 2
            LTYPE_LIVE = 3
            LTYPE_REDUCELIVE = 4
            LTYPE_QUAKE = 5
            LTYPE_WIND= 6
            LTYPE_SNOW = 7
            LTYPE_OTHER = 8

            LoadPatterns = cLoadPatterns(SapModel.LoadPatterns)
            ret = LoadPatterns.Add("DEAD", eLoadPatternType(LTYPE_DEAD), 1, True)
            ret = LoadPatterns.Add("RAIN", eLoadPatternType(LTYPE_OTHER),0, True)
            ret = LoadPatterns.Add("WIND-TEKAN", eLoadPatternType(LTYPE_WIND),0, True)
            ret = LoadPatterns.Add("WIND-HISAP", eLoadPatternType(LTYPE_WIND),0, True)

            # Define the known net pressure coefficients for windward and leeward directions obstructed wind flow
            coefficients = {
                7.5: {
                    "windward": 1.1,
                    "leeward": -0.3
                },
                15: {
                    "windward": 1.1,
                    "leeward": -0.4
                },
                22.5: {
                    "windward": 1.1,
                    "leeward": 0.1
                },
                30: {
                    "windward": 1.3,
                    "leeward": 0.3
                },
                37.5: {
                    "windward": 1.3,
                    "leeward": 0.6
                },
                45: {
                    "windward": 1.1,
                    "leeward": 0.9
                },                                
            }

            # Function to get the net pressure coefficient for a given roof angle and wind direction
            def get_coefficient(sudut, windward_coeff, leeward_coeff):
                if sudut in coefficients:
                # If the angle is one of the known values, return the corresponding coefficient for the given direction
                    if windward_coeff and not leeward_coeff:
                        return coefficients[sudut]["windward"]
                    elif leeward_coeff and not windward_coeff:
                        return coefficients[sudut]["leeward"]
                    else:
                        raise ValueError("Unknown wind direction")
                else:
                # If the angle is not one of the known values, interpolate between the two nearest values for the given direction
                    angle_list = sorted(coefficients.keys())
                    i = angle_list.index(max(a for a in angle_list if a <= sudut))
                    angle1 = angle_list[i]
                    angle2 = angle_list[i+1]
                    coeff1 = coefficients[angle1]["windward"] if windward_coeff else coefficients[angle1]["leeward"]
                    coeff2 = coefficients[angle2]["windward"] if windward_coeff else coefficients[angle2]["leeward"]
                    return coeff1 + (sudut - angle1) * (coeff2 - coeff1) / (angle2 - angle1)
            
            windward_coeff = True
            leeward_coeff = False
            CNw = get_coefficient(sudut, windward_coeff, leeward_coeff)
                
            windward_coeff = False
            leeward_coeff = True
            CNl = get_coefficient(sudut, windward_coeff, leeward_coeff)

            ret = LoadPatterns.AutoWind.SetASCE710("WIND-TEKAN", 2, 0, 1, 1, 1, 1, 1, True, (8000+tinggi), 8000, (kec_angin_mph), 1, 1, 0.85, 0.85, 0.85)
            ret = LoadPatterns.AutoWind.SetASCE710("WIND-HISAP", 2, 0, 1, 1, 1, 1, 1, True, (8000+tinggi), 8000, (kec_angin_mph), 1, 1, 0.85, 0.85, 0.85)
                
            #kombinasi beban
            ret = SapModel.SelectObj.PropertyArea("Penutup atap angin tekan")
            ret = AreaObj.SetLoadUniform("Penutup atap angin tekan","RAIN", 0.063, 6, True, "Global",eItemType(2))
            ret = AreaObj.SetLoadWindPressure("Penutup atap angin tekan", "WIND-TEKAN", 1, CNw, eItemType(2))

            ret = SapModel.SelectObj.PropertyArea("Penutup atap angin hisap")
            ret = AreaObj.SetLoadUniform("Penutup atap angin hisap","RAIN", 0.063, 6, True, "Global",eItemType(2))
            ret = AreaObj.SetLoadWindPressure("Penutup atap angin hisap", "WIND-HISAP", 1, CNl, eItemType(2))

            #save model
            File = cFile(SapModel.File)
            ret = File.Save(ModelPath)
            
            #set alalysis options
            Analyze = cAnalyze(SapModel.Analyze)
            ret = Analyze.SetActiveDOF([True,True,True,False,False,False])

            #run model (this will create the analysis model)
            ret = Analyze.RunAnalysis()

            #design
            DesignColdFormed = cDesignColdFormed(SapModel.DesignColdFormed)

            ret = DesignColdFormed.SetCode("AISI-16")
            ret = DesignColdFormed.StartDesign()
            ret = SapModel.SetPresentUnits(eUnits(kN_mm_C))

            #set case and combo output selections
            # output cold form design to csv
            Results = cAnalysisResults(SapModel.Results)
            Setup = cAnalysisResultsSetup(Results.Setup)

            # output cold form design to csv
            ret = Setup.DeselectAllCasesAndCombosForOutput()
            ret = Setup.SetComboSelectedForOutput('DCLD2', True)
            ret = Setup.SetComboSelectedForOutput('DCLD3', True)
            ret = Setup.SetComboSelectedForOutput('DCLD4', True)
            ret = Setup.SetComboSelectedForOutput('DCLD5', True)
            ret = Setup.SetComboSelectedForOutput('DCLD6', True)
            ret = Setup.SetComboSelectedForOutput('DCLD7', True)
            ret = Setup.SetComboSelectedForOutput('DCLD8', True)
            ret = Setup.SetComboSelectedForOutput('DCLD9', True)

            def export_Cold_Formed_Summary_Data(csvPath):
                NumberItems = 2
                FrameName = []
                Ratio = []
                RatioType = []
                Location = []
                ComboName = []
                ErrorSummary = []
                WarningSummary = []
                SAuto = ''
                PropName = ''

                save_PropName = []
                save_framename = []
                save_ratio = []
                save_RatioType = []
                save_Location = []
                save_ComboName = []
                save_ErrorSummary = []
                save_WarningSummary = []
                save_elm = []

                for i in range(FrameObj.Count()) :
                    [ret, NumberItems, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary] = DesignColdFormed.GetSummaryResults(str(i+1), NumberItems, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary)
                    [ret, PropName,SAuto] = FrameObj.GetSection(str(i+1), PropName,SAuto)

                    save_PropName.append("".join(PropName))
                    save_framename.append("".join(FrameName))
                    save_ratio.append("".join(str(item) for item in Ratio if isinstance(item, (int, float))))
                    save_RatioType.append("".join(str(item) for item in RatioType if isinstance(item, (int, float))))
                    save_Location.append("".join(str(item) for item in Location if isinstance(item, (int, float))))
                    save_ComboName.append("".join(ComboName))
                    save_ErrorSummary.append("".join(ErrorSummary))
                    save_WarningSummary.append("".join(WarningSummary))
                
                
                data = {
                    'FrameName': save_framename
                    , 'Prop Name' : save_PropName
                    , 'Message' : ''
                    , 'Ratio': save_ratio
                    , 'Ratio Type' : save_RatioType
                    , 'Location' : save_Location
                    , 'Combo Name' : save_ComboName
                    , 'Error Summary' : save_ErrorSummary
                    , 'Warning Summary' : save_WarningSummary
                    }

                data = pd.DataFrame(data)

                for i in range(len(data)):
                    if float(data.Ratio[i]) >= 1:
                        data.loc[i, 'Message'] = "Overstress"
                    else:
                        data.loc[i, 'Message'] = "No Message"
                data.to_csv((csvPath+'Cold Formed Summary Result '+'angin '+str(kec_angin_mps)+' '+str(sudut)+' '+str(panjang)), index=False)

            export_Cold_Formed_Summary_Data(csvPath)
            
            ret = File.Save()

print('Done')