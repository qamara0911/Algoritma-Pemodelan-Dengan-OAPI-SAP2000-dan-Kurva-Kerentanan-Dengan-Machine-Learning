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

N_batang_bawah = 5

Jarak_kuda_kuda = 1000

Banyak_kudakuda = 10

Jarak_gording_max = 750 

for segmen in range (1200, 1600, 100) : 
    
    Panjang = segmen*N_batang_bawah

    for tephra in np.arange(0,1.1,0.1) :
        for Sudut in range(15,46,5):
            print('tephra = '+ str(tephra) +', sudut= '+str(Sudut)+', panjang= '+str(Panjang))
            Tinggi = float((math.tan(math.radians(Sudut))*(0.5*Panjang)))
            numberof_divided_area = float(math.sqrt(Tinggi**2+(0.5*Panjang)**2)/Jarak_gording_max)
            numberof_divided_area = round(numberof_divided_area)
            tephra =str(tephra)

            #set the following flag to True to execute on a remote computer
            
            Remote = False            

            #if the above flag is True, set the following variable to the hostname of the remote computer

            #remember that the remote computer must have SAP2000 installed and be running the CSiAPIService.exe
            
            RemoteComputer = "SpareComputer-DT"
            
            # set the following flag to True to attach to an existing instance of the program

            # otherwise a new instance of the program will be started

            AttachToInstance = True

            # set the following flag to True to manually specify the path to SAP2000.exe

            # this allows for a connection to a version of SAP2000 other than the latest installation

            # otherwise the latest installed version of SAP2000 will be launched

            SpecifyPath = False

            # if the above flag is set to True, specify the path to SAP2000 below

            ProgramPath = R"C:\Program Files\Computers and Structures\SAP2000 25\SAP2000.exe"
            # full path to the model

            # set it to the desired path of your model

            APIPath = R'f:/Data D/TUGAS TH.4/TGA/ATAP/SAP'+'/'+str(tephra)+'/'+str(Sudut)+'/'+str(Panjang)
            csvPath = R'f:/Data D/TUGAS TH.4/TGA/ATAP/CSV/'

            if not os.path.exists(APIPath):
                try:
                    os.makedirs(APIPath)
                except OSError:
                    pass       
            ModelPath = APIPath + os.sep + 'APIFile Atap sudut'+'beban tephra' + str(tephra) + 'kpa' + str(Sudut) + str(Panjang)

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
            SapModel=cSapModel(mySAPObject.SapModel)

            #initialize model
            SapModel.InitializeNewModel()

            #create new blank model
            File = cFile(SapModel.File)
            ret=File.NewBlank()

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

            #set weight dan mass (berat jenis(kg/m3), massa jenis(kg/m3))
            ret = PropMaterial.SetWeightAndMass('Baja Ringan', 7850, 800)
            ret = SapModel.SetPresentUnits(eUnits(N_mm_C))
            #set material isotropic property data (modulus elastisitas(MPa), angka poisson, koefisien ekspansi termal(C))
            ret = PropMaterial.SetMPIsotropic('Baja Ringan', 200000, 0.3, 0.00001170)
            #set other properties for cold formed materials (Fy(Mpa), Fu(MPa), kinematic)
            ret = PropMaterial.SetOColdFormed("Baja Ringan", 550, 550, 1 )

            #define meterial property
            #define material penutup atap
            ret=SapModel.SetPresentUnits(eUnits(kgf_m_C))
            #set material name
            ret=PropMaterial.SetMaterial('Material penutup atap', eMatType(MATERIAL_NODESIGN))

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
            #set nama, material, dimensi(tinggi, lebar, tebal, radius, tinggi bibir)
            ret = PropFrame.SetColdC("C75", 'Baja Ringan', 75, 34, 1.0, 1.5, 10)

            #define area section properties
            #set material
            ret = PropMaterial.AddQuick('Material penutup atap', eMatType(MATERIAL_NODESIGN))
            #set nama, material shell-thin, tebal
            ret = PropArea.SetShell("Penutup atap", 1, "Material penutup atap", 0, 0.1, 0.1)
            
            #set stiffness modification
            #modifikasi area section property/stiffness 
            ModValue = [0, 0, 0, 0, 0, 0, 1, 1, 1, 1]
            ret = PropArea.SetModifiers('Penutup atap', ModValue)

            #add frame object by coordinates
            #add coordinates
            nodes=[]
            nodes.append([0, 0, 0])
            nodes.append([(1/2*Panjang/N_batang_bawah), 0, Tinggi/N_batang_bawah])

            nodes.append([(2/2*Panjang/N_batang_bawah), 0, 0])
            nodes.append([(3/2*Panjang/N_batang_bawah), 0, (3*Tinggi/N_batang_bawah)])

            nodes.append([(4/2*Panjang/N_batang_bawah), 0, 0])
            nodes.append([(5/2*Panjang/N_batang_bawah), 0, (5*Tinggi/N_batang_bawah)])

            nodes.append([(6/2*Panjang/N_batang_bawah), 0, 0])
            nodes.append([(7/2*Panjang/N_batang_bawah), 0, (3*Tinggi/N_batang_bawah)])

            nodes.append([(8/2*Panjang/N_batang_bawah), 0, 0])
            nodes.append([(9/2*Panjang/N_batang_bawah), 0, (Tinggi/N_batang_bawah)])

            nodes.append([(10/2*Panjang/N_batang_bawah), 0, 0])
            nodes=np.array(nodes)
            
            PointObj = cPointObj(SapModel.PointObj)
            for i in range(len(nodes)):
                ret=PointObj.AddCartesian(nodes[i,0],nodes[i,1],nodes[i,2],str(i+1))

            #add frame
            bars=[]
            #Batang Bawah
            bars.append([1,3])
            bars.append([3,5])
            bars.append([5,7])
            bars.append([7,9])
            bars.append([9,11])

            #Batang Atas
            bars.append([1,2])
            bars.append([2,4])
            bars.append([4,6])
            bars.append([6,8])
            bars.append([8,10])
            bars.append([10,11])

            #Batang Tengah
            bars.append([2,3])
            bars.append([3,4])
            bars.append([4,5])
            bars.append([5,6])
            bars.append([6,7])
            bars.append([7,8])
            bars.append([8,9])
            bars.append([9,10])
            bars=np.array(bars)
            
            FrameObj = cFrameObj(SapModel.FrameObj)
            for i in range(len(bars)):
                ret=FrameObj.AddByPoint(str(bars[i,0]),str(bars[i,1]),'FrameName'+str(i+1),'C75')

            #tumpuan
            Restraint=[True, True, True, False, False, False]
            ret=PointObj.SetRestraint('1', Restraint)
            ret=PointObj.SetRestraint('11', Restraint)

            ii=[False, False, False, True, True, True]
            jj=[False, False, False, True, True, True]
            StartValue=[0,0,0,0,0,0]
            EndValue=[0,0,0,0,0,0]
            for i in range(len(bars)):
                ret=FrameObj.SetReleases(str(i+1), ii, jj, StartValue, EndValue)

            View = cView(SapModel.View)
            ret=View.RefreshView(0, False)

            #replicate kuda-kuda
            ret = SapModel.SelectObj.All(False)

            ObjectType=[1,2]
            EditGeneral = cEditGeneral(SapModel.EditGeneral)
            #sumbu replikasi, jarak, pergeseran awal, banyak replikasi, point object
            ret=EditGeneral.ReplicateLinear(0, Jarak_kuda_kuda, 0, Banyak_kudakuda-1, 1, '', ObjectType, False)

            #add area object kiri
            x = [0,(5/2*Panjang/N_batang_bawah),(5/2*Panjang/N_batang_bawah),0]
            y = [0,0,Jarak_kuda_kuda,Jarak_kuda_kuda]
            z = [0,(3/3*Tinggi),+(3/3*Tinggi),0]
            Name = '1'
            UserName = ''
            PropName = "Penutup atap"
            AreaObj = cAreaObj(SapModel.AreaObj)
            area = AreaObj.AddByCoord(4, x, y, z, Name, PropName, UserName)

            x = [(5/2*Panjang/N_batang_bawah),(10/2*Panjang/N_batang_bawah),(10/2*Panjang/N_batang_bawah),(5/2*Panjang/N_batang_bawah)]
            y = [0,0,Jarak_kuda_kuda,Jarak_kuda_kuda]
            z = [(3/3*Tinggi),0,0,(3/3*Tinggi)]
            Name = '2'
            UserName = ''
            PropName = "Penutup atap"
            area = AreaObj.AddByCoord(4, x, y, z, Name, PropName, UserName)

            #divide area kanan, 0 = object
            ret = AreaObj.SetSelected("1", True, eItemType(0))

            EditArea = cEditArea(SapModel.EditArea)
            ret = EditArea.Divide("1", 1, 1, '', numberof_divided_area, 1)

            ret = AreaObj.SetSelected("2", True, eItemType(0))
            ret = EditArea.Divide("2", 1, 1, '', numberof_divided_area, 1)

            #replicate penutup atap
            ret = SapModel.SelectObj.PropertyArea('Penutup atap', False)
            ObjectType=[5]
            ret=EditGeneral.ReplicateLinear(0, Jarak_kuda_kuda, 0, Banyak_kudakuda-2, 1, '', ObjectType, False)

            #beban
            kN_m_C = 6
            ret=SapModel.SetPresentUnits(eUnits(kN_m_C))
            LTYPE_DEAD = 1
            LTYPE_SUPERDEAD = 2
            LTYPE_LIVE = 3
            LTYPE_REDUCELIVE = 4
            LTYPE_QUAKE = 5
            LTYPE_WIND= 6
            LTYPE_SNOW = 7
            LTYPE_OTHER = 8

            LoadPatterns = cLoadPatterns(SapModel.LoadPatterns)
            ret = LoadPatterns.Add("DEAD", eLoadPatternType(1), 1, True)
            ret = LoadPatterns.Add("TEPHRA", eLoadPatternType(8), 0, True)
 
            #beban tephra 6 = Z direction, 2 = selected object
            ret = SapModel.SelectObj.PropertyArea("Penutup atap")
            ret = AreaObj.SetLoadUniform("Penutup atap","TEPHRA", float(tephra), 6, True, "Global",eItemType(2))

            #add combo
            DesignColdFormed = cDesignColdFormed(SapModel.DesignColdFormed)
            #matikan kombinasi otomatis, 0 = linear additive
            ret = DesignColdFormed.SetComboAutoGenerate(False)
            ret = SapModel.RespCombo.Add("COMB1", 0)

            #add load case to combo, 0 = load case
            ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType(0), "DEAD", 1.2)
            ret = SapModel.RespCombo.SetCaseList("COMB1", eCNameType(0), "TEPHRA", 1.0)

            DesignColdFormed = cDesignColdFormed(SapModel.DesignColdFormed)
            ret= DesignColdFormed.SetComboStrength("COMB1", True)

            Analyze = cAnalyze(SapModel.Analyze)
            ret = Analyze.SetRunCaseFlag("COMB1", True)

            #set alalysis options (Ux, Uy, Uz, Rx, Ry, Rz)
            ret = Analyze.SetActiveDOF([True,True,True,False,False,False])

            #save model
            File = cFile(SapModel.File)
            ret = File.Save(ModelPath)
            
            #set alalysis options
            ret = Analyze.SetActiveDOF([True,True,True,False,False,False])

            #run model (this will create the analysis model)
            ret = Analyze.RunAnalysis()

            #design
            
            ret = DesignColdFormed.SetCode("AISI-16")
            ret=DesignColdFormed.StartDesign()

            ret=SapModel.SetPresentUnits(eUnits(kN_mm_C))
            
            #check if combo is selected
            Results = cAnalysisResults(SapModel.Results)
            ret = Results.Setup.GetComboSelectedForOutput("COMB1", True)
            
            #set case and combo output selections
            # output cold form design to csv
            Results = cAnalysisResults(SapModel.Results)
            Setup = cAnalysisResultsSetup(Results.Setup)

            ret=Setup.DeselectAllCasesAndCombosForOutput()
            ret=Setup.SetComboSelectedForOutput('COMB1', True)
    
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

                for i in range(SapModel.FrameObj.Count()) :
                    [ret, NumberItems, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary] = DesignColdFormed.GetSummaryResults(str(i+1), NumberItems, FrameName, Ratio, RatioType, Location, ComboName, ErrorSummary, WarningSummary)
                    [ret,PropName,SAuto] = FrameObj.GetSection(str(i+1), PropName,SAuto)

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
                data.to_csv((csvPath+'Cold Formed Summary Result '+'tephra '+str(tephra)+' '+str(Sudut)+' '+str(Panjang)), index=False)

            export_Cold_Formed_Summary_Data(csvPath)

            ret = File.Save()

print('Done')

