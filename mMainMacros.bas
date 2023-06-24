Attribute VB_Name = "mMainMacros"
'Macro_Main
Option Explicit
'-------------------------------------------------------------------------------------------------------------------
'----- CAM Automaion Macro Program Version
'-------------------------------------------------------------------------------------------------------------------
Public Const TRUCAMAUTOMATION_VERSION = "TRUCAM 2.0.2"
Public Const LASTRELEASED = "2023.05.08"
Public Const ABUTMENT_TYPE = "TA"
'-------------------------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------------------------
'----- CAM Automaion Setup Value
'-------------------------------------------------------------------------------------------------------------------
Public Const DEFAULT_TOLERANCE = "0.01"
Public Const ENDMILLTEMPLATE_LAYERNAME1MAIN = "[nn]DEG CROSS BALL ENDMILL" 'ENDMILL Template Layer Name - Main
Public Const ENDMILLTEMPLATE_LAYERNAME2LIMIT = "[nn]DEG 경계소재" 'ENDMILL Template Layer Name - Limit
Public Const ENDMILLTEMPLATE_LAYERNAME3MARGIN = "[nn]DEG 마진" 'ENDMILL Template Layer Name - Margin
Public Const ENDMILLTEMPLATE_FEATURECHAIN = "[nn]DEG ENDMILL ChainFeature"
Public Const ENDMILLTEMPLATE_FREEFORMNAME = "Template ENDMILL [경계소재-1+[nn]DEG 경계소재+경계소재-3]"
Public Const TRUDEFAULT_SMASHMINIMUMFACEANGLE = 20
'-------------------------------------------------------------------------------------------------------------------

Public Const CSIDL_DESKTOP = &H0   ' Desktop (namespace root)
Public Const CSIDL_DESKTOPDIRECTORY = &H10 ' Desktop folder ([user] profile)
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 ' Desktop folder (All Users profile)
Public Const MAX_PATH = 260
Public Const NOERROR = 0

Public Type shiEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As shiEMID
End Type

Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOREPOSITION = &H200
Public Const SWP_NOSIZE = &H1

Public Enum TruLeftRight
    TRU_LEFT = 0
    TRU_RIGHT = 1
    TRU_BOTH = 2
End Enum

Public Enum TruTopBottom
    TRU_BOTTOM = 0
    TRU_TOP = 1
    TRU_BOTH = 2
End Enum

Public Enum WindowState
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
    SW_MAX = 11
End Enum

'-------------------------------------------------------------------------------------------------------------------
'----- Get values from STL File Name in bracket
'----- The STL file name must be like blahblah_(XXX-CS-TA14,3).stl
Public Enum TruParameterCode
    PARA_ITEMCODE = 0
    PARA_PROGRAMNO = 1
    PARA_HOWMANYSECTION = 2
    PARA_SIZE = 9
End Enum
'-------------------------------------------------------------------------------------------------------------------


Public Enum TruHowManySectionsCode
    HOWMANY_3 = 3
    HOWMANY_4 = 4
    'Default HowManySections 4
    HOWMANY_DEFAULT = 4
End Enum

Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Sub InitTruCAM()
    
'Common Values
    Dim c_strPrcPath As String
    c_strPrcPath = Application.Configuration.GetFileDirectory(espFileTypeTemplate) & "\Presetfiles\" & ABUTMENT_TYPE
    
    Call Application.OutputWindow.Clear
    Application.OutputWindow.Text ("TRUCAMAUTOMATION_VERSION: " & TRUCAMAUTOMATION_VERSION & vbCrLf)
    Application.OutputWindow.Text ("LASTRELEASED: " & LASTRELEASED & vbCrLf)
    Application.OutputWindow.Text ("ABUTMENT_TYPE: " & ABUTMENT_TYPE & vbCrLf)
    
    If Not CheckPresetFiles(ABUTMENT_TYPE) Then
        Call MsgBox("Preset files are not founded. You must check the folder." + vbCrLf + c_strPrcPath, vbOKOnly)
    End If
    
End Sub

Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim IDL As ITEMIDLIST
    Dim sPath As String
    Dim iReturn As Long
    
    iReturn = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    
    If iReturn = NOERROR Then
        sPath = Space(512)
        iReturn = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        sPath = RTrim$(sPath)
        If Asc(Right(sPath, 1)) = 0 Then sPath = Left$(sPath, Len(sPath) - 1)
        GetSpecialfolder = sPath
        Exit Function
    End If
    GetSpecialfolder = ""
End Function
Sub ResetLayer(strLayerName As String)
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = strLayerName) Then
            Call Document.Layers.Remove(strLayerName)
            Exit For
        End If
    Next
    
    Call Document.Layers.Add(strLayerName)

End Sub
Sub AddLayersForAutomation()
    ResetLayer ("STL")
    ResetLayer ("DummyOperation")

End Sub

 
Function GetWorkFolder() As String
    Dim strDesk As String
    strDesk = GetSpecialfolder(CSIDL_DESKTOP)
    GetWorkFolder = strDesk + "\작업\"
    'or
    'strDesk = GetSpecialFolder(CSIDL_DESKTOPDIRECTORY)
    'or
    'strDesk = GetSpecialFolder(CSIDL_COMMON_DESKTOPDIRECTORY)
End Function
Function GetWorkScanfileFolder() As String
    GetWorkScanfileFolder = GetWorkFolder + "스캔파일\"
End Function
Function GetWorkEspritfileFolder() As String
    GetWorkEspritfileFolder = GetWorkFolder + "작업저장\"
End Function
 
Function GetFilenameWithoutExtension(ByVal FileName)
  Dim Result, i
  Result = FileName
  i = InStrRev(FileName, ".")
  If (i > 0) Then
    Result = Mid(FileName, 1, i - 1)
  End If
  GetFilenameWithoutExtension = Result
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#1 [1] Select STL file & locate properly.
Public Sub ClickBtn1()
    
    On Error GoTo 0
    
    If MsgBox("[1]Select STL file & locate properly.", vbYesNo, "CAM Automation") = vbYes Then
        
        Dim strOriginalEsppritFileName As String
        strOriginalEsppritFileName = Document.FileName
        
        'FindOrientationShortestDimension
        'FindOrientationSmallestArea
        Dim strSTLFilePath As String
        Dim File As stcFileStruct
        '// fill values (not required)
        File.strDialogtitle = "Select file to open"
        File.strFilter = "STL files (*.stl)|*.stl|All files (*.*)|*.*" '// use same format as
        
        'Task1 stl 스캔 파일 open 시 자동으로 기본설정 파일(.esp)과 해당 stl 파일 불러오기
        'Set default STL file path
        If (Strings.Right(Document.Name, 4) = ".esp") Then
            File.strFileName = Strings.Left(Document.Name, Strings.Len(Document.Name) - 4) + ".stl"
        Else
            File.strFileName = Document.Name + ".stl"
        End If
    
        If Dir(GetWorkFolder, vbDirectory) = "" Then
            If MsgBox(GetWorkFolder + " is not found. Do you want to make the folder?", vbYesNo) = vbYes Then
                MkDir GetWorkFolder
            End If
            If Dir(GetWorkScanfileFolder, vbDirectory) = "" Then
                If MsgBox(GetWorkScanfileFolder + " is not found. Do you want to make the folder?", vbYesNo) = vbYes Then
                    MkDir GetWorkScanfileFolder
                End If
            End If
            If Dir(GetWorkEspritfileFolder, vbDirectory) = "" Then
                If MsgBox(GetWorkEspritfileFolder + " is not found. Do you want to make the folder?", vbYesNo) = vbYes Then
                    MkDir GetWorkEspritfileFolder
                End If
            End If
            MsgBox ("Workfolders has been generated. Please check folders and files and try it again.")
            Exit Sub
        End If
    
    
        Dim strFileExists As String
        strFileExists = Dir(GetWorkScanfileFolder + File.strFileName)
        
        If strFileExists = "" Then
        'The selected file doesn't exist
        'Get the file manually.
            ShowOpenDialog File
            If File.strFileName = "" Then
                Call MsgBox("Please select an STL file and try it again.")
                Exit Sub
            End If
            strSTLFilePath = File.strFileName
            Document.SaveAs (GetWorkEspritfileFolder + GetFilenameWithoutExtension(File.strFileTitle) + ".esp")
            
            If strOriginalEsppritFileName = Document.FileName Then
                Exit Sub
            End If
        Else
            strSTLFilePath = GetWorkScanfileFolder + File.strFileName
        End If
    
        'CommonDialog Control
        '// pass stcFileStruct
        '// get return values (passed back through type)
        'strSTLFilePath = ".\Core Dental Studio_16731_1_Hiossen ET Regular_pm.stl"
        
        
        'If Not CopyFileToBackup(File.strFileTitle, "", Replace(File.strFileName, File.strFileTitle, ""), "E:\Esprit\개발관련\Step1\STLFiles\B\") Then
        'Task1 stl 스캔 파일 open 시 자동으로 기본설정 파일(.esp)과 해당 stl 파일 불러오기
        'esp file already copied and open with external program.
        'Document.SaveAs (GetWorkEspritfileFolder + GetFilenameWithoutExtension(File.strFileTitle) + ".esp")
        'If Then
        '    MsgBox ("Back-up The STL file is failed. Please check it and try it again.")
        '    'Exit Sub
        'End If
        
        
        'If strOriginalEsppritFileName = Document.FileName Then
        '    Exit Sub
        'End If

        
        Dim strTurning As String
        strTurning = ""
        'Call GetSTL(strSTLFilePath)
        Call CheckSTL(strSTLFilePath)
        Document.Refresh
        
        Dim cCondition As Boolean
        cCondition = True
        Do While cCondition
            Select Case MsgBox("Is the STL properly located?", vbYesNoCancel)
                Case vbYes
                    If CheckSTLInTheCircle() Then
                        Call SelectSTL_Model
                        cCondition = False
                    End If
                Case vbNo
                    strTurning = InputBox("Enter X or Y and degree like (X,90).", "CAM Automation - To transform the STL", "X,0")
                    Call TurnSTL(strTurning)
                    cCondition = True
                Case vbCancel
                    cCondition = False
                Case Else
                    cCondition = False
            End Select
        Loop
    Else
        Exit Sub
    End If


    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        End If
    Next

    For Each ly In Document.Layers
        If (ly.Name = "경계소재-1") Then
            ly.Visible = False
        End If
    Next
    Document.Refresh

    Load frmSTLRotate
    Call frmSTLRotate.RunDirectionCheck(1)
    frmSTLRotate.Show (vbModeless)

    Exit Sub
    
Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#1-2 [1-2] Generate toolpaths for [FRONT TURNING]. Please make sure the STL properly located."
Public Sub ClickBtn1_2()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: GenerateFreeFormsCROSSBALLENDMILL
' Description: Generate ENDMILL Freeforms from Template.
' Created By: Ian Pak(Tru IT/SW)
' Created At: 03/17/2023
' Last Updated: 03/17/2023
'
' Parameter
' pBaseWorkPlaneName As String:
' Optional pHowManySections As Integer = 3:
' Optional pByDeg As Integer = 120:
'
' Usage: Call GenerateFreeFormsCROSSBALLENDMILL(pBaseWorkPlaneName, pHowManySections, pByDeg)

    On Error GoTo Err_handler
    
    Load frmSTLRotate
    frmSTLRotate.Hide
    Unload frmSTLRotate
    
    Dim nErrorCode As Integer
    nErrorCode = 0

    If MsgBox("[2] Generate toolpaths for [FRONT TURNING]. Please make sure the STL properly located.", vbYesNo, "CAM Automation") = vbYes Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. Step1_2 + Step2_4: STL + 경계소재-1 합한 개체의 TurningProfile로 FRONT TURNING Layer에 가공 관련 개체들 자동 생성

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '1-1. Step1_2: Front Turning & lift the horizontal line in the "경계소재-1" layer
        nErrorCode = Step1_2
        'When Step1_2 failed, exit Sub
        If nErrorCode > 0 Then
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '1-2. Step2_4: Auto process to generate tool path in the Front Turning layer
            nErrorCode = Step2_4
        Else
            Call GetLayer("FRONT TURNING", 1)
            Exit Sub
        End If
        
        'When Step2_4 failed, exit Sub
        If nErrorCode < 0 Then
            Select Case nErrorCode
            Case -991
                Call MsgBox("Cannot find a parallel segment.", vbCritical, "Error in Front Turning(Step2_4)")
                Call GetLayer("FRONT TURNING", 1)
                Exit Sub
            Case Else
                Call MsgBox("Error in Step2_4(Draw toolpath segments and arcs automatically.).", vbCritical, "Error in Front Turning(Step2_4)")
                Call GetLayer("FRONT TURNING", 1)
                Exit Sub
            End Select
        End If
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '1-3. Step2_6: connect segments and arcs into the only featurechain. And make the featurechain into tool path in the Front Turning layer
        nErrorCode = Step2_6
        'When Step2_6 failed, exit Sub
        If nErrorCode < 0 Then
            Select Case nErrorCode
            Case -991
                Call MsgBox("More than 2 Chain features are made. Please check it.", vbCritical, "Error in Front Truning.(Step2_6)")
                Call GetLayer("FRONT TURNING", 1)
                Exit Sub
            Case Else
                Call MsgBox("Error in Step2_6(Create A chain feature.).", vbCritical, "Error in Front Turning(Step2_6)")
                Call GetLayer("FRONT TURNING", 1)
                Exit Sub
            End Select
        End If
    Else
        Exit Sub
    End If
   
ClickBtn1_2END:
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        End If
    Next
    Exit Sub

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#2 [2] Generate toolpaths for [ROUGH ENDMILL R6.0]. Please make sure the STL properly located."
Public Sub ClickBtn2()
    If Set_AttributeValues() <= 0 Then
                If MsgBox("[2-1] Select How-many-work-sections and generate CROSS BALL ENDMILL R0.75 Layers and Freeforms. Please make sure the STL properly located.", vbYesNo, "CAM Automation") = vbYes Then
                        Unload frmBaseWorkPlaneDeg
                        Load frmBaseWorkPlaneDeg
                        frmBaseWorkPlaneDeg.Show
                        
                        'GetBaseWorkPlaneAndStepDgreeAndRoughHowMany(strParse As String, strType As String)
        'Function Get_nRoughStepByDegree() As Integer
        'Function Get_nRoughHowMany() As Integer
        'Function Get_strBaseWorkPlaneName() As String
                        
                        'Dim nBaseWorkPlaneDeg As Integer
                        'Dim nRoughHowMany As Integer
                        'nBaseWorkPlaneDeg = Get_nRoughStepByDegree
                        'nRoughHowMany = Get_nRoughHowMany
                        
        '1. Generate ENDMILL Layers from Template
        '2. Copy ENDMILL Template - Freeforms
                        If Get_bSetBaseWorkPlane() Then
                                Call generateEndmillTemplates(strBaseWorkPlaneName_pub, Get_nHowManySections(), 360 / Get_nHowManySections)
                                If Not (GetLayer(ENDMILLTEMPLATE_LAYERNAME1MAIN) Is Nothing) Then
                                        Document.Layers.Remove (ENDMILLTEMPLATE_LAYERNAME1MAIN)
                                End If
                        Else
                                Call MsgBox("Base Work Plane/How-many sections/Rough Mill values are not set-up yet. Please check it first.")
                                Exit Sub
                        End If

                End If
    Else
        Call MsgBox("It is set from the saved value.(How-many-work-sections, base work pane, and etc..)")
    End If
        
    'For Test Remove Layer
    'If MsgBox("[*] Delete template layer?", vbYesNo, "CAM Automation") = vbYes Then
    '    Document.Layers.Remove ("[nn]DEG CROSS BALL ENDMILL")
    'End If
        
    If MsgBox("[2-2] Generate toolpaths for [ROUGH ENDMILL R6.0]. Please make sure the STL properly located.", vbYesNo, "CAM Automation") = vbYes Then
'3. Generate Solid Mill Turn (ROUGH END MILL)
'4. Open Form: frmCreateBorderSolidObject
        'generateSolidmilTurnWithBaseWorkPlane (strBaseWorkPlaneName_pub)
        If generateSolidmilTurnWithBaseWorkPlane(strBaseWorkPlaneName_pub, Get_nRoughStepByDegree, Get_nRoughHowMany) Then
            'Document.ActivePlane = Document.Planes("0DEG")
            If MsgBox("Reorder Operation and show checking Rough Endmill.", vbYesNo) = vbYes Then
                Call ReorderOperation
                Unload frmCreateBorderSolidObject
                Load frmCreateBorderSolidObject
                frmCreateBorderSolidObject.MultiPage1.Value = 0
                
                'Show 선반소재(MaskLatheStock)
                Call Document.Windows.ActiveWindow.SetMask(espViewMaskLatheStock, True)
                Document.Refresh
                frmCreateBorderSolidObject.Show (vbModeless)
                
            End If
        Else
            Call MsgBox("It is failed [2-2] Generate toolpaths for [ROUGH ENDMILL R6.0].", vbOKOnly)
        End If
    ElseIf MsgBox("[2-3] Reorder Operation and show checking Rough Endmill.", vbYesNo) = vbYes Then
                Call ReorderOperation
                Unload frmCreateBorderSolidObject
                Load frmCreateBorderSolidObject
                frmCreateBorderSolidObject.MultiPage1.Value = 0
                
                'Show 선반소재(MaskLatheStock)
                Call Document.Windows.ActiveWindow.SetMask(espViewMaskLatheStock, True)
                Document.Refresh
                frmCreateBorderSolidObject.Show (vbModeless)
    Else
        Exit Sub
    End If
End Sub
Public Sub generateEndmillTemplates(pBaseWorkPlaneName As String, Optional pHowManySections As Integer = 3, Optional pByDeg As Integer = 120)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Created By: Ian Pak(Tru IT/SW)
' Created At: mm/dd/yyyy
' Last Updated: mm/dd/yyyy
'
' Parameter
' pBaseWorkPlaneName As String:
' Optional pHowManySections As Integer = 3:
' Optional pByDeg As Integer = 120:
'

'Common Values

'Work variables

'Initialize
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. Generate Layers for ENDMILL works.
    Call GenerateLayerCROSSBALLENDMILL(pBaseWorkPlaneName, pHowManySections, pByDeg)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2. Copy Template ENDMILL ChainFeature to the generated layers.
'Work variables

    Dim strmSelectionIndex As String
    Dim mSelection As Esprit.SelectionSet
    Dim strWorkPlane As String
    strWorkPlane = Document.ActivePlane.Name
    
    Dim goRef As Esprit.graphicObject
    Dim plRef As Esprit.Plane
    Dim lyCurr As Esprit.Layer
    Dim strLayerName As String
    Dim strCopiedChainFeatureName As String
    strLayerName = ""
    strCopiedChainFeatureName = ""
    
    Dim strChkLayerName As String
    Dim strTempLayer As String
    
    Dim dUnit As Double
    Dim iLine As Esprit.Line
    Dim i As Integer
    
    '2-1. Copy ChainFeature & Rotate it according to the WorkPlaneDegree
    i = 0
    For i = 0 To pHowManySections - 1
        'Initialize mSelection
        dUnit = (GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * i) Mod 360
        strmSelectionIndex = CStr(dUnit)
        
        strLayerName = Replace(ENDMILLTEMPLATE_LAYERNAME1MAIN, "[nn]", strmSelectionIndex)
        strCopiedChainFeatureName = Replace(ENDMILLTEMPLATE_FEATURECHAIN, "[nn]", strmSelectionIndex)
        Set mSelection = CopyTemplateChainFeature(ENDMILLTEMPLATE_FEATURECHAIN, ENDMILLTEMPLATE_LAYERNAME1MAIN, strCopiedChainFeatureName, strLayerName)
        
        'Rotate the copied ChainFeature to align the WorkPlane Degree
        'Get the copied ChainFeature
        If mSelection.Count = 0 Then
            MsgBox ("Cannot find the ChainFeature in the layer of " + ENDMILLTEMPLATE_LAYERNAME1MAIN + ". Please check it first.")
            Exit Sub
        End If
        Set iLine = getTheOriginAxis("U")
    
        'Rotate by the Axis & Degrees from the parameter
        Call mSelection.Rotate(iLine, Get_DegreeToRadian(dUnit), 0)
        'For Debug
        'Document.Refresh
    Next
    'Document.Refresh
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'3. Generate ENDMILL Freeforms from Template.
    Call GenerateFreeFormsCROSSBALLENDMILL(pBaseWorkPlaneName, pHowManySections, pByDeg)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'4. Refresh Document
    'Document.Refresh
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finalize. Release Resources


generateEndmillTemplatesEND:
End Sub
Sub GenerateLayerCROSSBALLENDMILL(pBaseWorkPlaneName As String, Optional pHowManySections As Integer = 3, Optional pByDeg As Integer = 120)
'pBaseWorkPlaneName:
'pHowManySections:
'pByDeg:

    On Error GoTo 0
    
'Common Values

'Work variables
    Dim t_strLayerName As String
    Dim t_strChkLayerName As String
    t_strLayerName = ""
    t_strChkLayerName = ""
    
    Dim dUnit As Double
    Dim i As Integer
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. Generate Main Layers - DEG CROSS BALL ENDMILL
    i = 0
    For i = 0 To pHowManySections - 1
        dUnit = (GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * i) Mod 360
        t_strLayerName = Replace(ENDMILLTEMPLATE_LAYERNAME1MAIN, "[nn]DEG", CStr(dUnit) + "DEG")
        t_strChkLayerName = createLayer(t_strLayerName)
        If t_strChkLayerName <> t_strLayerName Then
            Call MsgBox("Cannot generate a Layer [" + t_strLayerName + "].", vbOKOnly, "Alert")
            GoTo GenerateLayerCROSSBALLENDMILLEND
        End If
    Next
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2. Generate Limit Layers - DEG 경계소재
    i = 0
    For i = 0 To pHowManySections - 1
        dUnit = (GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * i) Mod 360
        t_strLayerName = Replace(ENDMILLTEMPLATE_LAYERNAME2LIMIT, "[nn]DEG", CStr(dUnit) + "DEG")
        t_strChkLayerName = createLayer(t_strLayerName)
        If t_strChkLayerName <> t_strLayerName Then
            Call MsgBox("Cannot generate a Layer [" + t_strLayerName + "].", vbOKOnly, "Alert")
            GoTo GenerateLayerCROSSBALLENDMILLEND
        End If
    Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'3. Generate Margin Layers - DEG 마진
    i = 0
    For i = 0 To pHowManySections - 1
        dUnit = (GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * i) Mod 360
        t_strLayerName = Replace(ENDMILLTEMPLATE_LAYERNAME3MARGIN, "[nn]DEG", CStr(dUnit) + "DEG")
        t_strChkLayerName = createLayer(t_strLayerName)
        If t_strChkLayerName <> t_strLayerName Then
            Call MsgBox("Cannot generate a Layer [" + t_strLayerName + "].", vbOKOnly, "Alert")
            GoTo GenerateLayerCROSSBALLENDMILLEND
        End If
    Next

GenerateLayerCROSSBALLENDMILLEND:
End Sub

Function CopyTemplateChainFeature(pTemplateFeatureChainName As String, pTemplateLayerName As String, pCopiedChainFeatureName As String, pWorkLayerName As String) As Esprit.SelectionSet
'Work Variable
    Dim fcTemplate As Esprit.FeatureChain
    Dim lyOriginal As Esprit.Layer
    Set lyOriginal = Document.ActiveLayer
    Dim lyCurr As Esprit.Layer
    Dim mSelection As Esprit.SelectionSet
    Dim goRef As Esprit.graphicObject
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Set fcTemplate = GetFeatureChain(pTemplateFeatureChainName)
    setLayersFor (pWorkLayerName)
    Set lyCurr = Document.ActiveLayer
    
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(pWorkLayerName)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(pWorkLayerName)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.GraphicObjectType = espFeatureChain) Then
            If (.Layer.Name = pTemplateLayerName) Then
                If (.ComGraphicObject.Name = pTemplateFeatureChainName) Then
                    If (.Key > 0) Then
                    .Grouped = True
                    Call mSelection.Add(goRef)
                    Exit For
                End If
            End If
        End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    If mSelection.Count = 0 Then
        MsgBox ("Cannot find the Template FeatureChain of " + pTemplateFeatureChainName + ". Please check it first.")
        Exit Function
    End If
    'Copy the template ChainFeature
    Call mSelection.ChangeLayer(lyCurr, 1)
    
    'Get the Copied ChainFeature as a SelectionSets
    'reset mSelection
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    'search the copied FeatureChain.
    For Each goRef In Esprit.Document.GraphicsCollection
        With goRef
        If (.GraphicObjectType = espFeatureChain) Then
            If (.Layer.Name = pWorkLayerName) Then
                If (.ComGraphicObject.Name = pTemplateFeatureChainName) Then
                    If (.Key > 0) Then
                    '.ComGraphicObject.Name = Replace(ENDMILLTEMPLATE_FEATURECHAIN, "[nn]", strmSelectionIndex)
                    .ComGraphicObject.Name = pCopiedChainFeatureName
                    .Grouped = True
                    'Add it into mSelection.
                    Call mSelection.Add(goRef)
                    Exit For
                End If
            End If
        End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
CopyTemplateChainFeatureEND:
    Set CopyTemplateChainFeature = mSelection
End Function


Function GenerateFreeFormsCROSSBALLENDMILL(pBaseWorkPlaneName As String, Optional pHowManySections As Integer = 3, Optional pByDeg As Integer = 120) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: GenerateFreeFormsCROSSBALLENDMILL
' Description: Generate ENDMILL Freeforms from Template.
' Created By: Ian Pak(Tru IT/SW)
' Created At: 03/17/2023
' Last Updated: 03/17/2023
'
' Parameter
' pBaseWorkPlaneName As String:
' Optional pHowManySections As Integer = 3:
' Optional pByDeg As Integer = 120:
'
' Usage: Call GenerateFreeFormsCROSSBALLENDMILL(pBaseWorkPlaneName, pHowManySections, pByDeg)

    On Error GoTo Err_handler

'Return Value
    Dim rtnValue As Integer
    rtnValue = 0
    
'Common Values
    Dim M_TechnologyUtility As EspritTechnology.TechnologyUtility
    Dim FileName As String
    Dim tech() As EspritTechnology.Technology
    Dim techLMPPNew As EspritTechnology.TechLatheMoldParallelPlanes

    Set M_TechnologyUtility = Document.TechnologyUtility
    FileName = "ENDMILL_freeform.prc"
    tech = M_TechnologyUtility.OpenProcess(EspritUserFolder(ABUTMENT_TYPE) & FileName)
    Set techLMPPNew = tech(0)
    
    Dim c_bTopZLimit As Double
    Dim c_bBottomZLimit  As Double
    Dim c_bFullClearance As Double
    Dim c_bClearance  As Double
    Dim c_bStepOver_Upper  As Double
    Dim c_bStepOver_Lower  As Double
    Dim c_bStepPercentOfDiameter_Upper  As Double
    Dim c_bStepPercentOfDiameter_Lower  As Double
    Dim c_nPositionOnBoundaryProfile_Upper As espMoldPositionOnBoundaryProfile
    Dim c_nPositionOnBoundaryProfile_Lower As espMoldPositionOnBoundaryProfile
    
    Select Case GetParameterValue(PARA_SIZE)
    Case "14":
        c_bTopZLimit = 14                   '리미트>상단Z리미트 [상단,하단]
        c_bBottomZLimit = 0                 '리미트>바닥Z리미트 [상단,하단]
        c_bFullClearance = 14               '링크>전체여유 [상단,하단]
        c_bClearance = 7                    '링크>여유 [상단,하단]
        c_bStepOver_Upper = 0.075           '툴패스>스텝오버 [상단]
        c_bStepOver_Lower = 0.045           '툴패스>직경%값 [상단]
        c_bStepPercentOfDiameter_Upper = 5  '툴패스>스텝오버 [하단]
        c_bStepPercentOfDiameter_Lower = 3  '툴패스>직경%값 [하단]
        c_nPositionOnBoundaryProfile_Upper = espMoldPositionOnBoundaryProfileOutside    '리미트>경계>경계선 프로파일에 위치 [상단]
        c_nPositionOnBoundaryProfile_Lower = espMoldPositionOnBoundaryProfileInside     '리미트>경계>경계선 프로파일에 위치 [하단]
    Case "10", Default:
        c_bTopZLimit = 10                   '리미트>상단Z리미트 [상단,하단]
        c_bBottomZLimit = 0                 '리미트>바닥Z리미트 [상단,하단]
        c_bFullClearance = 10               '링크>전체여유 [상단,하단]
        c_bClearance = 5                    '링크>여유 [상단,하단]
        c_bStepOver_Upper = 0.075           '툴패스>스텝오버 [상단]
        c_bStepPercentOfDiameter_Upper = 5  '툴패스>직경%값 [상단]
        c_bStepOver_Lower = 0.045           '툴패스>스텝오버 [하단]
        c_bStepPercentOfDiameter_Lower = 3  '툴패스>직경%값 [하단]
        c_nPositionOnBoundaryProfile_Upper = espMoldPositionOnBoundaryProfileOutside    '리미트>경계>경계선 프로파일에 위치 [상단]
        c_nPositionOnBoundaryProfile_Lower = espMoldPositionOnBoundaryProfileInside     '리미트>경계>경계선 프로파일에 위치 [하단]
    End Select
    
    'Upper, Lower
    techLMPPNew.TopZLimit = c_bTopZLimit
    techLMPPNew.BottomZLimit = c_bBottomZLimit
    techLMPPNew.FullClearance = c_bFullClearance
    techLMPPNew.Clearance = c_bClearance
    techLMPPNew.StepOver = c_bStepOver_Upper
    techLMPPNew.StepPercentOfDiameter = c_bStepPercentOfDiameter_Upper
    techLMPPNew.PositionOnBoundaryProfile = c_nPositionOnBoundaryProfile_Upper
    
    Dim c_goSTL As Esprit.graphicObject
    Set c_goSTL = Get_STLObject()
    
'Work variables
    Dim fffWork As Esprit.FreeFormFeature
    Dim opWork As Esprit.Operation
    
    Dim dUnit As Double
    Dim i As Integer
    Dim strmSelectionIndex As String
    Dim strLayerName As String
    
'Initialize
    dUnit = 0
    i = 0
    strmSelectionIndex = ""
    strLayerName = ""
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BEGINS

    For i = 0 To pHowManySections - 1
        'dUnit = GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * i
        dUnit = (GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * i) Mod 360
        strmSelectionIndex = CStr(dUnit)
        strLayerName = Replace(ENDMILLTEMPLATE_LAYERNAME1MAIN, "[nn]", strmSelectionIndex)
        
        'Change Layer & WorkPlane
        setLayersFor (strLayerName)
        setWorkPlane (strmSelectionIndex + "DEG")
        
        'Add FreeFormFeature in the layer
        Set fffWork = Nothing
        Set fffWork = Document.FreeFormFeatures.Add()
        Call fffWork.Add(c_goSTL, espFreeFormPartSurfaceItem)
        'Change the FreeFormFeature name
        fffWork.Name = Replace(Replace(ENDMILLTEMPLATE_FREEFORMNAME, "Template ENDMILL", "ENDMILL-" + CStr(i + 1)), "[nn]", strmSelectionIndex)
        
        'Generate & Add operation in the freeform
        'Upper Operation
        techLMPPNew.StepOver = c_bStepOver_Upper
        techLMPPNew.StepPercentOfDiameter = c_bStepPercentOfDiameter_Upper
        techLMPPNew.PositionOnBoundaryProfile = c_nPositionOnBoundaryProfile_Upper
        Set opWork = Document.Operations.Add(techLMPPNew, fffWork)
        opWork.Name = "8-" + CStr(i * 2 + 1) + ". " + strmSelectionIndex + "DEG CROSS BALL ENDMILL R0.75"
        opWork.Suppress = False
        
        'Lower Operation
        techLMPPNew.StepOver = c_bStepOver_Lower
        techLMPPNew.StepPercentOfDiameter = c_bStepPercentOfDiameter_Lower
        techLMPPNew.PositionOnBoundaryProfile = c_nPositionOnBoundaryProfile_Lower
        Set opWork = Document.Operations.Add(techLMPPNew, fffWork)
        opWork.Name = "8-" + CStr(i * 2 + 2) + ". " + strmSelectionIndex + "DEG-1 CROSS BALL ENDMILL R0.75"
        opWork.Suppress = True
    
    Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'END

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finalize. Release Resources

    Set techLMPPNew = Nothing
    Set tech(0) = Nothing
    Set opWork = Nothing
    
    'To fix the last fff missing STL Part object.
'    fffWork.RemoveAll
    Set fffWork = Nothing
    Set c_goSTL = Nothing

GenerateFreeFormsCROSSBALLENDMILLEND:
    GenerateFreeFormsCROSSBALLENDMILL = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Function

Public Function createLayer(pstrLayerName As String) As String
'pstrLayerName

'Return variables
    Dim rtn_createLayer As String
    rtn_createLayer = ""
    
'Common Values
'Work variables
    Dim t_ly As Esprit.Layer
    Dim t_lyNew As Esprit.Layer

    
    For Each t_ly In Document.Layers
        If (t_ly.Name = pstrLayerName) Then
            Call Document.Layers.Remove(pstrLayerName)
            Exit For
        End If
    Next
    Set t_lyNew = Document.Layers.Add(pstrLayerName)
    rtn_createLayer = t_lyNew.Name
    
createLayerEND:
    createLayer = rtn_createLayer
End Function

Private Function generateSolidmilTurnWithBaseWorkPlane(pBaseWorkPlaneName As String, Optional pByDeg As Integer = 90, Optional pRoughHowMany As Integer = 2) As Boolean
'Return Value
    Dim rtnValue As Boolean
    rtnValue = False

    Dim strRtnName(3) As String
    Dim nRtn(3) As Integer
    nRtn(0) = 0
    nRtn(1) = 0
    nRtn(2) = 0
    strRtnName(0) = pBaseWorkPlaneName
    strRtnName(1) = CStr(GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * 1) + "DEG"
    strRtnName(2) = CStr(GetDegreeNumberInt(pBaseWorkPlaneName) + pByDeg * 2) + "DEG"
    
    Dim i As Integer
    Dim bChecker As Boolean
    i = 0
    bChecker = True
    For i = 0 To pRoughHowMany - 1
        If MsgBox("It is processing to " + strRtnName(i), vbYesNo) = vbYes Then
'''''''''''''''''''''''''''''''''''''''''''
'Generate Solid Mill Turn (ROUGH END MILL)
            nRtn(i) = generateSolidmilTurn(strRtnName(i), "ROUGH ENDMILL R6.0", CStr(i + 1))
            bChecker = (nRtn(i) = 1) 'To make sure each of nRtn() is true
        End If
    Next

    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        End If
    Next
    
    rtnValue = bChecker

'Original Code before 2023.03.01.
'        Dim nRtn_0Deg As Integer
'        Dim nRtn_120Deg As Integer
'        Dim nRtn_240Deg As Integer
'        nRtn_0Deg = 0
'        nRtn_120Deg = 0
'        nRtn_240Deg = 0
'
'        If MsgBox("It is processing to 0DEG.", vbYesNo) = vbYes Then
'            nRtn_0Deg = generateSolidmilTurn("0DEG", "ROUGH ENDMILL R6.0", "1")
'        End If
'        If MsgBox("It is processing to 120DEG.", vbYesNo) = vbYes Then
'            nRtn_120Deg = generateSolidmilTurn("120DEG", "ROUGH ENDMILL R6.0", "2")
'        End If
'        If MsgBox("It is processing to 240DEG.", vbYesNo) = vbYes Then
'            nRtn_240Deg = generateSolidmilTurn("240DEG", "ROUGH ENDMILL R6.0", "3")
'        End If
'
'        Dim ly As Esprit.Layer
'        For Each ly In Document.Layers
'            If (ly.Name = "STL" Or ly.Name = "기본값") Then
'                ly.Visible = True
'            End If
'        Next
'
'        If (nRtn_0Deg = 1 And nRtn_120Deg = 1 And nRtn_240Deg = 1) Then
'            If MsgBox("Reorder Operation and show checking Rough Endmill.", vbYesNo) = vbYes Then
'                Call ReorderOperation
'                Unload frmCreateBorderSolidObject
'                Load frmCreateBorderSolidObject
'                frmCreateBorderSolidObject.MultiPage1.Value = 0
'
'                'Show 선반소재(MaskLatheStock)
'                Call Document.Windows.ActiveWindow.SetMask(espViewMaskLatheStock, True)
'                Document.Refresh
'                frmCreateBorderSolidObject.Show (vbModeless)
'
'            End If
'        End If

generateSolidmilTurnWithBaseWorkPlaneEND:
    generateSolidmilTurnWithBaseWorkPlane = rtnValue
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#3 [3] Create 경계소재2 - border material 2nd
Public Sub ClickBtn3()
    If Set_AttributeValues() <= 0 Then
        Call MsgBox("Attribute Value is not set up yet. Please check it first.", vbOKOnly)
        Exit Sub
    Else
        Call MsgBox("It is set from the saved value.(How-many-work-sections, base work pane, and etc..)")
    End If

    Unload frmCreateBorderSolidObject
    Load frmCreateBorderSolidObject
    frmCreateBorderSolidObject.MultiPage1.Value = 1
    frmCreateBorderSolidObject.Show (vbModeless)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#4 Rebuild Freeform
Public Sub ClickBtn4()
    'Temp for test
    'Call ClickBtn2
    
    If MsgBox("[4] Rebuild Freeform With Part & Check Elements. " & vbCrLf & "* Template Layer '" + ENDMILLTEMPLATE_LAYERNAME1MAIN + "' will be deleted.", vbYesNo, "CAM Automation") = vbYes Then
        If Not (GetLayer(ENDMILLTEMPLATE_LAYERNAME1MAIN) Is Nothing) Then
            Document.Layers.Remove (ENDMILLTEMPLATE_LAYERNAME1MAIN)
        End If
        
        If Set_AttributeValues() <= 0 Then
            Call MsgBox("Attribute Value is not set up yet. Please check it first.", vbOKOnly)
            Exit Sub
        End If
        
        Call SetBoundaryOperationAll
        Call RebuildFreeformWithCheckElements
        Call RebuildFreeformWithNewSTLAsAPartElement
    Else
        Exit Sub
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' BTN#5 [5] Create Margin
Public Sub ClickBtn5()
    If Set_AttributeValues() <= 0 Then
        Call MsgBox("Attribute Value is not set up yet. Please check it first.", vbOKOnly)
        Exit Sub
    End If

    Unload frmCreateMargin
    Load frmCreateMargin
    frmCreateMargin.Show (vbModeless)
End Sub


Public Sub ClickBtnR()
    If MsgBox("[R] Reorder operations.", vbYesNo, "CAM Automation") = vbYes Then
        Call ReorderOperation
        Unload frmNCCodeReady
        Load frmNCCodeReady
        frmNCCodeReady.Show (vbModeless)
    Else
        Exit Sub
    End If
End Sub

Public Sub ClickBtnT()
    Dim strPGMNumber As String
    strPGMNumber = GetProgramNumber()
    '1. 데이터 페이지에 프로그램번호 입력
    Document.LatheMachineSetup.ProgramNumber = strPGMNumber
    
    Dim hdTemp As Esprit.Head
    For Each hdTemp In Document.LatheMachineSetup.Heads
        hdTemp.ProgramNumber = strPGMNumber
    Next
    
    Document.MillMachineSetup.ProgramNumber = strPGMNumber
    
    '2. Engraving Program Number
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "STL" Or ly.Name = "기본값") Then
            ly.Visible = True
        ElseIf (ly.Name = "TEXT") Then
            ly.Visible = True
            Document.ActiveLayer = ly
        Else
            ly.Visible = False
        End If
    Next
    Document.Refresh
    
    Unload frmPGMText
    Load frmPGMText
    frmPGMText.Show (vbModeless)
    
End Sub

Private Function CheckSTLInTheCircle() As Boolean
'Is The STL in the Circle?

Call GetPartProfileSTL

'Function IntersectCircleAndArcsSegments(ByRef cBound As Esprit.Circle, _
'            ByRef go2 As Esprit.graphicObject) As Esprit.Point

    Dim cBound As Esprit.Circle
    Dim go2 As Esprit.graphicObject
    Dim pntIntersect As Esprit.Point
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
        
    Dim crl As Esprit.Circle
    For Each crl In Esprit.Document.Circles
        With crl
        If (.Layer.Name = "기본값") Then
            If (.Key > 0) Then
                Set cBound = crl
                .Grouped = True
                .Layer.Visible = True
                Set layerObject = .Layer
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next

    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSegment Or .GraphicObjectType = espArc)) Then
            If (.Key > 0) Then
                Set go2 = graphicObject
                .Grouped = True
                .Layer.Visible = True
                Set pntIntersect = IntersectCircleAndArcsSegments(cBound, go2)
                Call Document.GraphicsCollection.Remove(go2.GraphicsCollectionIndex)
                Document.Refresh
                
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    'Dim bInTheCircle As Boolean
    CheckSTLInTheCircle = (pntIntersect Is Nothing)
    
    If (CheckSTLInTheCircle And Not (cBound Is Nothing)) Then
        Call Document.GraphicsCollection.Remove(cBound.GraphicsCollectionIndex)
    End If
    
    Document.Refresh


End Function


Public Sub OnTop(hwnd As Long)
    '
    ' put hWnd always on top
    '
    Call SetWindowPos(hwnd, -1, 0, 0, 0, 0, &H2 Or &H1)

End Sub


Public Sub OffTop(hwnd As Long)

    Call SetWindowPos(hwnd, -2, 0, 0, 0, 0, &H2 Or &H1)

End Sub
Public Function GetParameterStringFromFileName() As String

    Dim strFileName As String
    Dim strCodes() As String
    Dim strParameter As String
    

    strFileName = Document.FileName
    strCodes = Split(strFileName, "(")

    strParameter = strCodes(UBound(strCodes))
    strCodes = Split(strParameter, ")")
    strParameter = strCodes(LBound(strCodes))

    'PGMCode Check Logic recommended.

GetParameterStringFromFileNameEND:
    GetParameterStringFromFileName = strParameter

End Function
Public Function GetParameterValue(pParameter As TruParameterCode) As String
    Dim rtnStrValue As String
    Dim strParameter As String
    Dim strCodes() As String
    
    strParameter = GetParameterStringFromFileName()
    strCodes = Split(strParameter, ",")
    
    If UBound(strCodes) < 1 Or UBound(strCodes) > 2 Then
        MsgBox ("FileName is not properly set. Should end with (LibraryCode, ProgramNumber, HowManySections/optional)")
    ElseIf pParameter = TruParameterCode.PARA_ITEMCODE Then
        rtnStrValue = strCodes(TruParameterCode.PARA_ITEMCODE)
    ElseIf pParameter = TruParameterCode.PARA_SIZE Then
        rtnStrValue = Right(strCodes(TruParameterCode.PARA_ITEMCODE), 2)
    ElseIf pParameter = TruParameterCode.PARA_PROGRAMNO Then
        rtnStrValue = strCodes(TruParameterCode.PARA_PROGRAMNO)
    ElseIf pParameter = TruParameterCode.PARA_HOWMANYSECTION _
       And UBound(strCodes) = 1 Then
        rtnStrValue = CStr(TruHowManySectionsCode.HOWMANY_DEFAULT)
    ElseIf pParameter = TruParameterCode.PARA_HOWMANYSECTION _
       And UBound(strCodes) = 2 Then
        rtnStrValue = strCodes(TruParameterCode.PARA_HOWMANYSECTION)
    Else
        MsgBox ("FileName is not properly set. Should end with (LibraryCode, ProgramNumber, HowManySections/optional)")
    End If
    
'PGMCode Check Logic recommended.

GetParameterValueEND:
    GetParameterValue = rtnStrValue

End Function

Public Function GetProgramNumber() As String
    'PGMCode Check Logic recommended.
    GetProgramNumber = GetParameterValue(PARA_PROGRAMNO)
End Function
Public Function GetHowManySectionsCode() As String

    'strParameter = GetParameterStringFromFileName()
    'strParameter = "XXX-CS-TA14,1234,3"

GetHowManySectionsCodeEND:
    GetHowManySectionsCode = GetParameterValue(PARA_HOWMANYSECTION)

End Function

Public Function CheckSTLDirectionLine() As Boolean
'Is The STL in the Circle?

Call GetPartProfileSTL(0.01)

'Function IntersectCircleAndArcsSegments(ByRef cBound As Esprit.Circle, _
'            ByRef go2 As Esprit.graphicObject) As Esprit.Point

    Dim sTestSegment As Esprit.Segment
    Dim go2 As Esprit.graphicObject
    Dim pntIntersect As Esprit.Point
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
        
    Dim seg As Esprit.Segment
    For Each seg In Esprit.Document.Segments
        With seg
        'If (.Layer.Name = "방향체크") Then
        If InStr(1, .Layer.Name, "방향체크") <> 0 Then
            If (.Key > 0) Then
                Set sTestSegment = seg
                .Grouped = True
                .Layer.Visible = True
                Set layerObject = .Layer
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next

    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSegment Or .GraphicObjectType = espArc)) Then
            If (.Key > 0) Then
                Set go2 = graphicObject
                .Grouped = True
                .Layer.Visible = True
                If (pntIntersect Is Nothing) Then
                    Set pntIntersect = IntersectSegmentAndArcsSegments(sTestSegment, go2)
                    Call Document.GraphicsCollection.Remove(go2.GraphicsCollectionIndex)
                Else
                    Call Document.GraphicsCollection.Remove(go2.GraphicsCollectionIndex)
                End If
                Document.Refresh
                
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    
    'Dim bInTheCircle As Boolean
    CheckSTLDirectionLine = Not (pntIntersect Is Nothing)
    
    'If (CheckSTLDirectionLine And Not (sTestSegment Is Nothing)) Then
    '    Call Document.GraphicsCollection.Remove(sTestSegment.GraphicsCollectionIndex)
    'End If
    
    Document.Refresh


End Function

Public Function GetConnection(pstrReturnClass As String) As String
    
    Dim strReturn As String
    Dim strConnectionType As String
    strReturn = ""
    strConnectionType = ""
    
    If HasDirectionCheckSegmentLayer Then
        Select Case pstrReturnClass
        Case "Style"
            strReturn = GetConnectionStyle
        Case "Angle"
            strConnectionType = GetConnectionStyle
            If InStr(1, strConnectionType, "Error_") Then
                strReturn = strConnectionType
            Else
                strReturn = GetTurningAngle
            End If
        Case Else
            strReturn = "Error_NotSupportedReturnClass"
        End Select
    Else
        strReturn = "Error_MissingDirectionCheckSegmentLayer"
    End If
    
    GetConnection = strReturn

End Function
Private Function HasDirectionCheckSegmentLayer() As Boolean
    
    Dim lyrObject As Esprit.Layer
    HasDirectionCheckSegmentLayer = False
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
            If InStr(1, .Name, "방향체크") <> 0 Then
                HasDirectionCheckSegmentLayer = True
                Exit For
            End If
        End With
    Next

End Function

Private Function GetConnectionStyle() As String
    
    Dim lyrObject As Esprit.Layer
    Dim strConnectionType As String
    strConnectionType = ""
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If InStr(1, .Name, "방향체크") <> 0 Then
            strConnectionType = Mid(Replace(.Name, "방향체크", ""), 1, 1)
            Select Case strConnectionType
            Case "H"
            Case "K"
            Case "O"
            Case "S"
            Case "X"
            Case "T"
                strConnectionType = strConnectionType
            Case Else
                strConnectionType = "Error_NotInRegeisteredConnectionStyleCode"
            End Select

            Exit For
        End If
        End With
    Next
    
    GetConnectionStyle = strConnectionType

End Function

Private Function GetTurningAngle() As String
    
    Dim lyrObject As Esprit.Layer
    Dim strTurningAngle As String
    strTurningAngle = ""
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If InStr(1, .Name, "방향체크") <> 0 Then
            strTurningAngle = Replace(.Name, "방향체크" + GetConnectionStyle, "")
            If Not IsNumeric(strTurningAngle) Then
                strTurningAngle = "Error_IsNotNumber"
            End If
            Exit For
        End If
        End With
    Next
    
    GetTurningAngle = strTurningAngle

End Function

Public Function IsAlpha(s) As Boolean
    IsAlpha = Len(s) And Not s Like "*[!a-zA-Z]*"
End Function


Public Function CheckPresetFiles(Optional pStrAbutmentType As String = "TA") As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Created By: Ian Pak(Tru IT/SW)
' Created At: 04/05/2023
' Last Updated: 04/05/2023
'
' Parameter
' pStrAbutmentType As String:
' Optional pStrAbutmentType As String = "TA":
' pStrAbutmentType: "TA","ASC","T-L","AOT"
'
' Return Value
' If any file is existed in the PresetFiles folder or not.
' Usage: Call CheckPresetFiles("TA")

    On Error GoTo Err_handler

'Return Value
    Dim rtnValue As Boolean
    rtnValue = False

'Common Values
    Dim c_strPrcPath As String
    c_strPrcPath = Application.Configuration.GetFileDirectory(espFileTypeTemplate) & "\Presetfiles\" & ABUTMENT_TYPE

'Work variables

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Begin
    Application.OutputWindow.Visible = True
    Call Application.OutputWindow.Dock(espToolBarPositionBottom)
    Call Application.OutputWindow.Text("[Preset file path] " & c_strPrcPath & vbCrLf)
    rtnValue = Len(Dir(c_strPrcPath + "/*.*")) > 0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finalize. Release Resources

CheckPresetFilesEND:
    CheckPresetFiles = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Function
