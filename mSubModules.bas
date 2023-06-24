Attribute VB_Name = "mSubModules"
Option Private Module

'Step1
Sub GetSTL(strSTLFilePath As String)
'1. merge STL file in 'STL' Layer
    'Call Document.MergeFile("C:\Users\user\Desktop\기본설정\TEST\Osstem TS,GS Standard ver.1.esp")
    'Document.Refresh
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If (.Name = "STL") Then
            Document.ActiveLayer = lyrObject
        End If
        End With
    Next
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Document.GraphicsCollection.Remove (graphicObject.GraphicsCollectionIndex)
                Call MsgBox("Previous STL model has been deleted.", , "CAM Automation")
                Document.Refresh
            End If
        End If
        End With
    Next
    
    Call Document.MergeFile(strSTLFilePath)
    'Call Document.MergeFile("Select STL File")
    
    
    Document.Refresh

End Sub

Public Function CheckSTLInTheCircle() As Boolean
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
    
    If CheckSTLInTheCircle Then
        Call Document.GraphicsCollection.Remove(cBound.GraphicsCollectionIndex)
    End If
    
    Document.Refresh

'IntersectCircleAndArcsSegments(

End Function

Sub SelectSTL_Model()
'4. move the selected STL feature
    Dim mGraphicObject As Esprit.graphicObject
    Dim i As Long
    
     'Check the existing selection
    'Document.OpenUndoTransaction
    Dim stl As Esprit.STL_Model
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stl = graphicObject
                Call MoveSTL_Step1(stl)
                Document.Refresh
            End If
        End If
        End With
    Next
    
    'Call Document.CloseUndoTransaction(True)
    Set mGraphicObject = Nothing

    Dim nCount As Integer
    'nCount = Document.Group.Count
    'Call Document.SelectionSets.Item(1).Rotate(l, -90, 1)
    nCount = Document.Group.Count
    
    Call Application.OutputWindow.Text("PartStockLength(Before): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    Document.LatheMachineSetup.PartStockLength = Round(GetCutOffXRightEnd(), 2)
    Call Application.OutputWindow.Text("PartStockLength(Updated): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    
'    Dim sitm As Esprit.SelectionSet
''    sitm.Add (Document.Group.Item(1))
'    Set sitm = Document.SelectionSets.Add("STL")
End Sub


'This function will invert the input ptop orientation
Sub MoveSTL_Step1(ByRef stl As Esprit.STL_Model)
    
    If stl Is Nothing Then Exit Sub 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stl)  'Add the stl_model to the selection object
    End With
    
    ' Call mSelection.Translate(5, 0, 0)
    '1. Rotate
    Dim radian As Double
    Dim degree As Double
    Dim radians_angle As Double
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '3)  Y축기준 -90도 회전
    'radian = -90 * PI / 180
    Call TurnSTL("Y,-90")
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    '4)  X축기준 (Hexa: -30도 회전 / Square: - 45도 회전)
    Dim strTurning As String
        
    Dim cCondition As Boolean
    cCondition = True
    
    'strTurning = InputBox("Select (H)exa/(O)cta/(S)quare or input (X,-30)", "CAM Automation - To transform the STL", "H")
    'Updated at 09.09.2021.
    'Auto Turning
    strTurning = "X,0"
    Do While cCondition
        Select Case Left(strTurning, 1)
        Case "H" 'Hexa
            Call TurnSTL("X,-30")
            cCondition = False
        Case "O" 'Octa
            Call TurnSTL("X,-45")
            cCondition = False
        Case "S" 'Square
            Call TurnSTL("X,-45")
            cCondition = False
        Case "X", "Y" 'manual input
            If TurnSTL(strTurning) > 0 Then
                cCondition = False
            Else
                strTurning = InputBox("Select (H)exa/(O)cta/(S)quare or input (X,-30)", "CAM Automation - To transform the STL", "X,-30")
            End If
        Case Else
            strTurning = InputBox("Select (H)exa/(O)cta/(S)quare or input (X,-30)", "CAM Automation - To transform the STL", "X,-30")
        End Select
    Loop
        
        
'    radian = -30 * PI / 180
'    'Call Document.Lines.Add(iPoint, 0, 0, 1)
'    Call mSelection.Rotate(iLine, radian, 0)
   
'    iLine.Grouped = True
'    Document.Lines.Remove (Document.Lines.Count)
'    Document.Points.Remove (Document.Points.Count)
    
    
'    For Each circleObject In Esprit.Document.Circles
'        With circleObject
'        If (.Layer.Number = 0) Then
'            .Grouped = True
'        End If
'        End With
'    Next
'
    If Document.Circles.Count > 0 Then
        Call Document.Circles.Remove(Document.Circles.Count)
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Translate move to aside axis X
    'Set mRefGraphicObject() = Document.FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
   ' Dim mRefGraphicObject() As Esprit.graphicObject = Document.FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    'Dim comCurves() As Object, plottedObjects() As Esprit.graphicObject, faults As EspritComBase.ComFaults
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim graphicObject As Esprit.graphicObject
    Dim dLeftEnd As Double
    dLeftEnd = GetSTLXEnd(0)
    '2. Move X
    Call mSelection.Translate(dLeftEnd * (-1) + 0.1, 0, 0)
    
    Dim r(2) As Double
    
    r(1) = GetCutOffXRightEnd
    r(2) = GetSTLXEnd(1)
        
    Dim nCnt As Integer
    nCnt = 0
    
    'Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        'If (.Layer.Number >= 9 And .Layer.Number <= 11 And (.GraphicObjectType = espSolidModel)) Then
        'If (.Layer.Number >= 9 And .Layer.Number <= 11) Then
        'If ((.Layer.Name = "BACK TURNING" Or .Layer.Name = "CUT-OFF" Or .Layer.Name = "CUF-OFF" Or .Layer.Name = "경계소재-1") And .TypeName <> "Operation") Then
        If ((.Layer.Name = "BACK TURNING" Or .Layer.Name = "CUT-OFF" Or .Layer.Name = "CUF-OFF" Or .Layer.Name = "경계소재-1" Or .Layer.Name = "SPECIAL" Or (InStr(1, .Layer.Name, "방향체크") <> 0)) And .TypeName <> "Operation") Then
            If (.Key > 0) Then
                'Set solidObject = graphicObject
                
                'Call Step2_ConnectionSet(graphicObject, r(1), r(2))
                'nCnt = nCnt + 1
                Call mSelection.Add(graphicObject)  'Add the stl_model to the selection object
                
                
            
            End If
        End If
        End With
    Next
    
    Call mSelection.Translate(r(2) - r(1), 0, 0, 0)
    Document.Refresh
    
    For Each ly In Document.Layers
        'Requested by Kwangho KJNM at 2018.06.
        'If (ly.Name = "BACK TURNING" Or ly.Name = "CUF-OFF" Or ly.Name = "CUT-OFF" Or ly.Name = "경계소재-1" Or ly.Name = "SPECIAL") Then
        If (ly.Name = "BACK TURNING" Or ly.Name = "CUF-OFF" Or ly.Name = "CUT-OFF" Or ly.Name = "경계소재-1" Or (InStr(1, ly.Name, "방향체크") <> 0)) Then
            ly.Visible = True
        End If
    Next
    
    Document.Refresh
    
End Sub
Function Step1_2() As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: Step1_2
' Description: Generate Turning Profile Lines for .
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
    
    Dim rSTLXEnd As Double
    rSTLXEnd = GetSTLXEnd(1)
    If rSTLXEnd = -999 Then
        Call MsgBox("Cannot get the right end of the STL model in STL Layer. Please check it.")
        Exit Function
    End If
    
    Dim strTolerance As String
    'strTolerance = InputBox("Enter Tolerance For Turning Profile.", "CAM Automation - For Turning Profile", "0.0001")
    strTolerance = InputBox("Enter Tolerance For Turning Profile.", "CAM Automation - For Turning Profile", DEFAULT_TOLERANCE)
    
    rtnValue = GetTurningProfile(CDbl(strTolerance))
    If rtnValue > 0 Then
        rtnValue = GetTurningProfile_EditBoundary(rSTLXEnd)
    Else
        GoTo Step1_2END
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finalize. Release Resources

Step1_2END:
    Step1_2 = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
    
End Function

Sub Step2_ConnectionSet(ByRef graphicRef As Esprit.graphicObject, ByVal FromX As Double, ByVal ToX As Double)
    
    If graphicRef Is Nothing Then Exit Sub 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(graphicRef)  'Add the stl_model to the selection object
    End With
    
    Call mSelection.Translate(ToX - FromX, 0, 0, 0)
    
    Document.Refresh
End Sub



Function GetCutOffXRightEnd() As Double
'Get connection width from Layer 10. Cut Off
        
    On Error Resume Next

    GetCutOffXRightEnd = 0

    On Error GoTo 0
    
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    Dim sgmtObject As Esprit.Segment
    Dim sgmtCuttOff(2) As Esprit.Segment
    Dim dCuttOffX(3) As Double
    Dim dReturn As Double
    Dim i As Integer
    
    i = 0
    For Each sgmtObject In Esprit.Document.Segments
        With sgmtObject
        If (.Layer.Name = "CUF-OFF" Or .Layer.Name = "CUT-OFF") Then
            'Document.ActiveLayer = .Layer
            'sgmtObject.Grouped = True
'                sgmtCuttOff(Document.Group.Count) = sgmtObject
            i = i + 1
            dCuttOffX(i) = sgmtObject.XStart
        End If
        End With
    Next
    
    'Get X Right End
    dReturn = 0
    If (dCuttOffX(1) > dCuttOffX(2)) Then
        dReturn = dCuttOffX(1)
    Else
        dReturn = dCuttOffX(2)
    End If
    
    'GetCutOffXRightEnd = dCutOffX(3)
    Document.ActiveLayer = lyrObjectInitial
    
    GetCutOffXRightEnd = dReturn
    Document.Refresh
End Function


Function GetCutOffXEnd(penLR As TruLeftRight) As Double
'Get connection width from Layer 10. Cut Off
        
    On Error Resume Next

    GetCutOffXEnd = 0

    On Error GoTo 0
    
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    Dim sgmtObject As Esprit.Segment
    Dim sgmtCuttOff(2) As Esprit.Segment
    Dim dCuttOffX(3) As Double
    Dim dReturn As Double
    Dim i As Integer
    
    i = 0
    For Each sgmtObject In Esprit.Document.Segments
        With sgmtObject
        If (.Layer.Name = "CUF-OFF" Or .Layer.Name = "CUT-OFF") Then
            'Document.ActiveLayer = .Layer
            'sgmtObject.Grouped = True
'                sgmtCuttOff(Document.Group.Count) = sgmtObject
            i = i + 1
            dCuttOffX(i) = sgmtObject.XStart
        End If
        End With
    Next
    
    'Get X Right End
    dReturn = 0
    
    If penLR = TRU_RIGHT Then
        dReturn = maxValue(dCuttOffX(1), dCuttOffX(2))
    ElseIf penLR = TRU_LEFT Then
        dReturn = minValue(dCuttOffX(1), dCuttOffX(2))
    End If
    
    'GetCutOffXEnd = dCutOffX(3)
    Document.ActiveLayer = lyrObjectInitial
    
    GetCutOffXEnd = dReturn
    Document.Refresh
End Function

Function GetSTLXEnd(nDirectionCode As Integer) As Double
'Get connection width from Layer 10. Cut Off
        
'nDirectionCode = 0 : Left / 1 : Right / else : Error
    On Error Resume Next

    GetSTLXEnd = 0

    On Error GoTo 0
    
    'DirectionCode Check
    If Not (nDirectionCode = 0 Or nDirectionCode = 1) Then
        GoTo EndGetSTLXEnd
    End If
    
    
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim dReturn As Double
    
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
            End If
        End If
        End With
    Next
    
    If stlObject Is Nothing Then GetSTLXEnd = -1 'check if the object is valid
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("Temp")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("Temp")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
    End With
    
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    mRefGraphicObject = Document.FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espFeatureChains)
    For Each ly In Document.Layers
        If (ly.Name = "STL") Then
            mRefGraphicObject(0).Layer = ly
        End If
    Next
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = mRefGraphicObject(0).Layer.Name And (.GraphicObjectType <> espUnknown) And (.GraphicObjectType <> espWorkCoordinate)) Then
            If (.Key > 0) Then
                .Grouped = False
            End If
        End If
        If (.Layer.Name = mRefGraphicObject(0).Layer.Name And (.GraphicObjectType = espFeatureChain)) Then
            If (.Key > 0) Then
                Set returnedFC = graphicObject
            End If
        End If
        End With
    Next
    
    'Call mSelection.Translate(10, 0, 0)
    Dim graphicTemp As Esprit.graphicObject
    Dim lnLine As Esprit.Line
    Dim sgSegment As Esprit.Segment
    Dim sgArc As Esprit.Arc
    
    Dim dLeftEnd As Double
    Dim dRIghtEnd As Double
    Dim i As Integer
    
    dLeftEnd = 0
    dRIghtEnd = 0
    'For Each lnLine In Esprit.Document.Lines
    For i = 1 To returnedFC.Count
        
         If returnedFC.Item(i).TypeName = "Line" Then
            lnLine = returnedFC.Item(i)
            dRIghtEnd = lnLine.x
            dLeftEnd = lnLine.x
         
            'RightEnd
            If lnLine.x > dRIghtEnd Then
                dRIghtEnd = lnLine.x
            End If
            'LeftEnd
            If lnLine.x < dLeftEnd Then
                dLeftEnd = lnLine.x
            End If
         
         ElseIf returnedFC.Item(i).TypeName = "Segment" Then
            Set sgSegment = returnedFC.Item(i)
            'RightEnd
            If sgSegment.XEnd > sgSegment.XStart Then
                If sgSegment.XEnd > dRIghtEnd Then
                    dRIghtEnd = sgSegment.XEnd
                End If
            Else
                If sgSegment.XStart > dRIghtEnd Then
                    dRIghtEnd = sgSegment.XStart
                End If
            End If
            'LeftEnd
            If sgSegment.XStart < sgSegment.XEnd Then
                If sgSegment.XStart < dLeftEnd Then
                    dLeftEnd = sgSegment.XStart
                End If
            Else
                If sgSegment.XEnd < dLeftEnd Then
                    dLeftEnd = sgSegment.XEnd
                End If
            End If
         ElseIf returnedFC.Item(i).TypeName = "Arc" Then
            Set sgArc = returnedFC.Item(i)
            'RightEnd
            If sgArc.Extremity(espExtremityEnd).x > sgArc.Extremity(espExtremityStart).x Then
                If sgArc.Extremity(espExtremityEnd).x > dRIghtEnd Then
                    dRIghtEnd = sgArc.Extremity(espExtremityEnd).x
                End If
            Else
                If sgArc.Extremity(espExtremityStart).x > dRIghtEnd Then
                    dRIghtEnd = sgArc.Extremity(espExtremityStart).x
                End If
            End If
            'LeftEnd
            If sgArc.Extremity(espExtremityStart).x < sgArc.Extremity(espExtremityEnd).x Then
                If sgArc.Extremity(espExtremityStart).x < dLeftEnd Then
                    dLeftEnd = sgArc.Extremity(espExtremityStart).x
                End If
            Else
                If sgArc.Extremity(espExtremityEnd).x < dLeftEnd Then
                    dLeftEnd = sgArc.Extremity(espExtremityEnd).x
                End If
            End If
        
        Else
            MsgBox ("[GetSTLXEND]Not considered type: " & returnedFC.Item(i).TypeName)
        End If
        
    Next
    
    Dim SS As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set SS = .Item("tmpPartProfile")
        On Error GoTo 0
        If SS Is Nothing Then Set SS = .Add("tmpPartProfile")
    End With

    With SS
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(mRefGraphicObject)  'Add the stl_model to the selection object
    End With
    
    Call SS.Delete
    
    'Get X Right End
    'dReturn = returnedFC.BoundingBoxLength + startPoint.X
    'nDirectionCode = 0 : Left / 1 : Right / else : Error
    If nDirectionCode = 0 Then
        dReturn = dLeftEnd
    ElseIf nDirectionCode = 1 Then
        dReturn = dRIghtEnd
    Else
        dReturn = -999 'error
    End If
    'GetCutOffXRightEnd = dCutOffX(3)
    Document.ActiveLayer = lyrObjectInitial
    Document.Refresh
    
EndGetSTLXEnd:
    GetSTLXEnd = dReturn
End Function


Function GetTurningProfile(Optional ByVal bTolerance As Double = 0.1) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: GetTurningProfile
' Description: Generate Turning Profile with STL & "경계소재-1"(.Layer.Number = 11).
' Created By: Ian Pak(Tru IT/SW)
' Created At:
' Last Updated: 03/21/2023
'
' Parameter
' Optional ByVal bTolerance As Double = 0.1:
'
' Usage: rtnValue = GetTurningProfile(CDbl(strTolerance))
    'On Error Resume Next
    On Error GoTo Err_handler

'Return Value
    Dim rtnValue As Integer
    rtnValue = 0

    
    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    'Select Group: STL & 경계소재(11)
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
            End If
        ElseIf (.Layer.Name = "경계소재-1" And (.GraphicObjectType = espSolidModel)) Then
            If (.Key > 0) Then
                Set solidObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
        
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("subGetPartProfile")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("subGetPartProfile")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
        Call .Add(solidObject)  'Add the stl_model to the selection object
    End With
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim radian As Double
    Dim PI As Double
    PI = 3.14159265
    
    radian = 1 * PI / 180
    
    With Document
        .ActiveLayer = .Layers.Item(2) 'Set Turning Profile Layer to Active
        'mRefGraphicObject = .FeatureRecognition.CreateTurningProfile(mSelection, Document.Planes.Item("0DEG"), espTurningProfileOD, espSegmentsArcs, espTurningProfileLocationTop, 0.0001, 0.0001, radian)
        mRefGraphicObject = .FeatureRecognition.CreateTurningProfile(mSelection, Document.Planes.Item("0DEG"), espTurningProfileOD, espSegmentsArcs, espTurningProfileLocationTop, bTolerance, bTolerance, radian)
    End With
    
    Document.Refresh
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance
    rtnValue = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finalize. Release Resources

GetTurningProfileEND:
    GetTurningProfile = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Function

Function GetTurningProfile_EditBoundary(STLRightEndX As Double, Optional pYValue As Double = 1.25) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: GetTurningProfile_EditBoundary
' Description: Automated find the horizontal line in "경계소재-1" and lift it up by pValue.
' ** "FRONT TURNING" = (.Layer.Number = 1)
' Created By: Ian Pak(Tru IT/SW)
' Created At:
' Last Updated: 03/21/2023
'
' Parameter
' STLRightEndX As Double:
' Optional pYValue As Double = 1.25:
'
' Usage: rtnValue = GetTurningProfile_EditBoundary(rSTLXEnd)
    'On Error Resume Next
    'On Error GoTo 0
    On Error GoTo Err_handler

'Return Value
    Dim rtnValue As Integer
    rtnValue = 0
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    Dim segmentObject As Esprit.Segment
    Dim segmentSelected As Esprit.Segment
    
'Common variables
'Work variables
    
    'Layer visible reset
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    'Find the horizontal line in "FRONT TURNING"(.Layer.Number = 1)
        '0: when it fails in the first try, just give it up.
        'n: when it fails in the first try, try it again n times with reduced precision.
    Set segmentSelected = getTheHorizontalLine("FRONT TURNING", 2)

    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        Set mSelection = .Item("GetTurningProfile_EditBoundary")
        If mSelection Is Nothing Then Set mSelection = .Add("GetTurningProfile_EditBoundary")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(segmentSelected)  'Add the segment to the selection object
    End With
    
    'If cannot find the horizontal line in "경계소재-1" stop process and exit. To process it manually.
    If (mSelection.Count = 0) Then
        rtnValue = -991
        MsgBox ("Cannot be found the horizontal line in 경계소재-1. Please check the 경계소재-1, or process it manually.")
        Exit Function
    End If
    
    Call mSelection.Translate(0, pYValue, 0, 0)
    
    rtnValue = 1

GetTurningProfile_EditBoundaryEND:
    Dim ly As Esprit.Layer
    For Each ly In Document.Layers
        If (ly.Name = "FRONT TURNING") Then
            ly.Visible = True
        ElseIf (ly.Number = 1) Then
            ly.Visible = True
        End If
    Next
    Document.Refresh

    GetTurningProfile_EditBoundary = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Function

Function GetPartProfileSTL(Optional ByVal bTolerance As Double = 0.1) As Esprit.graphicObject()
    On Error Resume Next
    On Error GoTo 0
        
    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If
        
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    'Select Group: STL
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
                Set layerObject = .Layer
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("GetPartProfileSTL")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("GetPartProfileSTL")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
    End With
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Dim returnedFC As Esprit.FeatureChain
    
    Dim startPoint As Esprit.Point
    Dim midPoint As Esprit.Point
    Dim endPoint As Esprit.Point
    
    Dim radian As Double
    
    
    With Document
        .ActiveLayer = layerObject 'Set Turning Profile Layer to Active
        GetPartProfileSTL = .FeatureRecognition.CreatePartProfileShadow(mSelection, Document.Planes.Item("0DEG"), espSegmentsArcs)
    End With
    
    Document.Refresh
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance
    
End Function

Function TurnSTL(strParse As String) As Integer

    Document.OpenUndoTransaction
    Dim stl As Esprit.STL_Model
    Dim mSelection As Esprit.SelectionSet
    
    TurnSTL = 0
    
    Dim str() As String
    Dim dDegree As Double
    str = Split(strParse, ",")
    If strParse = "" Then
        Call MsgBox("Please set the parameters properly.")
        TurnSTL = -1
        Exit Function
    End If
    
    If Not (str(0) = "X" Or str(0) = "Y" Or str(0) = "Z") Then
        Call MsgBox("First Letter must be X, Y, or Z.")
        TurnSTL = -901
        Exit Function
    End If
    If Not (IsNumeric(str(1))) Then
        Call MsgBox("Second part must be numeric between -180 ~ 180.")
        TurnSTL = -902
        Exit Function
    End If
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stl = graphicObject
            End If
        End If
        End With
    Next

    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("tmpSTL")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("tmpSTL")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stl)  'Add the stl_model to the selection object
    End With
    
    Call Document.CloseUndoTransaction(True)
    
    
    ' Call mSelection.Translate(5, 0, 0)
    '1. Rotate
    Dim iLine As Esprit.Line
    Dim iPoint As Esprit.Point
    Set iPoint = Document.Points.Add(0, 0, 0)
    
    If str(0) = "X" Then
        Set iLine = Document.Lines.Add(iPoint, 1, 0, 0)
    Else
        If str(0) = "Y" Then
            Set iLine = Document.Lines.Add(iPoint, 0, 1, 0)
        End If
    End If
    
    Dim degree As Double
    degree = CDbl(str(1))
    
    'Rotate by the Axis & Degrees from the parameter
    radian = degree * PI / 180
    Call mSelection.Rotate(iLine, radian, 0)
    
    Document.GraphicsCollection.Remove (iPoint.GraphicsCollectionIndex)
    Document.GraphicsCollection.Remove (iLine.GraphicsCollectionIndex)
    
    Document.Refresh
    TurnSTL = 1
    
End Function


'Step1
Sub CheckSTL(strSTLFilePath As String)
'1. merge STL file in 'STL' Layer
    'Call Document.MergeFile("C:\Users\user\Desktop\기본설정\TEST\Osstem TS,GS Standard ver.1.esp")
    'Document.Refresh
    Dim lyrObjectInitial As Esprit.Layer
    Dim lyrObject As Esprit.Layer
    Set lyrObjectInitial = Document.ActiveLayer
    
    For Each lyrObject In Esprit.Document.Layers
        With lyrObject
        If (.Name = "STL") Then
            Document.ActiveLayer = lyrObject
        End If
        End With
    Next
    
    Dim nCount As Integer
    nCount = 0
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                nCount = nCount + 1
                'Document.GraphicsCollection.Remove (graphicObject.GraphicsCollectionIndex)
                'Call MsgBox("Previous STL model has been deleted.", , "CAM Automation")
                'Document.Refresh
            End If
        End If
        End With
    Next
    
    If nCount = 1 Then
        Exit Sub
    ElseIf nCount > 1 Then
        Call MsgBox("More than 2 STL models are in the document. They is being deleted.", , "CAM Automation")
        For Each graphicObject In Esprit.Document.GraphicsCollection
            With graphicObject
            If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
                If (.Key > 0) Then
                    Document.GraphicsCollection.Remove (graphicObject.GraphicsCollectionIndex)
                    'Call MsgBox("Previous STL model has been deleted.", , "CAM Automation")
                End If
            End If
            End With
        Next
        Call Document.MergeFile(strSTLFilePath)
    Else
        Call Document.MergeFile(strSTLFilePath)
    End If
    
    
    Document.Refresh

End Sub
Function GenerateMarginAreaFeatureChain(pstrWorkPlaneName As String, pstrMarginLayerName As String, Optional pdSmashMinFaceAngle As Double = 20, Optional pdInterval As Double = -1, Optional pstrTempLayerName As String = "TempSmash") As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: GenerateMarginAreaFeatureChain
' Description: Generate Margin Area FeatureChain.
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
    
    
    Dim nResult As Integer
    
    Dim dCutOffXLeftEnd As Double
    Dim dCutOffXRightEnd As Double
    Dim strLayerName As String
    
    Dim tmpSegment As Esprit.Segment
    Dim tmpArc As Esprit.Arc
    
    Dim mSelection As Esprit.SelectionSet
    
    Dim c_strCurrWorkPlane As String
    c_strCurrWorkPlane = Document.ActivePlane.Name
    
    strLayerName = SearchLayerName(pstrTempLayerName)
    If (strLayerName = pstrTempLayerName) Then
        Document.Layers.Remove (pstrTempLayerName)
    End If
    Document.Layers.Add (pstrTempLayerName)
    
    If pstrWorkPlaneName = "" Then
        pstrWorkPlaneName = GetStringDegree(pstrMarginLayerName)
    End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'1. Smash
    Call GetSmash(pstrTempLayerName, pdSmashMinFaceAngle)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'2. Copy to nnDEG layer.

    '2-1. Copy to nnDEG layer.
    nResult = CopyLayer(pstrTempLayerName, pstrMarginLayerName)
    If nResult <= 0 Then
        MsgBox ("Error in copy layer.")
    End If
    
    '2-2. Change active layer to the copied nnDEG 마진 layer
    setLayersFor (pstrMarginLayerName)
    
    '2-3. Align to Workplane for the nnDEG 마진 layer
    'Initialize mSelection
    Document.Group.Clear
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item(pstrMarginLayerName)
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add(pstrMarginLayerName)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    'Move WorkPlane to nnDEG for the nnDEG 마진 layer
    Document.ActivePlane = Document.Planes(pstrWorkPlaneName)
    'Get MajorGraphicObjects into mSelection from the nnDEG 마진 layer
    Set mSelection = SelectMajorGraphicObjectsInLayer(pstrMarginLayerName)
    'Align mSelection to nnDEG for the nnDEG 마진 layer
    Call mSelection.AlignPlane(Document.Planes(pstrWorkPlaneName), espAlignPlaneToGlobalXYZ)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'3. Trim1: Right side of Connection leftend
'4. Trim2: Z<0

    dCutOffXLeftEnd = GetCutOffXEnd(TRU_LEFT) - 0.5

    For Each goObject In Esprit.Document.GraphicsCollection
        With goObject
        If (.Layer.Name = pstrMarginLayerName) Then
            If (.GraphicObjectType = espSegment) Then
                Set tmpSegment = goObject
                With tmpSegment
                    If (minValue(.XStart, .XEnd) > dCutOffXLeftEnd) Then
                        .Grouped = True
                    ElseIf (maxValue(.ZStart, .ZEnd) <= 0) Then
                        .Grouped = True
                    Else
                        .Grouped = False
                    End If
                End With
            ElseIf (.GraphicObjectType = espArc) Then
                Set tmpArc = goObject
                With tmpArc
                    If (minValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) > dCutOffXLeftEnd) Then
                        .Grouped = True
                    ElseIf (maxValue(.Extremity(espExtremityStart).Z, .Extremity(espExtremityEnd).Z) <= 0) Then
                        .Grouped = True
                    Else
                        .Grouped = False
                    End If
                End With
            Else
                '.Grouped = False
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    'Document.Refresh
    Call Document.Group.DeleteAll
    'Call Document.GraphicsCollection.Remove(tmpSegment.GraphicsCollectionIndex)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'5. Find Top(X,Y,+), Bottom(X,Y,+)
    Dim dTop As Double
    Dim dBottom As Double
    Dim i As Integer
    
    Dim pntTop As Esprit.Point
    Dim pntBottom As Esprit.Point
    Set pntTop = Nothing
    
    '1) Get TopPoint
    Set pntTop = GetTopBottomPoint(TRU_TOP, pstrMarginLayerName)
    '2) Get BottomPoint
    Set pntBottom = GetTopBottomPoint(TRU_BOTTOM, pstrMarginLayerName)
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'6. Trim3: Trim exclude the Margin Line(connected elements with the pntTop Point)
    
    '1) To Get the Margin Line (connected elements with the pntTop Point)
    Dim Geometry As Esprit.graphicObject, TypeMask As Variant
    Set Geometry = pntTop
    
    Dim PropagationResults() As Object
    PropagationResults = PropagateGeometry(Geometry)
    Document.Group.Clear
    
    Debug.Print vbNewLine & "Propagation Results:"
    Debug.Print "#", "Type", "Key"
    Dim Element As Esprit.graphicObject
    For i = LBound(PropagationResults) To UBound(PropagationResults)
        Set Element = PropagationResults(i)
        Debug.Print i, Element.GuiTypeName, Element.Key
        Call Document.Group.Add(Element)
    Next
    Document.Refresh
    
    With Document.SelectionSets
        Set mSelection = .Item(pstrMarginLayerName)
        If mSelection Is Nothing Then Set mSelection = .Add(pstrMarginLayerName)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    If Document.Group.Count > 0 Then
        For i = 1 To Document.Group.Count
            Call mSelection.Add(Document.Group.Item(i))
        Next
    End If
    Document.Group.Clear
   
    '2) Trim elements excluding the Margin Line (connected elements with the pntTop Point)
    For Each goObject In Esprit.Document.GraphicsCollection
    With goObject
        If ((.Layer.Name = pstrMarginLayerName) _
            And (.GraphicObjectType = espSegment Or .GraphicObjectType = espArc Or .GraphicObjectType = espPoint)) Then
            If (Not mSelection.Contains(goObject)) Then
                Document.Group.Add (goObject)
                'Document.Refresh
            End If
        End If
    End With
    Next
    'Document.Refresh
    Document.Group.DeleteAll
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'7. Trim4: REMOVE Z<MAXValue(pntTop.Z, pntBottom.Z)
'
    Dim dLeftEndofMarginLine As Double
    dLeftEndofMarginLine = minValue(pntTop.x, pntBottom.x)
    For Each goObject In Esprit.Document.GraphicsCollection
        With goObject
        If (.Layer.Name = pstrMarginLayerName) Then
            If (.GraphicObjectType = espSegment) Then
                Set tmpSegment = goObject
                With tmpSegment
                    'If (maxValue(.XStart, .XEnd) < dLeftEndofMarginLine) Then
                    If (False) Then
                        .Grouped = True
                    ElseIf (maxValue(.YStart, .YEnd) > 0 And maxValue(.ZStart, .ZEnd) <= pntTop.Z) Then
                        .Grouped = True
                        'Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    ElseIf (minValue(.YStart, .YEnd) < 0 And maxValue(.ZStart, .ZEnd) <= pntBottom.Z) Then
                        .Grouped = True
                        'Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    Else
                        .Grouped = False
                    End If
                End With
            ElseIf (.GraphicObjectType = espArc) Then
                Set tmpArc = goObject
                With tmpArc
                    'If (maxValue(.Extremity(espExtremityStart).x, .Extremity(espExtremityEnd).x) < dLeftEndofMarginLine) Then
                    If (False) Then
                        .Grouped = True
                    ElseIf (maxValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) > 0 _
                        And maxValue(.Extremity(espExtremityStart).Z, .Extremity(espExtremityEnd).Z) <= pntTop.Z) Then
                        .Grouped = True
                        'Call Document.GraphicsCollection.Remove(.GraphicsCollectionIndex)
                    ElseIf (minValue(.Extremity(espExtremityStart).y, .Extremity(espExtremityEnd).y) < 0 _
                        And maxValue(.Extremity(espExtremityStart).Z, .Extremity(espExtremityEnd).Z) <= pntBottom.Z) Then
                        .Grouped = True
                    Else
                        .Grouped = False
                    End If
                End With
            Else
                '.Grouped = False
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    'Document.Refresh
    Call Document.Group.DeleteAll
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'8. Group rest of the segments and the arcs
    For Each goObject In Esprit.Document.GraphicsCollection
        With goObject
        If (.Layer.Name = pstrMarginLayerName) Then
            If (.GraphicObjectType = espSegment) Then
                .Grouped = True
            ElseIf (.GraphicObjectType = espArc) Then
                .Grouped = True
            Else
                '.Grouped = False
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
    'Document.Refresh

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'9. Generate Segments
    Dim pntPoint(4) As Point
    Dim sgSegment As Esprit.Segment
    
    Set pntPoint(0) = Document.Points.Add(-2 - pdInterval, pntTop.y + 2, 0)
    Set pntPoint(1) = Document.Points.Add(pntTop.x, pntTop.y + 2, pntTop.Z)
    Set pntPoint(2) = Document.Points.Add(pntBottom.x, pntBottom.y - 2, pntBottom.Z)
    Set pntPoint(3) = Document.Points.Add(-2 - pdInterval, pntBottom.y - 2, 0)
    
    Document.Segments.Add(pntPoint(0), pntPoint(1)).Grouped = True
    Document.Segments.Add(pntPoint(1), pntTop).Grouped = True
    Document.Segments.Add(pntBottom, pntPoint(2)).Grouped = True
    Document.Segments.Add(pntPoint(2), pntPoint(3)).Grouped = True
    Document.Segments.Add(pntPoint(0), pntPoint(3)).Grouped = True
    'Document.Refresh

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'10. Move LEFT X:1
'11. Align From nnDEG

    With Document.SelectionSets
        Set mSelection = .Item(pstrMarginLayerName)
        If mSelection Is Nothing Then Set mSelection = .Add(pstrMarginLayerName)
    End With
    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
    End With
    
    If Document.Group.Count <= 0 Then
        MsgBox ("Cannot find the smashed segments and arcs group")
        rtnValue = -991
        GoTo GenerateMarginAreaFeatureChainEND
    End If
    
    For i = 1 To Document.Group.Count
        Call mSelection.Add(Document.Group.Item(i))
    Next
    
    '10. Move LEFT X:1
    Call mSelection.Translate(pdInterval, 0, 0, 0)
    '11. Align From nnDEG
    Call mSelection.AlignPlane(Document.Planes(pstrWorkPlaneName), espAlignPlaneFromGlobalXYZ)
    'Document.Refresh

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'12. Create FeatureChain
    Dim fcChain As Esprit.FeatureChain
    Set fcChain = Nothing
    
    If mSelection.Count > 0 Then
        GraphicObj = Document.FeatureRecognition.CreateAutoChains(mSelection)
    End If
    
    If UBound(GraphicObj) > 0 Then
        MsgBox ("More than 1 FeatureChains are found. Please check it.")
        rtnValue = -995
        GoTo GenerateMarginAreaFeatureChainEND
    End If
    Set fcChain = GraphicObj(0)
    
    If fcChain Is Nothing Then
        rtnValue = -999
    Else
        rtnValue = 1
        Document.Layers.Remove (pstrTempLayerName)
        fcChain.Grouped = True
        Document.ActivePlane = Document.Planes(pstrWorkPlaneName)
        Document.Refresh
    End If

GenerateMarginAreaFeatureChainEND:
    GenerateMarginAreaFeatureChain = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description

End Function

Function GetSmash(pstrLayerName As String, Optional ByVal pdMinFaceAngle As Double = 20, Optional ByVal pstrAlignPlaneName As String = "0DEG", Optional ByVal pdTolerance As Double = 0.01) As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Name: GetSmash
' Description: Generate Smash with STL.
' Created By: Ian Pak(Tru IT/SW)
' Created At:
' Last Updated: 04/11/2023
'
' Parameter
' Optional ByVal bTolerance As Double = 0.01:
'
' Usage: rtnValue = GetSmash(CDbl(strTolerance))
    'On Error Resume Next
    On Error GoTo Err_handler

'Return Value
    Dim rtnValue As Integer
    rtnValue = 0

    
    'Tolerance set to 0.1 (default 공차)
    Dim bOriTolerance As Double
    bOriTolerance = Application.Configuration.ConfigurationFeatureRecognition.Tolerance
    If Document.SystemUnit = espInch Then
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance * 3.9
    Else
        Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bTolerance
    End If
        
    Dim layerObject As Esprit.Layer
    Dim graphicObject As Esprit.graphicObject
    Dim stlObject As Esprit.STL_Model
    Dim solidObject As Esprit.Solid
    
    'Select Group: STL
    For Each layerObject In Esprit.Document.Layers
        layerObject.Visible = False
    Next
    Document.Refresh
    
    For Each graphicObject In Esprit.Document.GraphicsCollection
        With graphicObject
        If (.Layer.Name = "STL" And (.GraphicObjectType = espSTL_Model)) Then
            If (.Key > 0) Then
                Set stlObject = graphicObject
                .Grouped = True
                .Layer.Visible = True
            End If
        ElseIf (.GraphicObjectType <> espUnknown) Then
            .Grouped = False
        End If
        End With
    Next
        
    setLayersFor (pstrLayerName)
    
    Dim mSelection As Esprit.SelectionSet
    With Document.SelectionSets
        On Error Resume Next
        Set mSelection = .Item("subGetSmash")
        On Error GoTo 0
        If mSelection Is Nothing Then Set mSelection = .Add("subGetSmash")
    End With

    With mSelection
        .RemoveAll
        .AddCopiesToSelectionSet = False
        Call .Add(stlObject)  'Add the stl_model to the selection object
        Call .Add(solidObject)  'Add the stl_model to the selection object
    End With
    
    Dim mRefGraphicObject() As Esprit.graphicObject
    Document.ActivePlane = Document.Planes(pstrAlignPlaneName) 'Must be 0DEG for align to XYZ
    Call mSelection.Smash(True, False, False, espWireFrameElementAll, pdTolerance, pdMinFaceAngle)
    Call mSelection.AlignPlane(Document.Planes(pstrAlignPlaneName), espAlignPlaneToGlobalXYZ)
    

    
    
    
    Document.Refresh
    'Tolerance back to original value
    Application.Configuration.ConfigurationFeatureRecognition.Tolerance = bOriTolerance
    rtnValue = 1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Finalize. Release Resources

GetSmashEND:
    GetSmash = rtnValue
    Exit Function

Err_handler:
    MsgBox Err.Number & "-" & Err.Description
End Function


Sub ReorderOperation()
    Dim Op As Esprit.Operation
    Dim i As Long
    Dim strOperationName() As String
    
    Call Application.OutputWindow.Text("PartStockLength(Before): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    Document.LatheMachineSetup.PartStockLength = Round(GetCutOffXRightEnd(), 2)
    Call Application.OutputWindow.Text("PartStockLength(Updated): " & CStr(Document.LatheMachineSetup.PartStockLength) & vbCrLf)
    
    Call DeleteDummyOperation
    With Document.Operations
        ReDim strOperationName(1 To .Count)
        For i = 1 To .Count
            Set Op = .Item(i)
            Call SetCustomLong(Op.CustomProperties, "SortOrder", i)
            strOperationName(i) = .Item(i).Name
            'MsgBox ("A[" & CStr(I) & "]" & strOperationName(I))
        Next i
    End With
    
    Call SelectionSortStrings(strOperationName())
    For i = 1 To Document.Operations.Count
        If (UBound(filter(strOperationName, Op.Name)) > -1) Then
            'MsgBox ("B[" & CStr(I) & "]" & strOperationName(I))
            Call SetCustomLong(getOperationByName(strOperationName(i)).CustomProperties, "SortOrder", i)
        End If
    Next i

    Call RestoreLastSavedOperationOrder

End Sub

Public Sub showAdvancedNCCode()
    Call Document.GUI.NCCodeAdvanced(True)
End Sub

