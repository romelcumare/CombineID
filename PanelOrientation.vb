Option Explicit
Option Compare Text
Const pi As Double = 3.14159265358979
Const Precision As Integer = 4



Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim swPart                  As SldWorks.PartDoc
Dim swAssy                  As SldWorks.AssemblyDoc
Dim swComp                  As SldWorks.Component2
Dim swSelData               As SldWorks.SelectData
Dim swBody                  As SldWorks.Body2
Dim vBodies                 As Variant
Dim longstatus              As Long
Dim boolstatus              As Boolean
Dim AssemblyPath            As String
Dim swAssyDocExt            As SldWorks.ModelDocExtension
Dim AssemblyName            As String

Dim DocName                 As String
Dim DocPath                 As String
Dim DocumentName            As String

Dim fso                     As FileSystemObject
Dim fileStream              As TextStream


Public CurrentProgress      As Double
Public BarWidth             As Long
Public ProgressPercentage   As Double
Public Progress             As Integer

Public CancelButton         As Boolean

Dim vAssemblyComps          As Variant
Dim ModelName               As String 'Model being processed

Sub main()


    Dim swConfMgr               As SldWorks.ConfigurationManager
    Dim swConf                  As SldWorks.Configuration
    Dim swRootComp              As SldWorks.Component2
    Dim PartList                As Variant
    Dim i                       As Integer
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swAssyDocExt = swModel.Extension
    AssemblyName = FileName(swModel.GetTitle())
         
    
    If Not swModel.GetType = SwConst.swDocASSEMBLY Then
        MsgBox "The active document is not an assembly model.", vbOKOnly
        End
    End If
    
    'On Error GoTo ErrorHandler
    
    'Create Log File
    DocName = swModel.GetTitle()
    DocPath = swModel.GetPathName
        
    If Right(DocName, 7) = ".sldasm" Then
        DocName = Left(DocName, InStrRev(DocName, ".") - 1)
    End If
    
    DocPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") - 1)
    DocPath = DocPath + "\" + DocName + "-LogFile.txt"
    Set fso = New FileSystemObject
    Set fileStream = fso.OpenTextFile(DocPath, ForAppending, True)
    fileStream.WriteLine ""
    fileStream.WriteLine "----------- " & Now & " -----------"
  
    
    Set swAssy = swModel
    
    AssemblyPath = swModel.GetPathName
    
'    CHAAAANGED THIS BACK!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
    'Deactivate view update
'    Dim modView As ModelView
'    Set modView = swModel.ActiveView
'    modView.EnableGraphicsUpdate = False
'    swAssy.FeatureManager.EnableFeatureTree = False
    
        
    'Get components for transformations
    vAssemblyComps = swAssy.GetComponents(False)
    
    'Show Progress Bar
    Progress = 0
    
    With ProgressBar
        .Bar.Width = 0
        .Text.Caption = "Getting Parts..."
        .Text2.Caption = "0% Complete"
        .Text3.Caption = "Processing..."
        .Show vbModeless
    End With
    
    CancelButton = False
    
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration
    Set swRootComp = swConf.GetRootComponent3(True)
            
    PartList = GetParts(swAssy)
    
               
    For i = 0 To UBound(PartList)
    
        MoveBody (PartList(i))
        Call UpdateBar(UBound(PartList), "Processing...")
        
        If CancelButton = True Then
            Debug.Print "CANCEL IS TRUE"
            Debug.Print ""
            Debug.Print ""
            Debug.Print ""
            Debug.Print ""
            Debug.Print ""
            Debug.Print ""
            fileStream.WriteLine "User Cancelled Macro"
            GoTo ExitCode
        End If
        
        
    Next i
    
 
ExitCode:
Call UpdateBar(UBound(PartList), "Rebuilding...")

Set swModel = swApp.ActivateDoc3(AssemblyPath, False, 1, 0)
swModel.ClearSelection2 (True)

boolstatus = swModel.Extension.Rebuild(1)

'    CHAAAANGED THIS BACK!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
''Activate view update
'modView.EnableGraphicsUpdate = True
'swAssy.FeatureManager.EnableFeatureTree = True

Unload ProgressBar


Debug.Print ""
Debug.Print "----------- Macro Finished -----------"
fileStream.WriteLine "----------- Macro Finished -----------"
'Close log file
fileStream.Close

Exit Sub

ErrorHandler:

'    CHAAAANGED THIS BACK!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!11
''Activate view update
'modView.EnableGraphicsUpdate = True
'swAssy.FeatureManager.EnableFeatureTree = True

Unload ProgressBar

Debug.Print ""
Debug.Print "----------- Macro Terminated while processing " & ModelName & "-----------"
fileStream.WriteLine "----------- Macro Terminated while processing " & ModelName & "-----------"
'Close log file
fileStream.Close

End Sub

Sub MoveBody(ByRef model As String)

Dim Component                       As Object
Dim swMoveCopyBodyFeatureData       As Object
Dim swFeature                       As Object
Dim swModelDocExt                   As ModelDocExtension
Dim MyMassProp                      As SldWorks.MassProperty
Dim swModelDoc                      As SldWorks.ModelDoc2
Dim swCustProp                      As CustomPropertyManager
Dim Axis                            As Variant
Dim Xaxis(2)                        As Variant
Dim LAxis                           As Double
Dim LXaxis                          As Double
Dim Angle                           As Double
Dim swFeatMgr                       As SldWorks.FeatureManager
Dim swFeat                          As SldWorks.Feature
Dim DotProd                         As Double
Dim vBodies                         As Variant
Dim swBody                          As SldWorks.Body2
Dim swOriginalBody                  As SldWorks.Body2
Dim swEntity                        As SldWorks.Entity
Dim swSelMgr                        As SldWorks.SelectionMgr
Dim swSelData                       As SldWorks.SelectData
Dim swSafeEnt                       As SldWorks.Entity
Dim vNorm                           As Variant
Dim longstatus                      As Long
Dim longwarnings                    As Long
Dim boolstatus                      As Boolean
Dim vBoundBox                       As Variant
Dim swXform                         As SldWorks.MathTransform
Dim n                               As Integer
Dim MassProperties                  As Object

'Debug.Print "Processing:", Mid(Model, InStrRev(Model, "\") + 1)

Set swModel = swApp.OpenDoc6(model, 1, 0, "", longstatus, longwarnings)

ModelName = Mid(model, InStrRev(model, "\") + 1)

'Debug.Print ModelName


'Check if document is a Part file
If swModel Is Nothing Then 'If nothing means that openDoc didnt work because is set to open parts.
    Exit Sub
End If

Set swPart = swModel
Set swSelMgr = swModel.SelectionManager
Set swSelData = swSelMgr.CreateSelectData
Set swModelDoc = swModel
Set swModelDocExt = swModelDoc.Extension

Set swCustProp = swModelDocExt.CustomPropertyManager("")

'Check if it's a hardware component
If swCustProp.Get("IS_HARDWARE") = "Yes" Or swCustProp.Get("FLIP") = "Yes" Or InStr(model, "\Hardwares\") <> 0 Then
    Debug.Print "Hardware component: " & ModelName
    fileStream.WriteLine "Hardware component: " & ModelName
    Exit Sub
End If


swPart.ClearSelection2 True

'Get Bodies
vBodies = swPart.GetBodies2(swAllBodies, True)

Dim TempVolume                      As Double
Dim BiggestBodyIndex                As Integer
Dim MassProps                       As Variant

'Get body properties to find biggest body
TempVolume = 0

For n = 0 To UBound(vBodies)

    'Debug.Print vBodies(n).Name
    Set swBody = vBodies(n)
    MassProps = swBody.GetMassProperties(1)
    'Debug.Print "Body Name: " & vBodies(n).Name & " Index: " & n & " Volume: " & MassProps(3)
    
    If MassProps(3) > TempVolume Then
    
        TempVolume = MassProps(3)
        BiggestBodyIndex = n
    End If
          
Next n

'Get Biggest Body
Set swBody = vBodies(BiggestBodyIndex)
'Debug.Print "Biggest Body is: " & swBody.Name

'Copy original body - Used for transformation
Set swOriginalBody = swBody.Copy

Set swEntity = SelectLargestFace(swBody)

If swEntity Is Nothing Then
    Debug.Print ModelName & ": Does not contain a flat face "
    fileStream.WriteLine ModelName & ": Does not contain a flat face "
    Exit Sub
End If

'Save Face to be used in next funtion
Set swSafeEnt = swEntity.GetSafeEntity

vNorm = swEntity.Normal
Debug.Print "Normal = (" & Round(vNorm(0), 6) & ", " & Round(vNorm(1), 6) & ", " & Round(vNorm(2), 6) & ")"

'FIRST ROTATION - Find largest face and mate it with Front Plane
If Not (Round(Abs(vNorm(0)), 6) = 0 And Round(Abs(vNorm(1)), 6) = 0 And Round(Abs(vNorm(2)), 6) = 1) Then 'Check if largest flat face is normal to Z

    'Add MoveCopyBody Feature
    For n = 0 To UBound(vBodies)     'Select all bodies
        boolstatus = swPart.Extension.SelectByID2(vBodies(n).Name, "SOLIDBODY", 0, 0, 0, True, 1, Nothing, 0)
        'Debug.Print "Selected Body: " & boolstatus
    Next n
    
    Set swFeature = swPart.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, 1)
    Set swMoveCopyBodyFeatureData = swFeature.GetDefinition
    boolstatus = swMoveCopyBodyFeatureData.AccessSelections(swModel, Nothing)
            
    swSelData.Mark = 1
    boolstatus = swSafeEnt.Select4(False, swSelData)
    
    ' Select plane normal to Z (1=Z, 2=Y, 3=X)
    Dim PlaneName As String
    
    PlaneName = SelectPlane(swModel, 1)
    
'    ########################################
    boolstatus = swPart.Extension.SelectByID2(PlaneName, "PLANE", 0, 0, 0, True, 1, Nothing, 0)

    If Not (boolstatus) Then
        boolstatus = swPart.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
    End If

    If Not (boolstatus) Then
        Debug.Print ""
        Debug.Print "####### Unable selection plane normal to Z in " & ModelName & " plane name must be 'Front' or 'Front Plane' #######"
        Debug.Print ""
        fileStream.WriteLine "####### Unable selection plane normal to Z in " & ModelName & " plane name must be 'Front' or 'Front Plane' #######"
    End If
'    ########################################
 
    swMoveCopyBodyFeatureData.AddMate Nothing, 0, 1, 0, 0, longstatus
    
    'Modify MoveCopyBody Definition
    swFeature.ModifyDefinition swMoveCopyBodyFeatureData, swPart, Component
    swMoveCopyBodyFeatureData.ReleaseSelectionAccess
    swModel.ViewZoomtofit2
    
Else
    fileStream.WriteLine "Already aligned with Z: " & ModelName
End If

'============================================================================================================================'

'SECOND ROTATION - Find longest edge with an adjacent perpendicular edge
Dim LongestEdge         As SldWorks.Edge
Dim swCurve             As SldWorks.Curve
Dim vCurveParam         As Variant
Dim EdgeLength          As Double
Dim EdgeVector(2)       As Double

Set LongestEdge = GetLongestEdge(swSafeEnt)

If Not LongestEdge Is Nothing Then

    Set swCurve = LongestEdge.GetCurve
    vCurveParam = LongestEdge.GetCurveParams2
    EdgeLength = swCurve.GetLength3(vCurveParam(6), vCurveParam(7))
    
    Dim i As Integer
    
    Set swSelMgr = swModel.SelectionManager
    Set swSelData = swSelMgr.CreateSelectData
    
        For i = 0 To 2
            EdgeVector(i) = (vCurveParam(i + 3) - vCurveParam(i)) / EdgeLength
        Next i
    
        'Obtain Angle
        Dim LengthAxis          As Double
        Dim LengthXAxis         As Double
    
        'Xaxis Definition
        Xaxis(0) = 1#
        Xaxis(1) = 0#
        Xaxis(2) = 0#
        
        DotProd = DotProduct(EdgeVector, Xaxis)
        
        LengthAxis = LengthVector(EdgeVector)
        LengthXAxis = LengthVector(Xaxis)
        
        Angle = (DotProd / (LengthAxis * LengthXAxis))
        Angle = Arccos(Angle)
        'Debug.Print "Angle with X is: " & Round(Angle * 180 / pi, Precision)
        
End If
    

If Not (Round(Angle * 180 / pi, Precision) = 0 Or Round(Angle * 180 / pi, Precision) = 180) Then

    'Get Bodies
    vBodies = swPart.GetBodies2(swAllBodies, True)
   
    'Flip rotation angle depending on orientation of edge
    If EdgeVector(1) < 0 Then
        Angle = -Angle
    End If
      
    'Select Bodies and add rotation
    For n = 0 To UBound(vBodies)     'Select all bodies
        boolstatus = swPart.Extension.SelectByID2(vBodies(n).Name, "SOLIDBODY", 0, 0, 0, True, 1, Nothing, 0)
    Next n
    
    Set swFeature = swPart.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, -Angle, 0, 0, False, 1)

End If

If LongestEdge Is Nothing Then

    Debug.Print ModelName & ": Does not contain a straight edge"
    fileStream.WriteLine ModelName & ": Does not contain a straight edge"
    
    ' Create mass properties such as axes of inertia
    Set MyMassProp = swModelDocExt.CreateMassProperty
        
    Axis = MyMassProp.PrincipleAxesOfInertia(0)
    
    'Xaxis Definition
    Xaxis(0) = 1
    Xaxis(1) = 0
    Xaxis(2) = 0
    
    DotProd = DotProduct(Axis, Xaxis)
    
    LAxis = LengthVector(Axis)
    LXaxis = LengthVector(Xaxis)
    
    Angle = (DotProd / (LAxis * LXaxis))
    
    Angle = Arccos(Angle)
        
    'Apply rotation if angle is different than zero
    If Round(Angle * 180 / pi, 2) <> 0 Then
    
        'Angle Precision set to 2 decimals
        'Debug.Print "Angle is: " & Round(Angle * 180 / pi, 10)
        
        'Get Bodies
        vBodies = swPart.GetBodies2(swAllBodies, True)
                   
        'Flip rotation angle depending on orientation of X axis of inertia
        If Axis(1) > 0 Then
            Angle = -Angle
        End If
                
        'Select Bodies and add rotation
        For n = 0 To UBound(vBodies)     'Select all bodies
            boolstatus = swPart.Extension.SelectByID2(vBodies(n).Name, "SOLIDBODY", 0, 0, 0, True, 1, Nothing, 0)
        Next n
        
        Set swFeature = swPart.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, Angle, 0, 0, False, 1)
    End If
End If

'Roll back and Roll to end. Added to fix issue since parts seem to be rolled back
Set swFeatMgr = swModelDoc.FeatureManager
boolstatus = swFeatMgr.EditRollback(swMoveRollbackBarToPreviousPosition, "")
boolstatus = swFeatMgr.EditRollback(swMoveRollbackBarToEnd, "")

'============================================================================================================================'

'GET TRANSFORMATION MATRIX - Move components back into position after applying move/copy body feature

'Get Bodies
vBodies = swPart.GetBodies2(swAllBodies, True)

Set swModel = swPart

'Get Biggest Body
TempVolume = 0

For n = 0 To UBound(vBodies)

    Set swBody = vBodies(n)
    MassProps = swBody.GetMassProperties(1)
    'Debug.Print "Body Name: " & vBodies(n).Name & " Index: " & n & " Volume: " & MassProps(3)
    
    If MassProps(3) > TempVolume Then
    
        TempVolume = MassProps(3)
        BiggestBodyIndex = n
    End If
      
Next n

Set swBody = vBodies(BiggestBodyIndex)
'Debug.Print "Transformed Body: " & swBody.Name

    
Set swXform = GetClosestTransform(swBody, swOriginalBody)

If Not swXform Is Nothing Then 'In case of coincidence not found

    Set swXform = swXform.Inverse

    '=================================== Apply the transformations ==================================='
    'Dim swAssem As SldWorks.AssemblyDoc
    'Repointing to swAssy since it sometimes fails
    'Set swAssy = swApp.OpenDoc6(AssemblyPath, swDocumentTypes_e.swDocASSEMBLY, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
    
    Dim swAssComp           As SldWorks.Component2
    Dim swCompT             As SldWorks.Component2
    Dim swCompXform         As SldWorks.MathTransform
    Dim PreFixed            As Boolean
    
    For i = 0 To UBound(vAssemblyComps)
    
        Set swAssComp = vAssemblyComps(i)
        
        'Check if part is the same as component
        If model = swAssComp.GetPathName Then
    
            'Debug.Print "Is Fixed? " & swAssComp.IsFixed
            'Debug.Print vAssemblyComps(i).Name
                    
            'Debug.Print Model
            
            PreFixed = False
    
            If swAssComp.IsFixed Then
                'Float Component - Only needed if multibody part is opened
                boolstatus = swAssyDocExt.SelectByID2(vAssemblyComps(i).Name & "@" & AssemblyName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                swAssy.UnfixComponent
                PreFixed = True
            End If
    
            Set swCompT = swAssy.GetComponentByName(vAssemblyComps(i).Name)
            
            If Not swCompT Is Nothing Then
                'Create Component Transform
                Set swCompXform = swCompT.Transform2
        
        
                ''Create vCompXfm only to print the transformation
                'Dim vCompXfm As Variant
                'vCompXfm = swCompXform.ArrayData
        
                'Debug.Print ""
                'Debug.Print "COMPONENT ASSEMBLY TRANSFORM", vAssemblyComps(i).Name
                'Debug.Print "Rotation:"
                'Debug.Print vbTab & Round(vCompXfm(0) * 1000, 4), Round(vCompXfm(1) * 1000, 4), Round(vCompXfm(2) * 1000, 4)
                'Debug.Print vbTab & Round(vCompXfm(3) * 1000, 4), Round(vCompXfm(4) * 1000, 4), Round(vCompXfm(5) * 1000, 4)
                'Debug.Print vbTab & Round(vCompXfm(6) * 1000, 4), Round(vCompXfm(7) * 1000, 4), Round(vCompXfm(8) * 1000, 4)
                'Debug.Print "Translation:"
                'Debug.Print vbTab & Round(vCompXfm(9) * 1000, 4), Round(vCompXfm(10) * 1000, 4), Round(vCompXfm(11) * 1000, 4)
                'Debug.Print "Scaling: " & vCompXfm(12)
        
                'Multiply part and component transformations
                swCompT.Transform2 = swXform.Multiply(swCompXform)
            Else
                Debug.Print "TRANSFORM SKIPPED for: " & vAssemblyComps(i).Name
            End If
            
            'Fix Component
            If PreFixed Then
                boolstatus = swAssyDocExt.SelectByID2(vAssemblyComps(i).Name & "@Adam", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
                swAssy.FixComponent
            End If
    
        End If
    
    Next i

End If


'============================================================================================================================'

'ADD BOUNDING BOX - Obtain bounding box of panel to add Length, Width and Thickness custom properties

vBoundBox = swPart.GetPartBox(True)

Dim Length      As Double
Dim Width       As Double
Dim Thickness   As Double

Length = Round(Str(vBoundBox(3) * 1000#) - Str(vBoundBox(0) * 1000#), 2)
Width = Round(Str(vBoundBox(4) * 1000#) - Str(vBoundBox(1) * 1000#), 2)
Thickness = Round(Str(vBoundBox(5) * 1000#) - Str(vBoundBox(2) * 1000#), 2)

'Debug.Print "Bounding Box Size = X:" & Length & ", Y:" & Width & ", Z:" & Thickness

'Add custom property
Set swCustProp = swModelDocExt.CustomPropertyManager("")

boolstatus = swCustProp.Add3("Length", swCustomInfoText, Length, swCustomPropertyReplaceValue)
boolstatus = swCustProp.Add3("Width", swCustomInfoText, Width, swCustomPropertyReplaceValue)
boolstatus = swCustProp.Add3("Thickness", swCustomInfoText, Thickness, swCustomPropertyReplaceValue)

'Flag model as dirty
swModel.SetSaveFlag
                    
'Debug.Print swModel.GetPathName
'Debug.Print " is it Dirty? " & swModel.GetSaveFlag
'swModel.Save2 (True)

End Sub

Public Function SelectLargestFace(swBody As SldWorks.Body2) As SldWorks.Entity
     
    Dim swSelData           As SldWorks.SelectData
    Dim swFace              As SldWorks.Face
    Dim swLargestFace       As SldWorks.Face
    Dim FaceArea            As Double
    Dim TempFaceArea        As Double
    Dim swSurf              As SldWorks.Surface

    
    Set swFace = swBody.GetFirstFace
    
    swModel.ClearSelection2 True
    
    Do While Not swFace Is Nothing

        FaceArea = swFace.GetArea
        
        'Check if face is planar
        Set swSurf = swFace.GetSurface
    
        If swSurf.Identity = 4001 Then '4001 = surface is planar

            If FaceArea > TempFaceArea Then
                Set swLargestFace = swFace
                TempFaceArea = FaceArea
            End If
        
        End If
               
        Set swFace = swFace.GetNextFace

    Loop
        
    Set SelectLargestFace = swLargestFace

End Function

Function DotProduct(Array1 As Variant, Array2 As Variant) As Double
    Dim i As Integer
    Dim temp As Double
    
    temp = 0
    
    For i = 0 To UBound(Array1)
        temp = temp + Array1(i) * Array2(i)
    Next i
    
    DotProduct = temp

End Function

Function LengthVector(Array1 As Variant) As Double
    Dim i As Integer
    Dim temp As Double
    
    temp = 0
    
    For i = 0 To UBound(Array1)
        temp = temp + Array1(i) ^ 2
    Next i
    
    LengthVector = Sqr(temp)
End Function

Function GetLongestEdge(Face As Variant) As Variant

    Dim vEdges                  As Variant
    Dim swCurve                 As SldWorks.Curve
    Dim swLongestCurve          As SldWorks.Curve
    Dim swCurveParam            As SldWorks.CurveParamData
    Dim i                       As Integer
    Dim j                       As Long
    Dim vCurveParam             As Variant
    Dim vCurveParamLongest      As Variant
    Dim EdgeLength              As Double
    Dim TempEdgeLength          As Double
    Dim TempNonPerpEdgeLength   As Double
    Dim LongestEdge             As SldWorks.Edge
    Dim LongestEdge2            As SldWorks.Edge
    Dim EdgeVector(2)           As Double
    Dim Perpen1                 As Boolean 'To check if previous edge perpendicular
    Dim Perpen2                 As Boolean 'To check if next edge perpendicular
    Dim vEndPoint               As Variant
    Dim vStartPoint             As Variant
   
    vEdges = Face.GetEdges
         
    'Find longest edge
    For j = 0 To UBound(vEdges)
        
        'Debug.Print ""
        'Debug.Print "Edge Number: " & j + 1
   
        'Check if edge is a straight line
        Set swCurve = vEdges(j).GetCurve
        
        If swCurve.Identity = 3001 Then
        
            If j = 0 Then
                Perpen1 = Perpendicular(vEdges(0), vEdges(UBound(vEdges)))
                Perpen2 = Perpendicular(vEdges(0), vEdges(1))
            ElseIf j = UBound(vEdges) Then
                Perpen1 = Perpendicular(vEdges(UBound(vEdges)), vEdges(0))
                Perpen2 = Perpendicular(vEdges(UBound(vEdges)), vEdges(UBound(vEdges) - 1))
            Else
                Perpen1 = Perpendicular(vEdges(j), vEdges(j + 1))
                Perpen2 = Perpendicular(vEdges(j), vEdges(j - 1))
            End If
              
            
                Set swCurve = vEdges(j).GetCurve
                vCurveParam = vEdges(j).GetCurveParams2
                EdgeLength = swCurve.GetLength3(vCurveParam(6), vCurveParam(7))
             
             If Perpen1 Or Perpen2 Then
               
                'Debug.Print "Length = " & Round(EdgeLength * 1000, Precision)
                      
                If EdgeLength > TempEdgeLength Then
                    TempEdgeLength = EdgeLength
                    Set LongestEdge = vEdges(j)
                End If
            
            Else
                'Debug.Print "==== Edge does not have adjacent perpendicular lines ===="
            
                If EdgeLength > TempNonPerpEdgeLength Then
                    TempNonPerpEdgeLength = EdgeLength
                    Set LongestEdge2 = vEdges(j)
                End If
            
            End If
            
        End If
 
    Next j
    
    If Not LongestEdge Is Nothing Then
        'Get longest edge data
        Set swCurve = LongestEdge.GetCurve
        vCurveParam = LongestEdge.GetCurveParams2
        EdgeLength = TempEdgeLength
    
        Set GetLongestEdge = LongestEdge
        
    ElseIf Not LongestEdge2 Is Nothing Then
        'Get longest edge data
        Set swCurve = LongestEdge2.GetCurve
        vCurveParam = LongestEdge2.GetCurveParams2
        EdgeLength = TempNonPerpEdgeLength
    
        Set GetLongestEdge = LongestEdge2
        
    Else
        fileStream.WriteLine "Could not find straight edge in: " & ModelName
        'Could not find straight edge in this component
        Set GetLongestEdge = Nothing
        Exit Function
    End If
            
End Function

Function Perpendicular(Edge1 As Variant, Edge2 As Variant) As Boolean 'Checks if two edges are perpendicular
    Dim swCurve                 As SldWorks.Curve
    Dim vCurvePara1             As Variant
    Dim vCurvePara2             As Variant
    Dim i                       As Integer
    Dim Vector1(2)              As Double
    Dim Vector2(2)              As Double
    Dim LengthVector1           As Double
    Dim LengthVector2           As Double
    Dim Angle                   As Double
    Dim DotProd                 As Double
    
    Set swCurve = Edge2.GetCurve
    
    'If edge is not a line then exit function
    If Not swCurve.Identity = 3001 Then
        Perpendicular = False
        'Debug.Print "Edge not line = " & swCurve.Identity
        Exit Function
    End If
    
    vCurvePara1 = Edge1.GetCurveParams2
    vCurvePara2 = Edge2.GetCurveParams2
    
    For i = 0 To 2
        Vector1(i) = vCurvePara1(i + 3) - vCurvePara1(i)
        Vector2(i) = -vCurvePara2(i + 3) + vCurvePara2(i)
    Next i
    
    DotProd = DotProduct(Vector1, Vector2)

    LengthVector1 = LengthVector(Vector1)
    LengthVector2 = LengthVector(Vector2)

    Angle = (DotProd / (LengthVector1 * LengthVector2))
    Angle = Arccos(Angle)
    Angle = Angle * 180 / pi
    Angle = Round(Angle, Precision)
    
    If Angle = 90 Then
        Perpendicular = True
    Else
        Perpendicular = False
    End If
    
    'Debug.Print "Angle is: " & Angle & "    Perpendicular? " & Perpendicular; ""
        
End Function


Function GetParts(Assembly As SldWorks.AssemblyDoc) As Variant

    Dim swBomQuant              As Object 'Key->Path, Value->Quantity
    Dim vComps                  As Variant
    Dim swComp                  As SldWorks.Component2
    Dim i                       As Integer
    Dim Path                    As String
    Dim Desc                    As String

    Set swBomQuant = CreateObject("Scripting.Dictionary")

    vComps = Assembly.GetComponents(False)
    
    For i = 0 To UBound(vComps)
        
        Set swComp = vComps(i)
        Debug.Print swComp.Name
        'Debug.Print swComp.IsSuppressed
       
        'Check if suppresed
        If swComp.IsSuppressed = False Then
            
            Path = swComp.GetPathName
            
            If swBomQuant.exists(Path) Then
                swBomQuant.Item(Path) = swBomQuant.Item(Path) + 1
            Else
                swBomQuant.Add Path, 1
            End If
        End If
    Next
    
    Dim vItems As Variant
    vItems = swBomQuant.Keys
    
'    For i = 0 To UBound(vItems)
'        Debug.Print vItems(i) & ", " & swBomQuant.Item(vItems(i))
'    Next

    GetParts = vItems
    
End Function

Public Function UpdateBar(Maxiter As Integer, ByVal Caption As String)
    Dim CurrentProgress     As Double
    Dim BarWidth            As Double
    Dim ProgressPercentage  As Double
    Dim a                   As Integer
    
    If Caption = "Rebuilding..." Then
        Progress = Progress - 1
        ProgressBar.Bar.BackColor = &HC000&
    End If
        
    Progress = Progress + 1
    Maxiter = Maxiter + 1 'Base 0
    
    'Update Progress Bar
    CurrentProgress = Progress / Maxiter
    BarWidth = ProgressBar.Frame.Width * CurrentProgress
    ProgressPercentage = Round(CurrentProgress * 100, 0)
    ProgressBar.Bar.Width = BarWidth - 0.015 * BarWidth
    'ProgressBar.Text2.Caption = progress
    ProgressBar.Text2.Caption = ProgressPercentage & "% Complete"
    ProgressBar.Text.Caption = Progress & " of " & Maxiter
    ProgressBar.Text3.Caption = Caption
    
    'If user uses stop button
    DoEvents
    
    If (ProgressPercentage / 10) Mod 2 = 0 Then
        ProgressBar.Image1.Visible = False
        ProgressBar.Image2.Visible = True
    Else
        ProgressBar.Image1.Visible = True
        ProgressBar.Image2.Visible = False
    End If

    
'    If CancelButton = True Then
'        Unload ProgressBar
'        Call Finalize
'        End
'    Else
'        DoEvents
'    End If

End Function

Function GetClosestTransform(swThisBody As SldWorks.Body2, swOtherBody As SldWorks.Body2) As SldWorks.MathTransform

    
    Dim transformsHits      As Object
    Dim swTransform         As SldWorks.MathTransform
    Dim key                 As Variant '???
    Set transformsHits = CreateObject("Scripting.Dictionary") 'for some reasons sometimes the first null element is added on creation
        
         
    If swThisBody.GetCoincidenceTransform2(swOtherBody, swTransform) Then
        
        If Not swTransform Is Nothing Then
            'Debug.Print "Could Coincide " & ModelName
       
        
            'Create vXfm only to print the transformation
            Dim vXfm As Variant
            
            vXfm = swTransform.ArrayData
            
'            Debug.Print "Rotation:"
'            Debug.Print vbTab & Round(vXfm(0) * 1000, 4), Round(vXfm(1) * 1000, 4), Round(vXfm(2) * 1000, 4)
'            Debug.Print vbTab & Round(vXfm(3) * 1000, 4), Round(vXfm(4) * 1000, 4), Round(vXfm(5) * 1000, 4)
'            Debug.Print vbTab & Round(vXfm(6) * 1000, 4), Round(vXfm(7) * 1000, 4), Round(vXfm(8) * 1000, 4)
'            Debug.Print "Translation:"
'            Debug.Print vbTab & Round(vXfm(9) * 1000, 4), Round(vXfm(10) * 1000, 4), Round(vXfm(11) * 1000, 4)
'            Debug.Print "Scaling: " & vXfm(12)
            
            Dim contains As Boolean
            contains = False
            For Each key In transformsHits.Keys
                If Not key Is Nothing Then
                    Dim tx As SldWorks.MathTransform
                    Set tx = key
                    If CompareTransforms(swTransform, tx) Then
                        transformsHits(tx) = transformsHits(tx) + 1
                        contains = True
                        Exit For
                    End If
                End If
            Next
            
            If Not contains Then
                transformsHits.Add swTransform, 1
            End If
            
        End If
         
    Else
        
        Debug.Print "CANNOT COINCIDE " & ModelName
        fileStream.WriteLine "Could not reset location of: " & ModelName
    End If
                
 
    
    Dim curMaxHit As Integer
    curMaxHit = 0
    
    For Each key In transformsHits.Keys
        If Not key Is Nothing Then
            Dim curTx As SldWorks.MathTransform
            Set curTx = key
            If transformsHits(curTx) > curMaxHit Then
                curMaxHit = transformsHits(curTx)
                Set GetClosestTransform = curTx
            End If
        End If
    Next

End Function

Function CompareTransforms(firstTransform As SldWorks.MathTransform, secondTransform As SldWorks.MathTransform) As Boolean
    
    Dim vFirstArrayData As Variant
    vFirstArrayData = firstTransform.ArrayData
    
    Dim vSecondArrayData As Variant
    vSecondArrayData = secondTransform.ArrayData
    
    Dim i As Integer
    
    For i = 0 To UBound(vFirstArrayData)
        If Not CompareValues(CDbl(vFirstArrayData(i)), CDbl(vSecondArrayData(i))) Then
            CompareTransforms = False
            Exit Function
        End If
    Next
    
    CompareTransforms = True
    
End Function

Function CompareValues(firstValue As Double, secondValue As Double, Optional tol As Double = 0.00000001) As Boolean
        
    CompareValues = Abs(secondValue - firstValue) <= tol
    
End Function

Function FileName(File As String) As String

    Dim FullFileName As String
    
    FullFileName = Right(File, Len(File) - InStrRev(File, "\"))

    If Right(FullFileName, 7) = ".sldasm" Or Right(FullFileName, 7) = ".sldprt" Then
        FileName = Left(FullFileName, Len(FullFileName) - 7)
    Else
        FileName = FullFileName
    End If

End Function

Function NamePath(Name As String) As String

    NamePath = Replace(Name, "/", "@")

End Function

Function SelectPlane(model As SldWorks.ModelDoc2, planeType As Integer) As String

    Dim planeIndex As Integer
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature

    Do While Not swFeat Is Nothing

        'Debug.Print swFeat.Description
        
        If swFeat.GetTypeName = "RefPlane" Then
        
            Debug.Print swFeat.Description
            
            planeIndex = planeIndex + 1
            
            'Debug.Print planeIndex
            
            If CInt(planeType) = planeIndex Then

                'swFeat.Select2 True, 0

                SelectPlane = swFeat.Description
                
                Exit Function

            End If

        End If
    
        Set swFeat = swFeat.GetNextFeature

    Loop
    
End Function

