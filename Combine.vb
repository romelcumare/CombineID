' V0.6 takes into account grain direction

'####Logic

' #Get all parts
' Delete CombineID property
' Explude parts without depth, width and thickness property
' exclude parts that have IS_HARDWARE and Combine property
' # Compare parts avoiding duplicate compares
' check if parts have the same material
' Get biggest bodies for both parts
' get transformation matrix between bodies
' if they can coincide then run CheckProperties

' #CheckProperties
' Get number of Edgebands
' if number of edgebands and laminates is different then exit

' Get symmetry type (Unique, Rotatable, Rotatable and Flippable, Fully Symmetric)
' Depending on the symmetry type, check possible rotation matrixes and run CheckRotation


' #CheckRotation
' Checks if the panel if flipped or rotated
' if flipped
'     Checks if laminates are the same, exit if not
'         check no edgebands are preset, exit

'     Check if edgebands are the same



Option Explicit
Option Compare Text

Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim swPart1                 As SldWorks.PartDoc
Dim swPart2                 As SldWorks.PartDoc
Dim swAssy                  As SldWorks.AssemblyDoc
Dim swComp                  As SldWorks.Component2
Dim longstatus              As Long
Dim boolstatus              As Boolean
Dim AssemblyPath            As String
Dim AssemblyName            As String
Dim DocName                 As String
Dim DocPath                 As String
Dim fso                     As FileSystemObject
Dim fileStream              As TextStream
Dim ModelName1              As String 'Model being processed
Dim ModelName2              As String 'Model to compair
Dim CombineID               As Integer
Dim GlobalCombineID         As Integer
Dim PartList                As Variant
Dim CanCoincide             As Boolean
Dim CanCoincideGlobal       As Boolean
Dim UserName                As String
Dim ErrorMessages           As Integer

Dim CurrentProgress         As Double
Dim BarWidth                As Long
Dim ProgressPercentage      As Double
Dim Progress                As Integer
Dim CancelButton            As Boolean



Sub main()

    ErrorMessages = 0

    Debug.Print ""
    Debug.Print "----------- Macro Started -----------"

    Dim i                       As Integer
    Dim j                       As Integer
    Dim NumofParts              As Integer
    Dim NumofChecks             As Integer

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    AssemblyName = FileName(swModel.GetTitle())
         
    If Not swModel.GetType = SwConst.swDocASSEMBLY Then
        MsgBox "The active document is not an assembly model.", vbOKOnly
        End
    End If
    
    On Error GoTo ErrorHandler
    
    ' Create Log File
    DocName = swModel.GetTitle()
    DocPath = swModel.GetPathName

    If Right(DocName, 7) = ".sldasm" Then
        DocName = Left(DocName, InStrRev(DocName, ".") - 1)
    End If

    UserName = Environ("USERNAME")
    DocPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") - 1)
    DocPath = DocPath + "\" + DocName + " - Combine Log.txt"
    Set fso = New FileSystemObject
    Set fileStream = fso.OpenTextFile(DocPath, ForAppending, True)
    fileStream.WriteLine ""
    fileStream.WriteLine "----------- " & UserName & " - " & Now & " -----------"
  
    Set swAssy = swModel
    AssemblyPath = swModel.GetPathName
    
    ' Deactivate view update
    Dim modView As ModelView
    Set modView = swModel.ActiveView
    modView.EnableGraphicsUpdate = False
    swAssy.FeatureManager.EnableFeatureTree = False
    
    ' Show Progress Bar
    Progress = 0
    
    With ProgressBar
        .Bar.Width = 0
        .Text.caption = "Getting Parts..."
        .Text2.caption = "0% Complete"
        .Text3.caption = "Processing..."
        .Show vbModeless
    End With
    
    CancelButton = False
                
    PartList = GetParts(swAssy)
    
'    GoTo ExitCode
'    GroupPanels (PartList)
    
    NumofParts = UBound(PartList) + 1
    'Debug.Print "Number of parts: " & NumofParts
    
    NumofChecks = (NumofParts * (NumofParts - 1)) / 2
    
    ' Set Combine IDs
    GlobalCombineID = 1
                   
    For i = 0 To UBound(PartList) - 1
    
        Call UpdateBar(UBound(PartList), "Processing...")
        
        For j = i To UBound(PartList)
    
            If Not i = j Then
                CombineParts (PartList(i)), (PartList(j))
            End If
            
            If CancelButton = True Then
                Debug.Print "CANCEL IS TRUE"
                fileStream.WriteLine "*----------- User Cancelled Macro -----------*"
                GoTo ExitCode
            End If
        
        Next j

        ' Increse GlobalCombineID is coincide was found
        If CanCoincideGlobal Then
            GlobalCombineID = GlobalCombineID + 1
            CanCoincideGlobal = False
        End If

    Next i
    
    Call UpdateBar(UBound(PartList), "Processing...")
    
ExitCode:

    Call UpdateBar(UBound(PartList), "Rebuilding...")

    Set swModel = swApp.ActivateDoc3(AssemblyPath, False, 1, 0)

    ' Rebuild Assembly
    boolstatus = swModel.EditRebuild3()

    swModel.ClearSelection2 (True)

    ' Activate view update
    modView.EnableGraphicsUpdate = True
    swAssy.FeatureManager.EnableFeatureTree = True

    Unload ProgressBar

    Debug.Print ""
    Debug.Print "----------- Macro Finished -----------"
    fileStream.WriteLine "-----------       Macro Finished       -----------"

    ' Close log file
    fileStream.Close
    
    If ErrorMessages = 1 Then
        MsgBox ErrorMessages & " error occured. Review log file", vbExclamation, "Error"
    ElseIf ErrorMessages > 1 Then
        MsgBox ErrorMessages & " errors occured. Review log file", vbExclamation, "Error"
    End If
    ' Exit Sub and don't proceed to ErrorHandler
    Exit Sub

ErrorHandler:

    Call UpdateBar(UBound(PartList), "Rebuilding...")

    Set swModel = swApp.ActivateDoc3(AssemblyPath, False, 1, 0)

    ' Rebuild Assembly
    boolstatus = swModel.EditRebuild3()

    swModel.ClearSelection2 (True)

    ' Activate view update
    modView.EnableGraphicsUpdate = True
    swAssy.FeatureManager.EnableFeatureTree = True

    Unload ProgressBar

    Debug.Print ""
    Debug.Print "----------- ERROR while processing " & ModelName1 & " and " & ModelName2 & "-----------"
    fileStream.WriteLine "----------- ERROR while processing " & ModelName1 & " and " & ModelName2 & "-----------"

    ' Close log file
    fileStream.Close
    
    MsgBox "Error Occured while processing: " & ModelName1 & " and " & ModelName2, vbExclamation, "Error"

End Sub


Sub CombineParts(Model1 As String, Model2 As String)

    Dim swModel1                        As SldWorks.ModelDoc2
    Dim swModel2                        As SldWorks.ModelDoc2
    Dim swModelDocExt1                  As ModelDocExtension
    Dim swModelDocExt2                  As ModelDocExtension
    Dim swModelDoc1                     As SldWorks.ModelDoc2
    Dim swModelDoc2                     As SldWorks.ModelDoc2
    Dim swCustProp1                     As CustomPropertyManager
    Dim swCustProp2                     As CustomPropertyManager
    Dim Material1                       As String
    Dim Material2                       As String
    Dim vBodies1                        As Variant
    Dim vBodies2                        As Variant
    Dim swBody1                         As SldWorks.Body2
    Dim swBody2                         As SldWorks.Body2
    Dim swFace                          As SldWorks.Entity
    Dim FaceNormal(15)                  As Variant
    Dim longstatus                      As Long
    Dim longwarnings                    As Long
    Dim boolstatus                      As Boolean
    Dim RotMatrix                       As SldWorks.MathTransform
    Dim i                               As Integer
    Dim j                               As Integer
    Dim n                               As Integer
    Dim SameLamEdge                     As Boolean
    Dim retval                          As Long
    
    ' Get Part1 Data
    'Debug.Print ""
    'Debug.Print "Processing:  " & FileName(Model1)
    
    ' Reset CanCoincide
    CanCoincide = False
    
    ModelName1 = FileName(Model1)
    Set swModel1 = swApp.OpenDoc6(Model1, 1, 0, "", longstatus, longwarnings)
    Set swPart1 = swModel1
    Set swModelDoc1 = swModel1
    Set swModelDocExt1 = swModelDoc1.Extension
    Material1 = swModelDoc1.MaterialIdName
    'Debug.Print "Material:  " & swModelDoc1.MaterialIdName
    Set swCustProp1 = swModelDocExt1.CustomPropertyManager("")

    ' Check if already has CombineID and use it
    If swCustProp1.Get("CombineID") = "" Then
        ' Potentially use GlobalCombineID if coincidense is found
        CombineID = GlobalCombineID
    Else
        ' Grab current CombineID
        'Debug.Print ModelName1 & ": Already has CombineID"
        CombineID = swCustProp1.Get("CombineID")
    End If

    ' Get Part2 Data
    'Debug.Print " Comparing:  " & FileName(Model2)

    ModelName2 = FileName(Model2)
    Set swModel2 = swApp.OpenDoc6(Model2, 1, 0, "", longstatus, longwarnings)
    Set swPart2 = swModel2
    Set swModelDoc2 = swModel2
    Set swModelDocExt2 = swModelDoc2.Extension
    Material2 = swModelDoc2.MaterialIdName
    'Debug.Print "Material:  " & swModelDoc2.MaterialIdName
    Set swCustProp2 = swModelDocExt2.CustomPropertyManager("")

    ' Check if Part2 already has CombineID and skip
    If swCustProp2.Get("CombineID") <> "" Then
        'Debug.Print ModelName2 & ": Already has CombineID and will not be checked"
        Exit Sub
    End If

    ' Check if both Parts have the same material. Skip coincidence check if not
    If Not Material1 = Material2 Then
        Exit Sub
    End If

    ' Get Biggest Body1
    Set swBody1 = GetBiggestBody(swPart1)
    'Debug.Print "Biggest Body1 is: " & swBody1.Name

    ' Get Biggest Body2
    Set swBody2 = GetBiggestBody(swPart2)
    'Debug.Print "Biggest Body2 is: " & swBody2.Name
        
    ' Get rotation matrix if bodies can coincide
    Set RotMatrix = GetTransformMatrix(swBody2, swBody1)
    
    If CanCoincide Then
'        ' Check Symmetry type
'        SymmetryType = GetSymmetryType(swBody1)
'        Debug.Print " Panel Type : " & SymmetryType
        
        ' Check if panels have the same number of laminates and edge
        SameLamEdge = CheckProperties(swModel1, swModel2, RotMatrix, swBody1)
    End If
    
    If (CanCoincide And SameLamEdge) Then
    
        CanCoincideGlobal = True
       
        ' Add CombineID custom property to both parts
        Set swCustProp2 = swModelDocExt2.CustomPropertyManager("")
        
        retval = swCustProp1.Add3("CombineID", swCustomInfoText, CombineID, swCustomPropertyReplaceValue)
        retval = swCustProp2.Add3("CombineID", swCustomInfoText, CombineID, swCustomPropertyReplaceValue)
                 
        ' Flag models as dirty
        swPart1.SetSaveFlag
        swPart2.SetSaveFlag
        'Debug.Print " is it Dirty? " & swPart1.GetSaveFlag
        'Debug.Print " is it Dirty? " & swPart2.GetSaveFlag
        
        'Debug.Print ""
        Debug.Print " CombineID " & CombineID & ": TRUE"
        Debug.Print "------------------------------------"
        
    ElseIf CanCoincide And SameLamEdge = False Then
        Debug.Print " Combine ID : False - Different Edgebands and Laminates"
        Debug.Print "------------------------------------"
        
    End If
            
End Sub

Function CheckProperties(swModel1 As SldWorks.ModelDoc2, swModel2 As SldWorks.ModelDoc2, RotMatrix As SldWorks.MathTransform, swBody1 As SldWorks.Body2) As Boolean

    Dim SymmetryType        As String
    Dim config1             As SldWorks.Configuration
    Dim config2             As SldWorks.Configuration
    Dim cusPropMgr1         As SldWorks.CustomPropertyManager
    Dim cusPropMgr2         As SldWorks.CustomPropertyManager
    Dim i                   As Integer
               
    Dim NumofEdgebands1     As Integer
    Dim NumofEdgebands2     As Integer
    Dim NumofLaminates1     As Integer
    Dim NumofLaminates2     As Integer
        
    Dim GrainAngle1         As Double
    Dim Top1                As Variant
    Dim Bottom1             As Variant
    Dim Front1              As Variant
    Dim Back1               As Variant
    Dim Left1               As Variant
    Dim Right1              As Variant
    
    Dim GrainAngle2         As Double
    Dim Top2                As Variant
    Dim Bottom2             As Variant
    Dim Front2              As Variant
    Dim Back2               As Variant
    Dim Left2               As Variant
    Dim Right2              As Variant
    
    Set config1 = swModel1.GetActiveConfiguration
    Set config2 = swModel2.GetActiveConfiguration
    
    Set cusPropMgr1 = config1.CustomPropertyManager
    Set cusPropMgr2 = config2.CustomPropertyManager
    
'    ' Custom SWOODDesign.cfg
' Get Panel1 Data
    '    ' Custom SWOODDesign.cfg
'    ' Get Panel1 Data
'    Top1 = Array(cusPropMgr1.Get("Swood_LamTop_Material"), cusPropMgr1.Get("Swood_LamTop_Thickness"))
'    Bottom1 = Array(cusPropMgr1.Get("Swood_LamBottom_Material"), cusPropMgr1.Get("Swood_LamBottom_Thickness"))
'    Front1 = Array(cusPropMgr1.Get("Swood_EBFront_Material"), cusPropMgr1.Get("Swood_EBFront_Thickness"))
'    Back1 = Array(cusPropMgr1.Get("Swood_EBBack_Material"), cusPropMgr1.Get("Swood_EBBack_Thickness"))
'    Left1 = Array(cusPropMgr1.Get("Swood_EBLeft_Material"), cusPropMgr1.Get("Swood_EBLeft_Thickness"))
'    Right1 = Array(cusPropMgr1.Get("Swood_EBRight_Material"), cusPropMgr1.Get("Swood_EBRight_Thickness"))
'
'    ' Get Panel2 Data
'    Top2 = Array(cusPropMgr2.Get("Swood_LamTop_Material"), cusPropMgr2.Get("Swood_LamTop_Thickness"))
'    Bottom2 = Array(cusPropMgr2.Get("Swood_LamBottom_Material"), cusPropMgr2.Get("Swood_LamBottom_Thickness"))
'    Front2 = Array(cusPropMgr2.Get("Swood_EBFront_Material"), cusPropMgr2.Get("Swood_EBFront_Thickness"))
'    Back2 = Array(cusPropMgr2.Get("Swood_EBBack_Material"), cusPropMgr2.Get("Swood_EBBack_Thickness"))
'    Left2 = Array(cusPropMgr2.Get("Swood_EBLeft_Material"), cusPropMgr2.Get("Swood_EBLeft_Thickness"))
'    Right2 = Array(cusPropMgr2.Get("Swood_EBRight_Material"), cusPropMgr2.Get("Swood_EBRight_Thickness"))

    ' Original SWOODDesign.cfg
    ' Get Panel1 Data
    Debug.Print " Grain angle1: " & cusPropMgr1.Get("SWOODCP_PanelGrainAngleInFrontView")
    If Not IsNumeric(cusPropMgr1.Get("SWOODCP_PanelGrainAngleInFrontView")) Then
        Debug.Print "Unable to get grain angle for " & ModelName1
        fileStream.WriteLine "  Unable to get grain angle for " & ModelName1 & ". Check if Front View if normal to top face"
        ErrorMessages = ErrorMessages + 1
    Else
        GrainAngle1 = cusPropMgr1.Get("SWOODCP_PanelGrainAngleInFrontView")
    End If
        
    Top1 = Array(cusPropMgr1.Get("SWOODCP_TopStockMaterial"), cusPropMgr1.Get("SWOODCP_TopStockThickness"), CDbl(cusPropMgr1.Get("SWOODCP_BottomStockGrainAngleInFrontView")))
    Bottom1 = Array(cusPropMgr1.Get("SWOODCP_BottomStockMaterial"), cusPropMgr1.Get("SWOODCP_BottomStockThickness"), CDbl(cusPropMgr1.Get("SWOODCP_BottomStockGrainAngleInFrontView")))
    Front1 = Array(cusPropMgr1.Get("SWOODCP_EdgeFrontMaterial"), cusPropMgr1.Get("SWOODCP_EdgeFrontThickness"))
    Back1 = Array(cusPropMgr1.Get("SWOODCP_EdgeBackMaterial"), cusPropMgr1.Get("SWOODCP_EdgeBackThickness"))
    Left1 = Array(cusPropMgr1.Get("SWOODCP_EdgeLeftMaterial"), cusPropMgr1.Get("SWOODCP_EdgeLeftThickness"))
    Right1 = Array(cusPropMgr1.Get("SWOODCP_EdgeRightMaterial"), cusPropMgr1.Get("SWOODCP_EdgeRightThickness"))

    ' Get Panel2 Data
    Debug.Print " Grain angle2: " & cusPropMgr2.Get("SWOODCP_PanelGrainAngleInFrontView")
    
    If Not IsNumeric(cusPropMgr2.Get("SWOODCP_PanelGrainAngleInFrontView")) Then
        Debug.Print "Unable to get grain angle for " & ModelName2
        fileStream.WriteLine "  Unable to get grain angle for " & ModelName2 & ". Check if Front View if normal to top face"
        ErrorMessages = ErrorMessages + 1
    Else
        GrainAngle2 = cusPropMgr2.Get("SWOODCP_PanelGrainAngleInFrontView")
    End If
    
    Top2 = Array(cusPropMgr2.Get("SWOODCP_TopStockMaterial"), cusPropMgr2.Get("SWOODCP_TopStockThickness"), CDbl(cusPropMgr2.Get("SWOODCP_BottomStockGrainAngleInFrontView")))
    Bottom2 = Array(cusPropMgr2.Get("SWOODCP_BottomStockMaterial"), cusPropMgr2.Get("SWOODCP_BottomStockThickness"), CDbl(cusPropMgr2.Get("SWOODCP_BottomStockGrainAngleInFrontView")))
    Front2 = Array(cusPropMgr2.Get("SWOODCP_EdgeFrontMaterial"), cusPropMgr2.Get("SWOODCP_EdgeFrontThickness"))
    Back2 = Array(cusPropMgr2.Get("SWOODCP_EdgeBackMaterial"), cusPropMgr2.Get("SWOODCP_EdgeBackThickness"))
    Left2 = Array(cusPropMgr2.Get("SWOODCP_EdgeLeftMaterial"), cusPropMgr2.Get("SWOODCP_EdgeLeftThickness"))
    Right2 = Array(cusPropMgr2.Get("SWOODCP_EdgeRightMaterial"), cusPropMgr2.Get("SWOODCP_EdgeRightThickness"))
    
    
    ' Get number of Edgebands
    If Front1(0) <> "" Then NumofEdgebands1 = NumofEdgebands1 + 1
    If Back1(0) <> "" Then NumofEdgebands1 = NumofEdgebands1 + 1
    If Left1(0) <> "" Then NumofEdgebands1 = NumofEdgebands1 + 1
    If Right1(0) <> "" Then NumofEdgebands1 = NumofEdgebands1 + 1
    
    If Front2(0) <> "" Then NumofEdgebands2 = NumofEdgebands2 + 1
    If Back2(0) <> "" Then NumofEdgebands2 = NumofEdgebands2 + 1
    If Left2(0) <> "" Then NumofEdgebands2 = NumofEdgebands2 + 1
    If Right2(0) <> "" Then NumofEdgebands2 = NumofEdgebands2 + 1

    
    ' Get number of Laminates
    If Top1(0) <> "" Then NumofLaminates1 = NumofLaminates1 + 1
    If Bottom1(0) <> "" Then NumofLaminates1 = NumofLaminates1 + 1
    
    If Top2(0) <> "" Then NumofLaminates2 = NumofLaminates2 + 1
    If Bottom2(0) <> "" Then NumofLaminates2 = NumofLaminates2 + 1
        
    'Debug.Print "Number of Edgebands 1: " & NumofEdgebands1
    'Debug.Print "Number of Edgebands 2: " & NumofEdgebands2
    'Debug.Print "Number of Laminates 1: " & NumofLaminates1
    'Debug.Print "Number of Laminates 2: " & NumofLaminates2
    
    If Not (NumofEdgebands1 = NumofEdgebands2 And NumofLaminates1 = NumofLaminates2) Then
        CheckProperties = False
        Debug.Print " Different number of Edgebands and Laminates"
        Exit Function
    End If
    
    
    '####################### move this check after grain check!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'    If NumofEdgebands1 + NumofEdgebands2 + NumofLaminates1 + NumofLaminates2 = 0 Then
'        Debug.Print " No Edgebands or Laminates present"
'        CheckProperties = True
'        Exit Function
'    ElseIf NumofEdgebands1 + NumofEdgebands2 = 0 Then Debug.Print " No Edgebands present"
'
'    ElseIf NumofLaminates1 + NumofLaminates2 = 0 Then Debug.Print " No Laminates present"
'
'    End If
    
    'Debug.Print "CheckProperties result: " & CheckProperties
        
    ' Check Symmetry type
    SymmetryType = GetSymmetryType(swBody1)
    Debug.Print " Panel Type : " & SymmetryType
        
    Dim swMathUtil            As SldWorks.mathUtility
    Dim Mdat(15)              As Double
    Dim SymmetricMatrix       As SldWorks.MathTransform
    
    Set swMathUtil = swApp.GetMathUtility
    
    If SymmetryType = "Unique" Then
                
        CheckProperties = CheckRotation(RotMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
   
        If CheckProperties Then Exit Function
        
    End If
    
    If SymmetryType = "Rotatable" Then
    
        CheckProperties = CheckRotation(RotMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 1"
        If CheckProperties Then Exit Function
        
        ' Matrix for Rotatable Panels
        Mdat(0) = -1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = -1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = 1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 2"
        If CheckProperties Then Exit Function
        
    End If
    
    If SymmetryType = "Rotatable and Flippable" Then
    
        CheckProperties = CheckRotation(RotMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 1"
        If CheckProperties Then Exit Function
        
        ' Matrixes for Rotatable and Flippable Panels
        Mdat(0) = -1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = -1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = 1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 2"
        If CheckProperties Then Exit Function
        
        Mdat(0) = -1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = 1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = -1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 3"
        If CheckProperties Then Exit Function
        
        Mdat(0) = 1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = -1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = -1:
    
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 4"
        If CheckProperties Then Exit Function
        
    End If
    
    If SymmetryType = "Fully Symmetric" Then
    
        CheckProperties = CheckRotation(RotMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 1"
        If CheckProperties Then Exit Function
        
        ' Matrixes for Fully Symmetrical Panels
        
        Mdat(0) = 0: Mdat(1) = 1: Mdat(2) = 0:
        Mdat(3) = -1: Mdat(4) = 0: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = 1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 2"
        If CheckProperties Then Exit Function
        
        Mdat(0) = -1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = -1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = 1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 3"
        If CheckProperties Then Exit Function
        
        Mdat(0) = 0: Mdat(1) = -1: Mdat(2) = 0:
        Mdat(3) = 1: Mdat(4) = 0: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = 1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 4"
        If CheckProperties Then Exit Function
        
        Mdat(0) = -1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = 1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = -1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 5"
        If CheckProperties Then Exit Function
        
        Mdat(0) = 0: Mdat(1) = -1: Mdat(2) = 0:
        Mdat(3) = -1: Mdat(4) = 0: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = -1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 6"
        If CheckProperties Then Exit Function
        
        Mdat(0) = 1: Mdat(1) = 0: Mdat(2) = 0:
        Mdat(3) = 0: Mdat(4) = -1: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = -1:
        
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 7"
        If CheckProperties Then Exit Function
        
        Mdat(0) = 0: Mdat(1) = 1: Mdat(2) = 0:
        Mdat(3) = 1: Mdat(4) = 0: Mdat(5) = 0:
        Mdat(6) = 0: Mdat(7) = 0: Mdat(8) = -1:
    
        Set SymmetricMatrix = swMathUtil.CreateTransform(Mdat)
        Set SymmetricMatrix = RotMatrix.Multiply(SymmetricMatrix)
        CheckProperties = CheckRotation(SymmetricMatrix, GrainAngle1, GrainAngle2, NumofEdgebands1, NumofEdgebands2, Top1, Top2, Bottom1, Bottom2, Front1, Front2, Back1, Back2, Left1, Left2, Right1, Right2)
        'Debug.Print " Check 8"
        If CheckProperties Then Exit Function
        
    End If
 End Function
 
 
 Function CheckRotation(RotMatrix As SldWorks.MathTransform, GrainAngle1 As Double, GrainAngle2 As Double, NumofEdgebands1 As Integer, NumofEdgebands2 As Integer, Top1 As Variant, Top2 As Variant, Bottom1 As Variant, Bottom2 As Variant, Front1 As Variant, Front2 As Variant, Back1 As Variant, Back2 As Variant, Left1 As Variant, Left2 As Variant, Right1 As Variant, Right2 As Variant) As Boolean
    
    ' Check if panel was flipped or rotated
    If Round(RotMatrix.ArrayData(8), 5) = 1 Then
    
        'Debug.Print " Panel not Flipped"
        
        ' Check Edgebands
        If RotMatrix.ArrayData(0) = 1 And RotMatrix.ArrayData(4) = 1 Then
            'Debug.Print " Panel not Rotated"
            'Debug.Print " Front => Front, Back => Back, Right => Right, Left => Left"

            'Check Panel grain angle
            If Not (GrainAngle1 = GrainAngle2) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If

            ' Check Laminates
            If Not (Compare(Top1, Top2) And Compare(Bottom1, Bottom2)) Then
                Debug.Print " Different Laminates"
                CheckRotation = False
                Exit Function
            End If
            
            If Compare(Front1, Front2) And Compare(Back1, Back2) And Compare(Right1, Right2) And Compare(Left1, Left2) Then
                CheckRotation = True
                Exit Function
            Else
                Debug.Print " Different Edgebands"
            End If
            
        End If
        
        If Round(RotMatrix.ArrayData(0), 5) = -1 And Round(RotMatrix.ArrayData(4), 5) = -1 Then
            ' Rotated 180 deg
            'Debug.Print " Front => Back, Back => Front, Right => Left, Left => Right"

            'Check Panel grain angle
            If Not (GrainAngle1 = GrainAngle2) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If

            ' Check Laminates
            If Not (Compare(Top1, Top2) And Compare(Bottom1, Bottom2)) Then
                Debug.Print " Different Laminates"
                CheckRotation = False
                Exit Function
            End If
            
            If Compare(Front1, Back2) And Compare(Back1, Front2) And Compare(Right1, Left2) And Compare(Left1, Right2) Then
                CheckRotation = True
                Exit Function
            Else
                Debug.Print " Different Edgebands"
            End If
            
        End If
        
        If Round(RotMatrix.ArrayData(1), 5) = 1 And Round(RotMatrix.ArrayData(3), 5) = -1 Then
            ' Rotated +90 deg
            'Debug.Print " Front => Right, Back => Left, Right => Back, Left => Front"

            If Not ((GrainAngle1 = GrainAngle2 + 90) Or (GrainAngle1 = GrainAngle2 - 90)) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If

            ' Check Laminates
            Dim Top2Rotated                         As Variant
            Dim Bottom2Rotated                      As Variant

            Top2Rotated = Top2
            Bottom2Rotated = Bottom2
    
            Top2Rotated(2) = Top2(2) + 90
            Bottom2Rotated(2) = Bottom2(2) + 90
            
            If Not (Compare(Top1, Top2Rotated) And Compare(Bottom1, Bottom2Rotated)) Then
                Debug.Print " Different Laminates"
                CheckRotation = False
                Exit Function
            End If
            
            If Compare(Front1, Right2) And Compare(Back1, Left2) And Compare(Right1, Back2) And Compare(Left1, Front2) Then
                CheckRotation = True
                Exit Function
            End If
            
        End If
        
        If Round(RotMatrix.ArrayData(1), 5) = -1 And Round(RotMatrix.ArrayData(3), 5) = 1 Then
            ' Rotated -90 deg
            'Debug.Print " Front => Left, Back => Right, Right => Front, Left => Back"

            If Not ((GrainAngle1 = GrainAngle2 + 90) Or (GrainAngle1 = GrainAngle2 - 90)) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If

            ' Check Laminates
            Dim Top2Rotated                         As Variant
            Dim Bottom2Rotated                      As Variant

            Top2Rotated = Top2
            Bottom2Rotated = Bottom2
    
            Top2Rotated(2) = Top2(2) - 90
            Bottom2Rotated(2) = Bottom2(2) - 90
            
            If Not (Compare(Top1, Top2Rotated) And Compare(Bottom1, Bottom2Rotated)) Then
                Debug.Print " Different Laminates"
                CheckRotation = False
                Exit Function
            End If
                
            ' Check Edgebands
            If Compare(Front1, Left2) And Compare(Back1, Right2) And Compare(Right1, Front2) And Compare(Left1, Back2) Then
                CheckRotation = True
                Exit Function
            End If
            
        End If
    
    ElseIf Round(RotMatrix.ArrayData(8), 5) = -1 Then
        'Debug.Print " Top => Bottom"
        
        ' Check Laminates
        If Not (Compare(Top1, Bottom2) And Compare(Bottom1, Top2)) Then
            Debug.Print " Different Laminates"
            CheckRotation = False
            Exit Function
        ' ElseIf NumofEdgebands1 + NumofEdgebands2 = 0 Then
        '     CheckRotation = True
        '     Exit Function
        End If
        
        ' Check Edgebands
        If Round(RotMatrix.ArrayData(1), 5) = -1 And Round(RotMatrix.ArrayData(3), 5) = -1 Then
            'Debug.Print " Front => Right, Back => Left, Right => Front, Left => Back "

            If Not ((GrainAngle1 = GrainAngle2 + 90) Or (GrainAngle1 = GrainAngle2 - 90)) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If
            
            If Compare(Front1, Right2) And Compare(Back1, Left2) And Compare(Right1, Front2) And Compare(Left1, Back2) Then
                CheckRotation = True
                Exit Function
            Else
                Debug.Print " Different Edgebands"
            End If
            
        End If
        
        If Round(RotMatrix.ArrayData(1), 5) = 1 And Round(RotMatrix.ArrayData(3), 5) = 1 Then
            'Debug.Print " Front => Left, Back => Right, Right => Back, Left => Front"

            If Not ((GrainAngle1 = GrainAngle2 + 90) Or (GrainAngle1 = GrainAngle2 - 90)) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If
            
            If Compare(Front1, Left2) And Compare(Back1, Right2) And Compare(Right1, Back2) And Compare(Left1, Front2) Then
                 CheckRotation = True
                Exit Function
            Else
                Debug.Print " Different Edgebands"
            End If
        End If
        
        If Round(RotMatrix.ArrayData(0), 5) = 1 And Round(RotMatrix.ArrayData(4), 5) = -1 Then
            'Debug.Print " Front => Back, Back => Front, Right => Right, Left => Left"

            If Not (GrainAngle1 = GrainAngle2) Then
                Debug.Print "Different panel grain angle"
                CheckRotation = False
                Exit Function
            End If
            
            If Compare(Front1, Back2) And Compare(Back1, Front2) And Compare(Right1, Right2) And Compare(Left1, Left2) Then
                CheckRotation = True
                Exit Function
            Else
                Debug.Print " Different Edgebands"
            End If
        End If
        
        If Round(RotMatrix.ArrayData(0), 5) = -1 And Round(RotMatrix.ArrayData(4), 5) = 1 Then
            'Debug.Print " Front => Front, Back => Back, Right => Left, Left => Right"

            If Not (GrainAngle1 = GrainAngle2) Then
                CheckRotation = False
                Debug.Print "Different panel grain angle"
                Exit Function
            End If
            
            If Compare(Front1, Front2) And Compare(Back1, Back2) And Compare(Right1, Left2) And Compare(Left1, Right2) Then
                CheckRotation = True
                Exit Function
            Else
                Debug.Print " Different Edgebands"
            End If
        End If
    End If
           
End Function


Function GetParts(Assembly As SldWorks.AssemblyDoc) As Variant

    Dim swBomQuant As Object 'Key->Path, Value->Quantity
    Dim vComps As Variant
    Dim swComp As SldWorks.Component2
    Dim i As Integer
    Dim Path As String
    Dim swSelModel                      As SldWorks.ModelDoc2
    Dim swModelDocExt                   As ModelDocExtension
    Dim swCustProp                      As CustomPropertyManager


    Set swBomQuant = CreateObject("Scripting.Dictionary")

    vComps = Assembly.GetComponents(False)
    
    For i = 0 To UBound(vComps)
        
        Set swComp = vComps(i)
        'Debug.Print vComps(i).Name
        
        ' Check if suppresed and if it's part model
        If swComp.IsSuppressed = False Then
        
            Set swSelModel = swComp.GetModelDoc2
            
            If swSelModel.GetType = 1 Then
          
                Set swModelDocExt = swSelModel.Extension
                Set swCustProp = swModelDocExt.CustomPropertyManager("")
               
                'Debug.Print swSelModel.GetType
                'Debug.Print swComp.Name
                'Debug.Print swComp.IsSuppressed
                
                Path = swComp.GetPathName
                'Debug.Print Path
                
                    ' Delete CombineID if it exists
                    boolstatus = swCustProp.Delete2("CombineID")
                
                    ' Check if it has LxWxT properties
                    If Not (swCustProp.Get("Length") = "" Or swCustProp.Get("Width") = "" Or swCustProp.Get("Thickness") = "") Then
                        'Check if it's a hardware component
                        If Not (swCustProp.Get("IS_HARDWARE") = "Yes" Or swCustProp.Get("Combine") = "No" Or InStr(Path, "\Hardwares\") <> 0) Then
            
                            'Debug.Print Path
                            
                            If swBomQuant.exists(Path) Then
                                swBomQuant.Item(Path) = swBomQuant.Item(Path) + 1
                            Else
                                swBomQuant.Add Path, 1
                            End If
                        Else
                            Debug.Print "Excluded Check 2: " & FileName(Path)
                        End If
                    
                    Else
                        Debug.Print "Excluded Check 1: " & FileName(Path)
                    End If
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


Function GetTransformMatrix(swThisBody As SldWorks.Body2, swOtherBody As SldWorks.Body2) As SldWorks.MathTransform

    Dim swTransform         As SldWorks.MathTransform
         
    If swThisBody.GetCoincidenceTransform2(swOtherBody, swTransform) Then
        
       Set GetTransformMatrix = swTransform
        
        If Not swTransform Is Nothing Then
            Debug.Print ""
            Debug.Print "------------------------------------"
            Debug.Print " Model 1    : " & ModelName1
            Debug.Print " Model 2    : " & ModelName2
            Debug.Print " Matrix     : " & True
        
            ' Create TDat to obtain Transform matrix data
            Dim TDat            As Variant
            Dim determinant     As Long
            
            TDat = swTransform.ArrayData
            
            ' Calculate Determinant
            determinant = TDat(0) * (TDat(4) * TDat(8) - TDat(5) * TDat(7)) - TDat(1) * (TDat(3) * TDat(8) - TDat(5) * TDat(6)) + TDat(2) * (TDat(3) * TDat(7) - TDat(4) * TDat(6))
            'Debug.Print "Determinant: " & Determinant
            
            'Debug.Print ""
            'Debug.Print " Rotation Matrix:"
            'Debug.Print vbTab & Round(TDat(0), 6), Round(TDat(1), 6), Round(TDat(2), 6)
            'Debug.Print vbTab & Round(TDat(3), 6), Round(TDat(4), 6), Round(TDat(5), 6)
            'Debug.Print vbTab & Round(TDat(6), 6), Round(TDat(7), 6), Round(TDat(8), 6)
            
            'Debug.Print "Translation:"
            'Debug.Print vbTab & Round(TDat(9), 4), Round(TDat(10), 4), Round(TDat(11), 4)
            
            'Debug.Print "Scaling: " & TDat(12)
            
            If determinant = -1 Then
                Debug.Print " Mirror     : " & ModelName1
                Debug.Print "------------------------------------"
            ' Check if scale is 1
            ElseIf TDat(12) = 1 Then
                Debug.Print " CombineID " & CombineID & ": " & ModelName1
                CanCoincide = True
            End If
              
        End If
         
    Else
        'Debug.Print "CANNOT COINCIDE with " & ModelName2
    End If
                    
End Function


Function GetBiggestBody(Part As SldWorks.PartDoc) As SldWorks.Body2

    Dim swBody                          As SldWorks.Body2
    Dim TempVolume                      As Double
    Dim BiggestBodyIndex                As Integer
    Dim MassProps                       As Variant
    Dim vBodies                         As Variant
    Dim n                               As Integer

    Part.ClearSelection2 True

    vBodies = Part.GetBodies2(swAllBodies, True)

    ' Get body properties to find biggest body
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

    ' Get Biggest Body
    Set GetBiggestBody = vBodies(BiggestBodyIndex)

End Function


Function FileName(file As String) As String

    Dim FullFileName As String
    
    FullFileName = Right(file, Len(file) - InStrRev(file, "\"))

    If Right(FullFileName, 7) = ".sldasm" Or Right(FullFileName, 7) = ".sldprt" Then
        FileName = Left(FullFileName, Len(FullFileName) - 7)
    Else
        FileName = FullFileName
    End If

End Function


Function NamePath(Name As String) As String

    NamePath = Replace(Name, "/", "@")

End Function


Function UpdateBar(Maxiter As Integer, ByVal caption As String)

    Dim CurrentProgress     As Double
    Dim BarWidth            As Double
    Dim ProgressPercentage  As Double
        
    If caption = "Rebuilding..." Then
        Progress = Progress - 1
        ProgressBar.Bar.BackColor = &HC000&
    End If
        
    Progress = Progress + 1
    Maxiter = Maxiter + 1 'Base 0
    
    ' Update Progress Bar
    CurrentProgress = Progress / Maxiter
    BarWidth = ProgressBar.Frame.Width * CurrentProgress
    ProgressPercentage = Round(CurrentProgress * 100, 0)
    ProgressBar.Bar.Width = BarWidth - 0.015 * BarWidth
    ProgressBar.Text2.caption = ProgressPercentage & "% Complete"
    ProgressBar.Text.caption = Progress & " of " & Maxiter
    ProgressBar.Text3.caption = caption
    
    ' If user uses stop button
    DoEvents
    
    If (ProgressPercentage / 10) Mod 2 = 0 Then
        ProgressBar.Image1.Visible = False
        ProgressBar.Image2.Visible = True
    Else
        ProgressBar.Image1.Visible = True
        ProgressBar.Image2.Visible = False
    End If

End Function


Function Compare(A As Variant, B As Variant) As Boolean

    Dim i As Integer
    For i = 0 To UBound(A)
        If A(i) <> B(i) Then
            Compare = False
            Exit Function
        End If
    Next i
    
    Compare = True

End Function


Function GetSymmetryType(Body As SldWorks.Body2) As String
    
    Dim swBody                  As SldWorks.Body2
    Dim MassProps               As Variant
    Dim CoM(2)                  As Variant
    Dim Box                     As Variant
    Dim BoxCoM(2)               As Double
    Dim BoxDiff(2)              As Double
    Dim i                       As Integer
    
    MassProps = Body.GetMassProperties(1)
    Box = Body.GetBodyBox
    
   ' Check differece between Body CoM and Box Com
    For i = 0 To 2
        CoM(i) = MassProps(i)
        BoxCoM(i) = (Box(i + 3) + Box(i)) / 2
        BoxDiff(i) = Round(BoxCoM(i) - CoM(i), 8)
    Next i
    
    If BoxDiff(0) = 0 And BoxDiff(1) = 0 And BoxDiff(2) = 0 And Box(3) - Box(0) = Box(4) - Box(1) Then
        GetSymmetryType = "Fully Symmetric"
        Exit Function
    End If
    
    If BoxDiff(0) = 0 And BoxDiff(1) = 0 And BoxDiff(2) = 0 Then
        GetSymmetryType = "Rotatable and Flippable"
        Exit Function
    End If
    
    If BoxDiff(0) = 0 And BoxDiff(1) = 0 And BoxDiff(2) <> 0 Then
        GetSymmetryType = "Rotatable"
        Exit Function
    End If

    GetSymmetryType = "Unique"

End Function


' NOT CURRENTLY USED
Function GroupPanels(List As Variant)
    Dim i               As Integer
    Dim Groups          As Variant
    Dim errors          As Long
    Dim warnings        As Long

    Dim swModel         As ModelDoc2
    Dim swModelDocExt   As ModelDocExtension
    Dim swCustProp      As CustomPropertyManager
    
    Dim FieldName       As String
    Dim UseCached       As Boolean
    Dim ValOut          As String
    Dim ResolvedValOut  As String
    Dim WasResolved     As Boolean
    Dim LinkToProperty  As Boolean


    Dim Length          As String
    Dim Width           As String
    Dim Thickness       As String
    
    For i = 0 To UBound(List)
        'Debug.Print List(i)
        Set swModel = swApp.OpenDoc6(List(i), swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)
        Set swModelDocExt = swModel.Extension
        Set swCustProp = swModelDocExt.CustomPropertyManager("")
        swCustProp.Get6 "Length", UseCached, ValOut, Length, WasResolved, LinkToProperty
        swCustProp.Get6 "Width", UseCached, ValOut, Width, WasResolved, LinkToProperty
        swCustProp.Get6 "Thickness", UseCached, ValOut, Thickness, WasResolved, LinkToProperty
        Debug.Print Length & " x " & Width & " x " & Thickness
        
        Groups = Length
        
    Next i

End Function










