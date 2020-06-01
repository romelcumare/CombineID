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

Dim CurrentProgress         As Double
Dim BarWidth                As Long
Dim ProgressPercentage      As Double
Dim Progress                As Integer
Dim CancelButton            As Boolean



Sub main()

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
    
    'On Error GoTo ErrorHandler
    
    'Create Log File
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
    
    'Deactivate view update
    Dim modView As ModelView
    Set modView = swModel.ActiveView
    modView.EnableGraphicsUpdate = False
    swAssy.FeatureManager.EnableFeatureTree = False
    
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
                
    PartList = GetParts(swAssy)
    
    NumofParts = UBound(PartList) + 1
    'Debug.Print "Number of parts: " & NumofParts
    
    NumofChecks = (NumofParts * (NumofParts - 1)) / 2
    
    'Set Combine IDs
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

        'Increse GlobalCombineID is coincide was found
        If CanCoincideGlobal Then
            GlobalCombineID = GlobalCombineID + 1
            CanCoincideGlobal = False
        End If

    Next i
    
    Call UpdateBar(UBound(PartList), "Processing...")
    
ExitCode:

    Call UpdateBar(UBound(PartList), "Rebuilding...")

    Set swModel = swApp.ActivateDoc3(AssemblyPath, False, 1, 0)

    'Rebuild Assembly
    boolstatus = swModel.EditRebuild3()

    swModel.ClearSelection2 (True)

    'Activate view update
    modView.EnableGraphicsUpdate = True
    swAssy.FeatureManager.EnableFeatureTree = True

    Unload ProgressBar

    Debug.Print ""
    Debug.Print "----------- Macro Finished -----------"
    fileStream.WriteLine "-----------       Macro Finished       -----------"

    'Close log file
    fileStream.Close
    
    'Exit Sub and don't proceed to ErrorHandler
    Exit Sub

ErrorHandler:

    Call UpdateBar(UBound(PartList), "Rebuilding...")

    Set swModel = swApp.ActivateDoc3(AssemblyPath, False, 1, 0)

    'Rebuild Assembly
    boolstatus = swModel.EditRebuild3()

    swModel.ClearSelection2 (True)

    'Activate view update
    modView.EnableGraphicsUpdate = True
    swAssy.FeatureManager.EnableFeatureTree = True

    Unload ProgressBar

    Debug.Print ""
    Debug.Print "----------- ERROR while processing " & ModelName1 & "-----------"
    fileStream.WriteLine "----------- ERROR while processing " & ModelName1 & "-----------"

    'Close log file
    fileStream.Close
    
    MsgBox "Error Occured while processing: " & ModelName1, vbExclamation, "Error"

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
    Dim FaceNormal                      As Variant
    Dim longstatus                      As Long
    Dim longwarnings                    As Long
    Dim boolstatus                      As Boolean
    Dim RotMatrix                       As SldWorks.MathTransform
    Dim i                               As Integer
    Dim j                               As Integer
    Dim n                               As Integer
    Dim SameLamEdge                     As Boolean
    Dim RetVal                          As Long
    
    ''Get Part1 Data
    'Debug.Print ""
    'Debug.Print "Processing:  " & FileName(Model1)
    
    'Rest CanCoincide
    CanCoincide = False
    
    ModelName1 = FileName(Model1)
    Set swModel1 = swApp.OpenDoc6(Model1, 1, 0, "", longstatus, longwarnings)
    Set swPart1 = swModel1
    Set swModelDoc1 = swModel1
    Set swModelDocExt1 = swModelDoc1.Extension
    Material1 = swModelDoc1.MaterialIdName
    'Debug.Print "Material:  " & swModelDoc1.MaterialIdName
    Set swCustProp1 = swModelDocExt1.CustomPropertyManager("")

    'Check if already has CombineID and use it
    If swCustProp1.Get("CombineID") = "" Then
        'Potentially use GlobalCombineID if coincidense is found
        CombineID = GlobalCombineID
    Else
        'Grab current CombineID
        'Debug.Print ModelName1 & ": Already has CombineID"
        CombineID = swCustProp1.Get("CombineID")
    End If

    ''Get Part2 Data
    'Debug.Print " Comparing:  " & FileName(Model2)

    ModelName2 = FileName(Model2)
    Set swModel2 = swApp.OpenDoc6(Model2, 1, 0, "", longstatus, longwarnings)
    Set swPart2 = swModel2
    Set swModelDoc2 = swModel2
    Set swModelDocExt2 = swModelDoc2.Extension
    Material2 = swModelDoc2.MaterialIdName
    'Debug.Print "Material:  " & swModelDoc2.MaterialIdName
    Set swCustProp2 = swModelDocExt2.CustomPropertyManager("")

    'Check if Part2 already has CombineID and skip
    If swCustProp2.Get("CombineID") <> "" Then
        'Debug.Print ModelName2 & ": Already has CombineID and will not be checked"
        Exit Sub
    End If

    'Check if both Parts have the same material. Skip coincidence check if not
    If Not Material1 = Material2 Then
        Exit Sub
    End If

    'Get Biggest Body1
    Set swBody1 = GetBiggestBody(swPart1)
    'Debug.Print "Biggest Body1 is: " & swBody1.Name

    'Get Biggest Body2
    Set swBody2 = GetBiggestBody(swPart2)
    'Debug.Print "Biggest Body2 is: " & swBody2.Name
        
    'Get rotation matrix if bodies can coincide
    Set RotMatrix = GetTransformMatrix(swBody2, swBody1)
    
    'Check if panels have the same number of laminates and edge
    SameLamEdge = CheckProperties(swModel1, swModel2)
    
    If (CanCoincide And SameLamEdge) Then
    
        CanCoincideGlobal = True
       
    
        'Add CombineID custom property to both parts
        Set swCustProp2 = swModelDocExt2.CustomPropertyManager("")
        
        RetVal = swCustProp1.Add3("CombineID", swCustomInfoText, CombineID, swCustomPropertyReplaceValue)
        RetVal = swCustProp2.Add3("CombineID", swCustomInfoText, CombineID, swCustomPropertyReplaceValue)
        
        
        'Determine if compared body was flipped
        Set swFace = SelectLargestFace(swBody2)
        
        FaceNormal = swFace.Normal
        'Debug.Print "Largest Face Normal Vector = (" & FaceNormal(0) & ", " & FaceNormal(1) & ", " & FaceNormal(2) & ")"
        
        'Multiply face vector and rotation matrix to determine if body was flipped
        Dim swMath As SldWorks.MathUtility
        Dim RotatedVector As MathVector
        
        Set swMath = swApp.GetMathUtility
        Set RotatedVector = swMath.CreateVector(FaceNormal)
        Debug.Print "FLIPPED Panel"
            
       
        'Flag models as dirty
        swPart1.SetSaveFlag
        swPart2.SetSaveFlag
        'Debug.Print " is it Dirty? " & swPart1.GetSaveFlag
        'Debug.Print " is it Dirty? " & swPart2.GetSaveFlag
        
        
        
    End If
            
End Sub

Function CheckProperties(swModel1 As SldWorks.ModelDoc2, swModel2 As SldWorks.ModelDoc2) As Boolean

    Dim config1             As SldWorks.Configuration
    Dim config2             As SldWorks.Configuration
    Dim cusPropMgr1         As SldWorks.CustomPropertyManager
    Dim cusPropMgr2         As SldWorks.CustomPropertyManager
    Dim lRetVal1            As Long
    Dim lRetVal2            As Long
    Dim nNbrProps1          As Long
    Dim nNbrProps2          As Long
    Dim j                   As Integer
    
    Dim vPropNames1         As Variant
    Dim vPropTypes1         As Variant
    Dim vPropValues1        As Variant
    Dim ValOut1             As String
    Dim ResolvedValOut1     As String
    Dim wasResolved1        As Boolean
    Dim linkToProp1         As Boolean
    Dim resolved1           As Variant
    Dim linkProp1           As Variant
    
    Dim vPropNames2         As Variant
    Dim vPropTypes2         As Variant
    Dim vPropValues2        As Variant
    Dim ValOut2             As String
    Dim ResolvedValOut2     As String
    Dim wasResolved2        As Boolean
    Dim linkToProp2         As Boolean
    Dim resolved2           As Variant
    Dim linkProp2           As Variant
    
    Dim NumofEdgebands1     As Integer
    Dim NumofEdgebands2     As Integer
    Dim NumofLaminates1     As Integer
    Dim NumofLaminates2     As Integer

    Dim i                   As Integer
    
    Set config1 = swModel1.GetActiveConfiguration
    Set config2 = swModel2.GetActiveConfiguration
    
    Set cusPropMgr1 = config1.CustomPropertyManager
    Set cusPropMgr2 = config2.CustomPropertyManager

    nNbrProps1 = cusPropMgr1.Count
    nNbrProps2 = cusPropMgr2.Count
    
    ' Gets the custom properties
    lRetVal1 = cusPropMgr1.GetAll3(vPropNames1, vPropTypes1, vPropValues1, resolved1, linkProp1)
    lRetVal2 = cusPropMgr2.GetAll3(vPropNames2, vPropTypes2, vPropValues2, resolved2, linkProp2)

    ' For each custom property, print its name, type, and evaluated value
    For j = 0 To nNbrProps1 - 1
        'Debug.Print "    Name1: " & vPropNames1(j) & " Value1: " & vPropValues1(j)

        If InStr(vPropValues1(j), "Edgeband") > 0 Then
            NumofEdgebands1 = NumofEdgebands1 + 1
        End If
        
    Next j
      
    For j = 0 To nNbrProps2 - 1
        'Debug.Print "    Name2: " & vPropNames2(j) & " Value2: " & vPropValues2(j)

        If InStr(vPropValues2(j), "Edgeband") > 0 Then
            NumofEdgebands2 = NumofEdgebands2 + 1
        End If

    Next j
    
    If cusPropMgr1.Get("SWOODCP_TopStockMaterial") <> "" Then
        NumofLaminates1 = NumofLaminates1 + 1
    End If
    
    If cusPropMgr1.Get("SWOODCP_BottomStockMaterial") <> "" Then
        NumofLaminates1 = NumofLaminates1 + 1
    End If
    
      If cusPropMgr2.Get("SWOODCP_TopStockMaterial") <> "" Then
        NumofLaminates2 = NumofLaminates2 + 1
    End If
    
    If cusPropMgr2.Get("SWOODCP_BottomStockMaterial") <> "" Then
        NumofLaminates2 = NumofLaminates2 + 1
    End If
    
    'Debug.Print "Number of Edgebands 1: " & NumofEdgebands1
    'Debug.Print "Number of Edgebands 2: " & NumofEdgebands2
    'Debug.Print "Number of Laminates 1: " & NumofLaminates1
    'Debug.Print "Number of Laminates 2: " & NumofLaminates2
    
    If NumofEdgebands1 = NumofEdgebands2 And NumofLaminates1 = NumofLaminates2 Then
        CheckProperties = True
    End If
    
    'Debug.Print "CheckProperties result: " & CheckProperties
           
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
        
        'Check if suppresed and if it's part model
        If swComp.IsSuppressed = False Then
        
            Set swSelModel = swComp.GetModelDoc2
            
            If swSelModel.GetType = 1 Then
          
                Set swModelDocExt = swSelModel.Extension
                Set swCustProp = swModelDocExt.CustomPropertyManager("")
               
                'Debug.Print swSelModel.GetType
                'Debug.Print swComp.Name
                'Debug.Print swComp.IsSuppressed
                
                Path = swComp.GetPathName
                
                    'Delete CombineID if it exists
                    boolstatus = swCustProp.Delete2("CombineID")
                
                    'Check if it has LxWxT properties
                    If Not (swCustProp.Get("Length") = "" Or swCustProp.Get("Width") = "" Or swCustProp.Get("Thickness") = "") Then
                        'Check if it's a hardware component
                        If Not (swCustProp.Get("IS_HARDWARE") = "Yes" Or InStr(Path, "\Hardwares\") <> 0) Then
            
                            'Path = swComp.GetPathName
                
                            If swBomQuant.exists(Path) Then
                                swBomQuant.Item(Path) = swBomQuant.Item(Path) + 1
                            Else
                                swBomQuant.Add Path, 1
                            End If
                            
                        End If
                    
                    Else
                        'Debug.Print "Excluded: " & FileName(Path)
                    End If
            End If
        End If
        
    Next
    
    Dim vItems As Variant
    vItems = swBomQuant.Keys
    
    'For i = 0 To UBound(vItems)
        'Debug.Print vItems(i) & ", " & swBomQuant.Item(vItems(i))
    'Next

    GetParts = vItems
    
End Function



Function GetTransformMatrix(swThisBody As SldWorks.Body2, swOtherBody As SldWorks.Body2) As SldWorks.MathTransform

    Dim swTransform         As SldWorks.MathTransform
         
    If swThisBody.GetCoincidenceTransform2(swOtherBody, swTransform) Then
        
       Set GetTransformMatrix = swTransform
        
        If Not swTransform Is Nothing Then
            Debug.Print ""
            Debug.Print "------------------------------------"
            Debug.Print " Matrix     : " & ModelName2
        
            'Create vXfm only to print the transformation
            Dim vXfm As Variant
            Dim Determinant As Long
            
            'Angles calculator
            'https://www.andre-gaschler.com/rotationconverter/
            vXfm = swTransform.ArrayData
            
            'Calculate Determinant
            Determinant = vXfm(0) * (vXfm(4) * vXfm(8) - vXfm(5) * vXfm(7)) - vXfm(1) * (vXfm(3) * vXfm(8) - vXfm(5) * vXfm(6)) + vXfm(2) * (vXfm(3) * vXfm(7) - vXfm(4) * vXfm(6))
            'Debug.Print "Determinant: " & Determinant
            'Debug.Print ""

            'Debug.Print "Rotation:"
            'Debug.Print vbTab & Round(vXfm(0), 4), Round(vXfm(1), 4), Round(vXfm(2), 4)
            'Debug.Print vbTab & Round(vXfm(3), 4), Round(vXfm(4), 4), Round(vXfm(5), 4)
            'Debug.Print vbTab & Round(vXfm(6), 4), Round(vXfm(7), 4), Round(vXfm(8), 4)
            'Debug.Print "Translation:"
            'Debug.Print vbTab & Round(vXfm(9), 4), Round(vXfm(10), 4), Round(vXfm(11), 4)
            'Debug.Print "Scaling: " & vXfm(12)
            
            If Determinant = -1 Then
                Debug.Print " Mirror     : " & ModelName1
            ElseIf vXfm(12) = 1 Then 'Check if scale is 1
                Debug.Print " CombineID " & CombineID & ": " & ModelName1
                CanCoincide = True
            End If
            Debug.Print "------------------------------------"
            
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
    Set GetBiggestBody = vBodies(BiggestBodyIndex)

End Function


Function SelectLargestFace(swBody As SldWorks.Body2) As SldWorks.Entity
     
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


Function UpdateBar(Maxiter As Integer, ByVal Caption As String)

    Dim CurrentProgress     As Double
    Dim BarWidth            As Double
    Dim ProgressPercentage  As Double
        
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

End Function




