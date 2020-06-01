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
Dim swAssyDocExt            As SldWorks.ModelDocExtension
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


Public CurrentProgress      As Double
Public BarWidth             As Long
Public ProgressPercentage   As Double
Public Progress             As Integer
Public CancelButton         As Boolean



Sub main()

    Debug.Print ""
    Debug.Print "----------- Macro Started -----------"


    Dim swConfMgr               As SldWorks.ConfigurationManager
    Dim swConf                  As SldWorks.Configuration
    Dim i                       As Integer
    Dim j                       As Integer
    Dim NumofParts              As Integer
    Dim NumofChecks             As Integer
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
    DocPath = DocPath + "\" + DocName + "-Combine Log.txt"
    Set fso = New FileSystemObject
    Set fileStream = fso.OpenTextFile(DocPath, ForAppending, True)
    fileStream.WriteLine ""
    fileStream.WriteLine "----------- " & Now & " -----------"
  
    
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
    
    Set swConfMgr = swModel.ConfigurationManager
    Set swConf = swConfMgr.ActiveConfiguration
            
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
                fileStream.WriteLine "User Cancelled Macro"
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

    'Rebuild Assembly
    'boolstatus = swModel.ForceRebuild3(True)

    Unload ProgressBar


    Debug.Print ""
    Debug.Print "----------- Macro Finished -----------"
    fileStream.WriteLine "----------- Macro Finished -----------"

    'Close log file
    fileStream.Close

    ' ----------- Exit Sub and don't proceed to ErrorHandler -----------
    Exit Sub

ErrorHandler:

    'Activate view update
    modView.EnableGraphicsUpdate = True
    swAssy.FeatureManager.EnableFeatureTree = True

    'Rebuild Assembly
    'boolstatus = swModel.ForceRebuild3(True)

    Unload ProgressBar

    Debug.Print ""
    Debug.Print "----------- ERROR while processing " & ModelName1 & "-----------"
    fileStream.WriteLine "----------- ERROR while processing " & ModelName1 & "-----------"

    'Close log file
    fileStream.Close

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



    ''Get Part1 Data
    'Debug.Print ""
    'Debug.Print "Processing:  " & FileName(Model1)

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
        CombineID = GlobalCombineID
    Else
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
        'Debug.Print ModelName2 & ": Already has CombineID"
        Exit Sub
    End If

    'Check if both Parts have the same material
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

    If CanCoincide Then
        'Add CombineID custom property to both parts
        Set swCustProp2 = swModelDocExt2.CustomPropertyManager("")
        
        boolstatus = swCustProp1.Add3("CombineID", swCustomInfoText, CombineID, swCustomPropertyReplaceValue)
        boolstatus = swCustProp2.Add3("CombineID", swCustomInfoText, CombineID, swCustomPropertyReplaceValue)
        
        'Determine if compared body was flipped
        Set swFace = SelectLargestFace(swBody2)
        
        FaceNormal = swFace.Normal
        'Debug.Print "Largest Face Normal Vector = (" & FaceNormal(0) & ", " & FaceNormal(1) & ", " & FaceNormal(2) & ")"
        
        'Multiply face vector and rotation matrix to determine if body was flipped
            Dim swMath As SldWorks.MathUtility
            Dim RotatedVector As MathVector
            
            Set swMath = swApp.GetMathUtility
            Set RotatedVector = swMath.CreateVector(FaceNormal)
            Set RotatedVector = RotatedVector.MultiplyTransform(RotMatrix)
            'Debug.Print "Rotated vector = (" & RotatedVector.ArrayData(0) & ", " & RotatedVector.ArrayData(1) & ", " & RotatedVector.ArrayData(2) & ")"

            If RotatedVector.ArrayData(2) = -FaceNormal(2) Then
                Debug.Print "FLIPPED Panel"
            End If
        
        'Flag models as dirty
        swPart1.SetSaveFlag
        swPart2.SetSaveFlag
        'Debug.Print " is it Dirty? " & swPart1.GetSaveFlag
        'Debug.Print " is it Dirty? " & swPart2.GetSaveFlag

        CanCoincide = False
    End If
            
End Sub


Function GetParts(Assembly As SldWorks.AssemblyDoc) As Variant

    Dim swBomQuant As Object 'Key->Path, Value->Quantity
    Dim vComps As Variant
    Dim swComp As SldWorks.Component2
    Dim i As Integer
    Dim Path As String
    Dim Desc As String
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
                CanCoincideGlobal = True
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


Public Function UpdateBar(Maxiter As Integer, ByVal Caption As String)
    Dim CurrentProgress     As Double
    Dim BarWidth            As Double
    Dim ProgressPercentage  As Double
    Dim A                   As Integer
    
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

End Function







