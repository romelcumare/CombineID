'Change Properties


Option Explicit
Option Compare Text

Dim Properties As Variant

   Properties = Array(_
                Array("one" , "est"), _ 
                Array("one" , "est"), _ 
                Array("one" , "est"), _ 
                Array("one" , "est"), _ 
                Array("one" , "est"), _ 
                Array("one" , "est"), _ 
                Array("one" , "est"), _ 
                )
   
   
Dim swApp                   As SldWorks.SldWorks
Dim swModel                 As SldWorks.ModelDoc2
Dim swPart                  As SldWorks.PartDoc
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
Dim ModelName               As String 'Model being processed


Dim PartList                As Variant
Dim UserName                As String
Dim ErrorMessages           As Integer

Dim CurrentProgress         As Double
Dim BarWidth                As Long
Dim ProgressPercentage      As Double
Dim Progress                As Integer
Dim CancelButton            As Boolean


Sub main()

    Properties(0,0) = "IS_NESTING"
    Properties(0,1) = "IS_NESTING"
    Properties(1,0) = "IS_NESTING"
    Properties(1,1) = "IS_NESTING"
    Properties(2,0) = "IS_NESTING"
    Properties(2,1) = "IS_NESTING"
    Properties(3,0) = "IS_NESTING"
    Properties(3,1) = "IS_NESTING"

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
    
    'On Error GoTo ErrorHandler
    
    ' Create Log File
    DocName = swModel.GetTitle()
    DocPath = swModel.GetPathName

    If Right(DocName, 7) = ".sldasm" Then
        DocName = Left(DocName, InStrRev(DocName, ".") - 1)
    End If

    UserName = Environ("USERNAME")
    DocPath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") - 1)
    DocPath = DocPath + "\" + DocName + " - Properties Log.txt"
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
                           
    For i = 0 To UBound(PartList)
    
        Call UpdateBar(UBound(PartList), "Processing...")
   
         ChangeProperties (PartList(i))

         If CancelButton = True Then
             Debug.Print "CANCEL IS TRUE"
             fileStream.WriteLine "*----------- User Cancelled Macro -----------*"
             GoTo ExitCode
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
    Debug.Print "----------- ERROR while processing " & ModelName & "-----------"
    

    ' Close log file
    fileStream.Close
    
    MsgBox "Error Occured while processing: " & ModelName, vbExclamation, "Error"

End Sub

Sub ChangeProperties(Model As String)

    Dim swModel                        As SldWorks.ModelDoc2
    Dim swModelDocExt                  As ModelDocExtension
    Dim swModelDoc                     As SldWorks.ModelDoc2
    Dim swCustProp                     As CustomPropertyManager
    Dim ConfigPropMgr                    As SldWorks.CustomPropertyManager
   
    Dim longstatus                      As Long
    Dim longwarnings                    As Long
    Dim boolstatus                      As Boolean
    Dim i                               As Integer
    Dim j                               As Integer
    Dim n                               As Integer
    Dim retval                          As Long
    Dim lRetVal As Long
    Dim vPropNames As Variant
    Dim vPropTypes As Variant
    Dim vPropValues As Variant
    Dim resolved As Variant
    Dim linkProp As Variant
    Dim nNbrProps As Long
    Dim custPropType As Long

    ModelName = FileName(Model)
    Set swModel = swApp.OpenDoc6(Model, 1, 0, "", longstatus, longwarnings)
    Set swPart = swModel
    Set swModelDoc = swModel
    Set swModelDocExt = swModelDoc.Extension
    Set ConfigPropMgr = swModel.GetActiveConfiguration.CustomPropertyManager
    Set swCustProp = swModelDocExt.CustomPropertyManager("")

    ' Get the number of custom properties for this configuration
    nNbrProps = swCustProp.Count
    Debug.Print "Number of properties for this configuration:            " & nNbrProps

    ' Gets the custom properties
    lRetVal = swCustProp.GetAll3(vPropNames, vPropTypes, vPropValues, resolved, linkProp)
    
    ' For each custom property, print its name, type, and evaluated value
    Debug.Print "-----------------"
    For j = 0 To nNbrProps - 1
        custPropType = swCustProp.GetType2(vPropNames(j))
        
'        if vPropNames = "IS_NESTING" or
        
        Debug.Print "    Name, swCustomInfoType_e value, and resolved value:  " & vPropNames(j) & ", "; custPropType & ", " & vPropValues(j)
    Next j



End Sub

Function ChangeValue()


End Function
Function GetParts(Assembly As SldWorks.AssemblyDoc) As Variant

    Dim swBomQuant As Object 'Key->Path, Value->Quantity
    Dim vComps As Variant
    Dim swComp As SldWorks.Component2
    Dim i As Integer
    Dim Path As String
    Dim swSelModel                      As SldWorks.ModelDoc2
    Dim swModelDocExt                   As ModelDocExtension
    
    Set swBomQuant = CreateObject("Scripting.Dictionary")

    vComps = Assembly.GetComponents(False)
    
    For i = 0 To UBound(vComps)
        
        Set swComp = vComps(i)
        'Debug.Print vComps(i).Name
        
        ' Check if suppresed and if it's part model
        If swComp.IsSuppressed = False Then
        
            Set swSelModel = swComp.GetModelDoc2
            
            If swSelModel.GetType = 1 Then
                         
                'Debug.Print swSelModel.GetType
                'Debug.Print swComp.Name
                'Debug.Print swComp.IsSuppressed
                
                Path = swComp.GetPathName
                'Debug.Print Path
                                
                If swBomQuant.exists(Path) Then
                    swBomQuant.Item(Path) = swBomQuant.Item(Path) + 1
                Else
                    swBomQuant.Add Path, 1
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















