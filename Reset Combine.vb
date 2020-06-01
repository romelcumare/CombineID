Option Explicit

Dim swApp               As SldWorks.SldWorks
Dim swModel             As SldWorks.ModelDoc2
Dim swAssy              As SldWorks.AssemblyDoc
Dim boolstatus          As Boolean
Dim AssemblyName        As String
Dim swAssyDocExt        As SldWorks.ModelDocExtension
Dim AssemblyPath        As String

Public CurrentProgress      As Double
Public BarWidth             As Long
Public ProgressPercentage   As Double
Public Progress             As Integer
Public Cancel               As Boolean

' --------- RESET ---------
Sub main() 

    Dim swConfMgr As SldWorks.ConfigurationManager
    Dim PartList As Variant
    'Dim swAssy As SldWorks.AssemblyDoc
    Dim i As Integer
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swAssyDocExt = swModel.Extension
    Set swAssy = swApp.ActiveDoc
    AssemblyPath = swModel.GetPathName
    AssemblyName = FileName(swModel.GetTitle())
    
    
    'Deactivate view update
    Dim modView As ModelView
    Set modView = swModel.ActiveView
    modView.EnableGraphicsUpdate = False
    swAssy.FeatureManager.EnableFeatureTree = False
  
    
    'Reset Origin
    Progress = 0

  
    If swModel.GetType = SwConst.swDocASSEMBLY Then

        Set swConfMgr = swModel.ConfigurationManager
        Set swAssy = swModel
        
        PartList = GetParts(swAssy)
               
        
    Else
        MsgBox "The active document is not an assembly model.", vbOKOnly, "Operational Feedback"
    End If

ExitCode:
    
    Set swModel = swApp.ActivateDoc3(AssemblyPath, False, 1, 0)

    Debug.Print "Rebuilding Model: " & swModel.GetTitle
    
    boolstatus = swModel.Extension.Rebuild(2)
    Debug.Print boolstatus
    
    
    'Activate view update
    modView.EnableGraphicsUpdate = True
    swAssy.FeatureManager.EnableFeatureTree = True
    
    'Rebuild Assembly
    boolstatus = swModel.EditRebuild3()
    Debug.Print "Rebuild status: " & boolstatus
    Debug.Print "Reset Finished"

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
        
        'Check if suppresed and is part model
        If swComp.IsSuppressed = False Then
        
            Set swSelModel = swComp.GetModelDoc2
            
            
            'If swComp.IsSuppressed = False And swSelModel.GetType = 1 Then
            If swSelModel.GetType = 1 Then
        
            
            Set swModelDocExt = swSelModel.Extension
            Set swCustProp = swModelDocExt.CustomPropertyManager("")
        
            
            'Debug.Print swSelModel.GetType
            'Debug.Print swComp.Name
            'Debug.Print swComp.IsSuppressed
            
            Path = swComp.GetPathName
 
            'Delete CombineID if it exists
            boolstatus = swCustProp.Delete2("CombineID")
            
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


Function FileName(File As String) As String

    Dim FullFileName As String
    
    FullFileName = Right(File, Len(File) - InStrRev(File, "\"))

    If Right(FullFileName, 7) = ".sldasm" Or Right(FullFileName, 7) = ".sldprt" Then
        FileName = Left(FullFileName, Len(FullFileName) - 7)
    Else
        FileName = FullFileName
    End If

End Function


