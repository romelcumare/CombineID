Option Explicit
Option Base 1




Sub Main()

    Dim SwApp                   As SldWorks.SldWorks
    Dim swModel                 As SldWorks.ModelDoc2

    

    Dim errCode                 As Long
    Dim el                      As Double
    Dim tl                      As Double

    Dim Disp  As Variant, Stress As Variant, Strn As Variant
   

    Dim SelNodeElemVariant      As Variant
    Dim SelNodeElemWarnings     As Variant

    Dim Name                    As Variant
    Dim NumPlots                As Integer
    Dim PlotNames()             As String
    Dim StudyName               As String
    Dim ActStudy                As Integer
    
    Dim Filepath                As String
    Dim PlotFolderPath          As String
    Dim i                       As Integer
    Dim j                       As Integer
    Dim k                       As Integer
    Dim longstatus              As Long
    Dim Part                    As Object
    Dim swModelView             As SldWorks.ModelView
    Dim vModelViewNames         As Variant
    Dim CustomViewNames()       As String
    Dim NumCustomViews          As Integer
    Dim TotalNumViews           As Integer
    Dim TotalNumImages          As Integer
     
    Dim CurrentProgress         As Double
    Dim ProgressPercentage      As Double
    Dim BarWidth                As Long
    
    Dim swModelDocExt As SldWorks.ModelDocExtension

    '==========================================================================
    
    'Get SolidWorks
    If SwApp Is Nothing Then Set SwApp = Application.SldWorks
    
    'Get Active Document
    Set Part = SwApp.ActiveDoc
    
   
    Set swModelDocExt = Part.Extension
   
 

    'Get Model File Path
    Set swModel = SwApp.ActiveDoc
    Filepath = Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\") - 1) & "\"
    Debug.Print "FileName = " & swModel.GetPathName
    PlotFolderPath = Filepath & "Result Plots\"
  
    Dim ViewData As Variant
    
         
    'Get model views
    NumCustomViews = 0

    
    ViewData = Part.GetStandardViewRotation(1)
    
    For i = 0 To UBound(ViewData)
        Debug.Print ViewData(i)
    Next i
    
    If ViewData(0) = 1 And ViewData(4) = 1 And ViewData(8) = 1 Then
        Debug.Print "Correct view orientation"
    Else
         Debug.Print "WRONG view orientation"
    End If

    
    
End Sub
    