Option Explicit
Option Compare Text
Const pi As Double = 3.14159265358979
Const Precision As Integer = 4
Dim qcd7e5646aa31d4a3b54db53750164e31                   As SldWorks.SldWorks
Dim r089f2e75bc0c2785f3502a8c94ce64fb                 As SldWorks.ModelDoc2
Dim b67c3f8109b77bd4b38e5630fd32c514f                  As SldWorks.PartDoc
Dim z4f76b89c98a78e0601bd278fdabfe3ab                  As SldWorks.AssemblyDoc
Dim bf79075ea2e17d0bf2eb4ec624b45fd6f                  As SldWorks.Component2
Dim wf0ca045f9be481927e0f67dfe17d1b77               As SldWorks.SelectData
Dim w33548d5cda925bb321073eb295ef6a62                  As SldWorks.Body2
Dim bfd8c89eae9803e5f705c3dd395249328                 As Variant
Dim bf49699a88461a2a540df64abb48b2f98              As Long
Dim b4e5c40e176840ea491c358d7b078c013              As Boolean
Dim b94b7ebfc3d693c01e015c80ebfdd0725            As String
Dim w7a695257e7af03642eb7fff3efb75bb3            As SldWorks.ModelDocExtension
Dim q02ed784cf0fc1dc30d0c7933763a602c            As String
Dim z56989495ec6c0363b02cdbc53cb1808d                 As String
Dim b383222a9bea4084b4a55b3485db67e60                 As String
Dim wfe98125bc4fc31287d895dbdb361bfb9            As String
Dim rab2bd3733535f62d8ead8ad301e13be2                     As FileSystemObject
Dim bfc5d43b8304fa915bdf038e967c4d02a              As TextStream
Public t382b9d8066debe4514dc179247019dcf      As Double
Public r2e00027eabe0b722f19a23ac90779560             As Long
Public qddcbd44bfca946d6461895d06978f141   As Double
Public qea414b44a6ee13e50a6ec2eb47ffefa7             As Integer
Public w9ef73b569aa1335c16e2f2abb37559eb         As Boolean
Dim t0f0bc4817eb80edbd3b60e38d6e44c34          As Variant
Dim n88c5392f2cfe40f6063575c761ab6c7d               As String
Sub main()
Dim w93baea52d603ba1f44a0151d3fcffce7               As SldWorks.ConfigurationManager
Dim b16c335cf956414eff434e97cc74587f5                  As SldWorks.Configuration
Dim bb14ddfe535e39e16ec3dd3b9f42fb988              As SldWorks.Component2
Dim r910c0775bffa167dbee861a2869f037e                As Variant
Dim z57fbbe9a55b7e76e8772bb12c27d0537                       As Integer
Set qcd7e5646aa31d4a3b54db53750164e31 = Application.SldWorks
Set r089f2e75bc0c2785f3502a8c94ce64fb = qcd7e5646aa31d4a3b54db53750164e31.ActiveDoc
Set w7a695257e7af03642eb7fff3efb75bb3 = r089f2e75bc0c2785f3502a8c94ce64fb.Extension
q02ed784cf0fc1dc30d0c7933763a602c = b5ebaf9eab0956be7d743ef1391203e5e(r089f2e75bc0c2785f3502a8c94ce64fb.GetTitle())
If Not r089f2e75bc0c2785f3502a8c94ce64fb.GetType = SwConst.swDocASSEMBLY Then
MsgBox "The active document is not an assembly model.", vbOKOnly
End
End If
z56989495ec6c0363b02cdbc53cb1808d = r089f2e75bc0c2785f3502a8c94ce64fb.GetTitle()
b383222a9bea4084b4a55b3485db67e60 = r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName
If Right(z56989495ec6c0363b02cdbc53cb1808d, 7) = ".sldasm" Then
z56989495ec6c0363b02cdbc53cb1808d = Left(z56989495ec6c0363b02cdbc53cb1808d, InStrRev(z56989495ec6c0363b02cdbc53cb1808d, ".") - 1)
End If
b383222a9bea4084b4a55b3485db67e60 = Left(r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName, InStrRev(r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName, "\") - 1)
b383222a9bea4084b4a55b3485db67e60 = b383222a9bea4084b4a55b3485db67e60 + "\" + z56989495ec6c0363b02cdbc53cb1808d + "-LogFile.txt"
Set rab2bd3733535f62d8ead8ad301e13be2 = New FileSystemObject
Set bfc5d43b8304fa915bdf038e967c4d02a = rab2bd3733535f62d8ead8ad301e13be2.OpenTextFile(b383222a9bea4084b4a55b3485db67e60, ForAppending, True)
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine ""
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "----------- " & Now & " -----------"
Set z4f76b89c98a78e0601bd278fdabfe3ab = r089f2e75bc0c2785f3502a8c94ce64fb
b94b7ebfc3d693c01e015c80ebfdd0725 = r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName
t0f0bc4817eb80edbd3b60e38d6e44c34 = z4f76b89c98a78e0601bd278fdabfe3ab.GetComponents(False)
qea414b44a6ee13e50a6ec2eb47ffefa7 = 0
With ProgressBar
.Bar.Width = 0
.Text.Caption = "Getting Parts..."
.Text2.Caption = "0% Complete"
.Text3.Caption = "Processing..."
.Show vbModeless
End With
w9ef73b569aa1335c16e2f2abb37559eb = False
Set w93baea52d603ba1f44a0151d3fcffce7 = r089f2e75bc0c2785f3502a8c94ce64fb.ConfigurationManager
Set b16c335cf956414eff434e97cc74587f5 = w93baea52d603ba1f44a0151d3fcffce7.ActiveConfiguration
Set bb14ddfe535e39e16ec3dd3b9f42fb988 = b16c335cf956414eff434e97cc74587f5.GetRootComponent3(True)
r910c0775bffa167dbee861a2869f037e = m1e98ad36788ac2ca18eb096d3a858f8e(z4f76b89c98a78e0601bd278fdabfe3ab)
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(r910c0775bffa167dbee861a2869f037e)
bf96188166df987c66304fae9a91c65e3 (r910c0775bffa167dbee861a2869f037e(z57fbbe9a55b7e76e8772bb12c27d0537))
Call q8b61244ae4822e03b47ad99c4fc71e0c(UBound(r910c0775bffa167dbee861a2869f037e), "Processing...")
If w9ef73b569aa1335c16e2f2abb37559eb = True Then
Debug.Print "CANCEL IS TRUE"
Debug.Print ""
Debug.Print ""
Debug.Print ""
Debug.Print ""
Debug.Print ""
Debug.Print ""
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "User Cancelled Macro"
GoTo ExitCode
End If
Next z57fbbe9a55b7e76e8772bb12c27d0537
ExitCode:
Call q8b61244ae4822e03b47ad99c4fc71e0c(UBound(r910c0775bffa167dbee861a2869f037e), "Rebuilding...")
Set r089f2e75bc0c2785f3502a8c94ce64fb = qcd7e5646aa31d4a3b54db53750164e31.ActivateDoc3(b94b7ebfc3d693c01e015c80ebfdd0725, False, 1, 0)
r089f2e75bc0c2785f3502a8c94ce64fb.ClearSelection2 (True)
b4e5c40e176840ea491c358d7b078c013 = r089f2e75bc0c2785f3502a8c94ce64fb.Extension.Rebuild(1)
Unload ProgressBar
Debug.Print ""
Debug.Print "----------- Macro Finished -----------"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "----------- Macro Finished -----------"
bfc5d43b8304fa915bdf038e967c4d02a.Close
Exit Sub
ErrorHandler:
Unload ProgressBar
Debug.Print ""
Debug.Print "----------- Macro Terminated while processing " & n88c5392f2cfe40f6063575c761ab6c7d & "-----------"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "----------- Macro Terminated while processing " & n88c5392f2cfe40f6063575c761ab6c7d & "-----------"
bfc5d43b8304fa915bdf038e967c4d02a.Close
End Sub
Sub bf96188166df987c66304fae9a91c65e3(ByRef model As String)
Dim be5418de9d969c5ca5cb44308f32da188                       As Object
Dim b67927ea18feebbaa0a49705f4fec3eb2       As Object
Dim b432b23cd29982f1c659de7c15266a10b                       As Object
Dim n5365312f3a16ea17d2b92ef20fa77d84                   As ModelDocExtension
Dim bbeb837e9ee564e64f8607696461893a4                      As SldWorks.MassProperty
Dim wc8c39c5599ddf73063181785c0caae37                      As SldWorks.ModelDoc2
Dim b75677af635f8cc2aab0c89a1ddfa26bb                      As CustomPropertyManager
Dim b26226ef0ed8b85195a2984e0f7903a1c                            As Variant
Dim t8fcc0989bc0894d1dac35c7f5e7d32df(2)                        As Variant
Dim q8e5489469f46ff6977f70647e6d160e6                           As Double
Dim qc459f7414f9a330324e431ffa95a66bb                          As Double
Dim b9b4c91e56a7252144a2745fe17913d02                           As Double
Dim b8cdfda0b5a2e71faa28ead12f563eb75                       As SldWorks.FeatureManager
Dim b1596bdb5b54bd50c7bd07ea6d9da6316                          As SldWorks.Feature
Dim bdcc132ec89e911e347bcd7c6e05d44ad                         As Double
Dim bfd8c89eae9803e5f705c3dd395249328                         As Variant
Dim w33548d5cda925bb321073eb295ef6a62                          As SldWorks.Body2
Dim b679764c23bf7d8a169575ebac35213ae                  As SldWorks.Body2
Dim m2ec76a3e1eaacc363dd270302878acd8                        As SldWorks.Entity
Dim b8c7867c698f52057ae4e664309a7beb6                        As SldWorks.SelectionMgr
Dim wf0ca045f9be481927e0f67dfe17d1b77                       As SldWorks.SelectData
Dim b9bf81dc083d4d5b9586264d46f62ebe8                       As SldWorks.Entity
Dim b22cae1e6e02d414d303cc97e1fadb4b9                           As Variant
Dim bf49699a88461a2a540df64abb48b2f98                      As Long
Dim b95f845acaeb1bed280585702474b8385                    As Long
Dim b4e5c40e176840ea491c358d7b078c013                      As Boolean
Dim q11d1d2401fd0d54e6cff88ed4b502891                       As Variant
Dim n0299b426f4d4f9713bc8d4efb436f885                         As SldWorks.MathTransform
Dim bd788acd2ab98604acf912471d7b2c47d                               As Integer
Dim mb3d3a137c10d8cf36360b9c4ca4e1ff9                  As Object
Set r089f2e75bc0c2785f3502a8c94ce64fb = qcd7e5646aa31d4a3b54db53750164e31.OpenDoc6(model, 1, 0, "", bf49699a88461a2a540df64abb48b2f98, b95f845acaeb1bed280585702474b8385)
n88c5392f2cfe40f6063575c761ab6c7d = Mid(model, InStrRev(model, "\") + 1)
If r089f2e75bc0c2785f3502a8c94ce64fb Is Nothing Then
Exit Sub
End If
Set b67c3f8109b77bd4b38e5630fd32c514f = r089f2e75bc0c2785f3502a8c94ce64fb
Set b8c7867c698f52057ae4e664309a7beb6 = r089f2e75bc0c2785f3502a8c94ce64fb.SelectionManager
Set wf0ca045f9be481927e0f67dfe17d1b77 = b8c7867c698f52057ae4e664309a7beb6.CreateSelectData
Set wc8c39c5599ddf73063181785c0caae37 = r089f2e75bc0c2785f3502a8c94ce64fb
Set n5365312f3a16ea17d2b92ef20fa77d84 = wc8c39c5599ddf73063181785c0caae37.Extension
Set b75677af635f8cc2aab0c89a1ddfa26bb = n5365312f3a16ea17d2b92ef20fa77d84.CustomPropertyManager("")
If b75677af635f8cc2aab0c89a1ddfa26bb.Get("IS_HARDWARE") = "Yes" Or b75677af635f8cc2aab0c89a1ddfa26bb.Get("FLIP") = "Yes" Or InStr(model, "\Hardwares\") <> 0 Then
Debug.Print "Hardware component: " & n88c5392f2cfe40f6063575c761ab6c7d
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "Hardware component: " & n88c5392f2cfe40f6063575c761ab6c7d
Exit Sub
End If
b67c3f8109b77bd4b38e5630fd32c514f.ClearSelection2 True
bfd8c89eae9803e5f705c3dd395249328 = b67c3f8109b77bd4b38e5630fd32c514f.GetBodies2(swAllBodies, True)
Dim bf1492617213aa582c0775df45383c89a                      As Double
Dim bb323a9df80869e4e0b85a937f96c1909                As Integer
Dim b42c949cef725fec3e619f59152e39db5                       As Variant
bf1492617213aa582c0775df45383c89a = 0
For bd788acd2ab98604acf912471d7b2c47d = 0 To UBound(bfd8c89eae9803e5f705c3dd395249328)
Set w33548d5cda925bb321073eb295ef6a62 = bfd8c89eae9803e5f705c3dd395249328(bd788acd2ab98604acf912471d7b2c47d)
b42c949cef725fec3e619f59152e39db5 = w33548d5cda925bb321073eb295ef6a62.GetMassProperties(1)
If b42c949cef725fec3e619f59152e39db5(3) > bf1492617213aa582c0775df45383c89a Then
bf1492617213aa582c0775df45383c89a = b42c949cef725fec3e619f59152e39db5(3)
bb323a9df80869e4e0b85a937f96c1909 = bd788acd2ab98604acf912471d7b2c47d
End If
Next bd788acd2ab98604acf912471d7b2c47d
Set w33548d5cda925bb321073eb295ef6a62 = bfd8c89eae9803e5f705c3dd395249328(bb323a9df80869e4e0b85a937f96c1909)
Set b679764c23bf7d8a169575ebac35213ae = w33548d5cda925bb321073eb295ef6a62.Copy
Set m2ec76a3e1eaacc363dd270302878acd8 = b57cd87b65610e085b9ea334d181b4d3f(w33548d5cda925bb321073eb295ef6a62)
If m2ec76a3e1eaacc363dd270302878acd8 Is Nothing Then
Debug.Print n88c5392f2cfe40f6063575c761ab6c7d & ": Does not contain a flat face "
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine n88c5392f2cfe40f6063575c761ab6c7d & ": Does not contain a flat face "
Exit Sub
End If
Set b9bf81dc083d4d5b9586264d46f62ebe8 = m2ec76a3e1eaacc363dd270302878acd8.GetSafeEntity
b22cae1e6e02d414d303cc97e1fadb4b9 = m2ec76a3e1eaacc363dd270302878acd8.Normal
If Not (Round(Abs(b22cae1e6e02d414d303cc97e1fadb4b9(0)), 6) = 0 And Round(Abs(b22cae1e6e02d414d303cc97e1fadb4b9(1)), 6) = 0 And Round(Abs(b22cae1e6e02d414d303cc97e1fadb4b9(2)), 6) = 1) Then
For bd788acd2ab98604acf912471d7b2c47d = 0 To UBound(bfd8c89eae9803e5f705c3dd395249328)
b4e5c40e176840ea491c358d7b078c013 = b67c3f8109b77bd4b38e5630fd32c514f.Extension.SelectByID2(bfd8c89eae9803e5f705c3dd395249328(bd788acd2ab98604acf912471d7b2c47d).Name, "SOLIDBODY", 0, 0, 0, True, 1, Nothing, 0)
Next bd788acd2ab98604acf912471d7b2c47d
Set b432b23cd29982f1c659de7c15266a10b = b67c3f8109b77bd4b38e5630fd32c514f.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, 1)
Set b67927ea18feebbaa0a49705f4fec3eb2 = b432b23cd29982f1c659de7c15266a10b.GetDefinition
b4e5c40e176840ea491c358d7b078c013 = b67927ea18feebbaa0a49705f4fec3eb2.AccessSelections(r089f2e75bc0c2785f3502a8c94ce64fb, Nothing)
wf0ca045f9be481927e0f67dfe17d1b77.Mark = 1
b4e5c40e176840ea491c358d7b078c013 = b9bf81dc083d4d5b9586264d46f62ebe8.Select4(False, wf0ca045f9be481927e0f67dfe17d1b77)
Dim bfe2456d61ce9d2156dd13d4610d6c600 As String
bfe2456d61ce9d2156dd13d4610d6c600 = z1c324e9bb56f7d832776568176458a38(r089f2e75bc0c2785f3502a8c94ce64fb, 1)
b4e5c40e176840ea491c358d7b078c013 = b67c3f8109b77bd4b38e5630fd32c514f.Extension.SelectByID2(bfe2456d61ce9d2156dd13d4610d6c600, "PLANE", 0, 0, 0, True, 1, Nothing, 0)
If Not (b4e5c40e176840ea491c358d7b078c013) Then
b4e5c40e176840ea491c358d7b078c013 = b67c3f8109b77bd4b38e5630fd32c514f.Extension.SelectByID2("Front", "PLANE", 0, 0, 0, True, 1, Nothing, 0)
End If
If Not (b4e5c40e176840ea491c358d7b078c013) Then
Debug.Print ""
Debug.Print "####### Unable selection plane normal to Z in " & n88c5392f2cfe40f6063575c761ab6c7d & " plane name must be 'Front' or 'Front Plane' #######"
Debug.Print ""
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "####### Unable selection plane normal to Z in " & n88c5392f2cfe40f6063575c761ab6c7d & " plane name must be 'Front' or 'Front Plane' #######"
End If
b67927ea18feebbaa0a49705f4fec3eb2.AddMate Nothing, 0, 1, 0, 0, bf49699a88461a2a540df64abb48b2f98
b432b23cd29982f1c659de7c15266a10b.ModifyDefinition b67927ea18feebbaa0a49705f4fec3eb2, b67c3f8109b77bd4b38e5630fd32c514f, be5418de9d969c5ca5cb44308f32da188
b67927ea18feebbaa0a49705f4fec3eb2.ReleaseSelectionAccess
r089f2e75bc0c2785f3502a8c94ce64fb.ViewZoomtofit2
Else
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "Already aligned with Z: " & n88c5392f2cfe40f6063575c761ab6c7d
End If
Dim e03dae960a9b73193a014489fce4d2a23         As SldWorks.Edge
Dim b013924ecd7b07b8b6a200a382e6eecec             As SldWorks.Curve
Dim zca0cbf9d814a99e312766daf4dadfba1         As Variant
Dim b0e8529dea2db3219bcad294bc3058fa4          As Double
Dim b06835aff6a749bb70de9b38cda505595(2)       As Double
Set e03dae960a9b73193a014489fce4d2a23 = bb387760d7b3294fb83f59422f1f3c1bb(b9bf81dc083d4d5b9586264d46f62ebe8)
If Not e03dae960a9b73193a014489fce4d2a23 Is Nothing Then
Set b013924ecd7b07b8b6a200a382e6eecec = e03dae960a9b73193a014489fce4d2a23.GetCurve
zca0cbf9d814a99e312766daf4dadfba1 = e03dae960a9b73193a014489fce4d2a23.GetCurveParams2
b0e8529dea2db3219bcad294bc3058fa4 = b013924ecd7b07b8b6a200a382e6eecec.GetLength3(zca0cbf9d814a99e312766daf4dadfba1(6), zca0cbf9d814a99e312766daf4dadfba1(7))
Dim z57fbbe9a55b7e76e8772bb12c27d0537 As Integer
Set b8c7867c698f52057ae4e664309a7beb6 = r089f2e75bc0c2785f3502a8c94ce64fb.SelectionManager
Set wf0ca045f9be481927e0f67dfe17d1b77 = b8c7867c698f52057ae4e664309a7beb6.CreateSelectData
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To 2
b06835aff6a749bb70de9b38cda505595(z57fbbe9a55b7e76e8772bb12c27d0537) = (zca0cbf9d814a99e312766daf4dadfba1(z57fbbe9a55b7e76e8772bb12c27d0537 + 3) - zca0cbf9d814a99e312766daf4dadfba1(z57fbbe9a55b7e76e8772bb12c27d0537)) / b0e8529dea2db3219bcad294bc3058fa4
Next z57fbbe9a55b7e76e8772bb12c27d0537
Dim ea23a9964dbed1dd3581cd93684e8ca16          As Double
Dim wa8a3a2551ee0b7f8266853684c78e325         As Double
t8fcc0989bc0894d1dac35c7f5e7d32df(0) = 1#
t8fcc0989bc0894d1dac35c7f5e7d32df(1) = 0#
t8fcc0989bc0894d1dac35c7f5e7d32df(2) = 0#
bdcc132ec89e911e347bcd7c6e05d44ad = b8f9b81f33696150cc970d5b4ee7b1cf4(b06835aff6a749bb70de9b38cda505595, t8fcc0989bc0894d1dac35c7f5e7d32df)
ea23a9964dbed1dd3581cd93684e8ca16 = ebb15456c1bc4f43896e5a35698434455(b06835aff6a749bb70de9b38cda505595)
wa8a3a2551ee0b7f8266853684c78e325 = ebb15456c1bc4f43896e5a35698434455(t8fcc0989bc0894d1dac35c7f5e7d32df)
b9b4c91e56a7252144a2745fe17913d02 = (bdcc132ec89e911e347bcd7c6e05d44ad / (ea23a9964dbed1dd3581cd93684e8ca16 * wa8a3a2551ee0b7f8266853684c78e325))
b9b4c91e56a7252144a2745fe17913d02 = Arccos(b9b4c91e56a7252144a2745fe17913d02)
End If
If Not (Round(b9b4c91e56a7252144a2745fe17913d02 * 180 / pi, Precision) = 0 Or Round(b9b4c91e56a7252144a2745fe17913d02 * 180 / pi, Precision) = 180) Then
bfd8c89eae9803e5f705c3dd395249328 = b67c3f8109b77bd4b38e5630fd32c514f.GetBodies2(swAllBodies, True)
If b06835aff6a749bb70de9b38cda505595(1) < 0 Then
b9b4c91e56a7252144a2745fe17913d02 = -b9b4c91e56a7252144a2745fe17913d02
End If
For bd788acd2ab98604acf912471d7b2c47d = 0 To UBound(bfd8c89eae9803e5f705c3dd395249328)
b4e5c40e176840ea491c358d7b078c013 = b67c3f8109b77bd4b38e5630fd32c514f.Extension.SelectByID2(bfd8c89eae9803e5f705c3dd395249328(bd788acd2ab98604acf912471d7b2c47d).Name, "SOLIDBODY", 0, 0, 0, True, 1, Nothing, 0)
Next bd788acd2ab98604acf912471d7b2c47d
Set b432b23cd29982f1c659de7c15266a10b = b67c3f8109b77bd4b38e5630fd32c514f.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, -b9b4c91e56a7252144a2745fe17913d02, 0, 0, False, 1)
End If
If e03dae960a9b73193a014489fce4d2a23 Is Nothing Then
Debug.Print n88c5392f2cfe40f6063575c761ab6c7d & ": Does not contain a straight edge"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine n88c5392f2cfe40f6063575c761ab6c7d & ": Does not contain a straight edge"
Set bbeb837e9ee564e64f8607696461893a4 = n5365312f3a16ea17d2b92ef20fa77d84.CreateMassProperty
b26226ef0ed8b85195a2984e0f7903a1c = bbeb837e9ee564e64f8607696461893a4.PrincipleAxesOfInertia(0)
t8fcc0989bc0894d1dac35c7f5e7d32df(0) = 1
t8fcc0989bc0894d1dac35c7f5e7d32df(1) = 0
t8fcc0989bc0894d1dac35c7f5e7d32df(2) = 0
bdcc132ec89e911e347bcd7c6e05d44ad = b8f9b81f33696150cc970d5b4ee7b1cf4(b26226ef0ed8b85195a2984e0f7903a1c, t8fcc0989bc0894d1dac35c7f5e7d32df)
q8e5489469f46ff6977f70647e6d160e6 = ebb15456c1bc4f43896e5a35698434455(b26226ef0ed8b85195a2984e0f7903a1c)
qc459f7414f9a330324e431ffa95a66bb = ebb15456c1bc4f43896e5a35698434455(t8fcc0989bc0894d1dac35c7f5e7d32df)
b9b4c91e56a7252144a2745fe17913d02 = (bdcc132ec89e911e347bcd7c6e05d44ad / (q8e5489469f46ff6977f70647e6d160e6 * qc459f7414f9a330324e431ffa95a66bb))
b9b4c91e56a7252144a2745fe17913d02 = Arccos(b9b4c91e56a7252144a2745fe17913d02)
If Round(b9b4c91e56a7252144a2745fe17913d02 * 180 / pi, 2) <> 0 Then
bfd8c89eae9803e5f705c3dd395249328 = b67c3f8109b77bd4b38e5630fd32c514f.GetBodies2(swAllBodies, True)
If b26226ef0ed8b85195a2984e0f7903a1c(1) > 0 Then
b9b4c91e56a7252144a2745fe17913d02 = -b9b4c91e56a7252144a2745fe17913d02
End If
For bd788acd2ab98604acf912471d7b2c47d = 0 To UBound(bfd8c89eae9803e5f705c3dd395249328)
b4e5c40e176840ea491c358d7b078c013 = b67c3f8109b77bd4b38e5630fd32c514f.Extension.SelectByID2(bfd8c89eae9803e5f705c3dd395249328(bd788acd2ab98604acf912471d7b2c47d).Name, "SOLIDBODY", 0, 0, 0, True, 1, Nothing, 0)
Next bd788acd2ab98604acf912471d7b2c47d
Set b432b23cd29982f1c659de7c15266a10b = b67c3f8109b77bd4b38e5630fd32c514f.FeatureManager.InsertMoveCopyBody2(0, 0, 0, 0, 0, 0, 0, b9b4c91e56a7252144a2745fe17913d02, 0, 0, False, 1)
End If
End If
Set b8cdfda0b5a2e71faa28ead12f563eb75 = wc8c39c5599ddf73063181785c0caae37.FeatureManager
b4e5c40e176840ea491c358d7b078c013 = b8cdfda0b5a2e71faa28ead12f563eb75.EditRollback(swMoveRollbackBarToPreviousPosition, "")
b4e5c40e176840ea491c358d7b078c013 = b8cdfda0b5a2e71faa28ead12f563eb75.EditRollback(swMoveRollbackBarToEnd, "")
bfd8c89eae9803e5f705c3dd395249328 = b67c3f8109b77bd4b38e5630fd32c514f.GetBodies2(swAllBodies, True)
Set r089f2e75bc0c2785f3502a8c94ce64fb = b67c3f8109b77bd4b38e5630fd32c514f
bf1492617213aa582c0775df45383c89a = 0
For bd788acd2ab98604acf912471d7b2c47d = 0 To UBound(bfd8c89eae9803e5f705c3dd395249328)
Set w33548d5cda925bb321073eb295ef6a62 = bfd8c89eae9803e5f705c3dd395249328(bd788acd2ab98604acf912471d7b2c47d)
b42c949cef725fec3e619f59152e39db5 = w33548d5cda925bb321073eb295ef6a62.GetMassProperties(1)
If b42c949cef725fec3e619f59152e39db5(3) > bf1492617213aa582c0775df45383c89a Then
bf1492617213aa582c0775df45383c89a = b42c949cef725fec3e619f59152e39db5(3)
bb323a9df80869e4e0b85a937f96c1909 = bd788acd2ab98604acf912471d7b2c47d
End If
Next bd788acd2ab98604acf912471d7b2c47d
Set w33548d5cda925bb321073eb295ef6a62 = bfd8c89eae9803e5f705c3dd395249328(bb323a9df80869e4e0b85a937f96c1909)
Set n0299b426f4d4f9713bc8d4efb436f885 = bc18d56f2833a99b219a49c707e4c0cac(w33548d5cda925bb321073eb295ef6a62, b679764c23bf7d8a169575ebac35213ae)
If Not n0299b426f4d4f9713bc8d4efb436f885 Is Nothing Then
Set n0299b426f4d4f9713bc8d4efb436f885 = n0299b426f4d4f9713bc8d4efb436f885.Inverse
Dim t30f0178cf0062da0c85c859c72620266           As SldWorks.Component2
Dim ecb34e868e1dc604578c704715874ebf5             As SldWorks.Component2
Dim b606e01e850f828bbbba8bb133f59cae2         As SldWorks.MathTransform
Dim baee2eeded58610e6e39630dabbd9128a            As Boolean
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(t0f0bc4817eb80edbd3b60e38d6e44c34)
Set t30f0178cf0062da0c85c859c72620266 = t0f0bc4817eb80edbd3b60e38d6e44c34(z57fbbe9a55b7e76e8772bb12c27d0537)
If model = t30f0178cf0062da0c85c859c72620266.GetPathName Then
baee2eeded58610e6e39630dabbd9128a = False
If t30f0178cf0062da0c85c859c72620266.IsFixed Then
b4e5c40e176840ea491c358d7b078c013 = w7a695257e7af03642eb7fff3efb75bb3.SelectByID2(t0f0bc4817eb80edbd3b60e38d6e44c34(z57fbbe9a55b7e76e8772bb12c27d0537).Name & "@" & q02ed784cf0fc1dc30d0c7933763a602c, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
z4f76b89c98a78e0601bd278fdabfe3ab.UnfixComponent
baee2eeded58610e6e39630dabbd9128a = True
End If
Set ecb34e868e1dc604578c704715874ebf5 = z4f76b89c98a78e0601bd278fdabfe3ab.GetComponentByName(t0f0bc4817eb80edbd3b60e38d6e44c34(z57fbbe9a55b7e76e8772bb12c27d0537).Name)
If Not ecb34e868e1dc604578c704715874ebf5 Is Nothing Then
Set b606e01e850f828bbbba8bb133f59cae2 = ecb34e868e1dc604578c704715874ebf5.Transform2
ecb34e868e1dc604578c704715874ebf5.Transform2 = n0299b426f4d4f9713bc8d4efb436f885.Multiply(b606e01e850f828bbbba8bb133f59cae2)
Else
Debug.Print "TRANSFORM SKIPPED for: " & t0f0bc4817eb80edbd3b60e38d6e44c34(z57fbbe9a55b7e76e8772bb12c27d0537).Name
End If
If baee2eeded58610e6e39630dabbd9128a Then
b4e5c40e176840ea491c358d7b078c013 = w7a695257e7af03642eb7fff3efb75bb3.SelectByID2(t0f0bc4817eb80edbd3b60e38d6e44c34(z57fbbe9a55b7e76e8772bb12c27d0537).Name & "@Adam", "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
z4f76b89c98a78e0601bd278fdabfe3ab.FixComponent
End If
End If
Next z57fbbe9a55b7e76e8772bb12c27d0537
End If
q11d1d2401fd0d54e6cff88ed4b502891 = b67c3f8109b77bd4b38e5630fd32c514f.GetPartBox(True)
Dim b81b400e697d89d3663cb93382ab0dda5      As Double
Dim Width       As Double
Dim bd62504a4f63646d093adb0d6e30b4a48   As Double
b81b400e697d89d3663cb93382ab0dda5 = Round(Str(q11d1d2401fd0d54e6cff88ed4b502891(3) * 1000#) - Str(q11d1d2401fd0d54e6cff88ed4b502891(0) * 1000#), 2)
Width = Round(Str(q11d1d2401fd0d54e6cff88ed4b502891(4) * 1000#) - Str(q11d1d2401fd0d54e6cff88ed4b502891(1) * 1000#), 2)
bd62504a4f63646d093adb0d6e30b4a48 = Round(Str(q11d1d2401fd0d54e6cff88ed4b502891(5) * 1000#) - Str(q11d1d2401fd0d54e6cff88ed4b502891(2) * 1000#), 2)
Set b75677af635f8cc2aab0c89a1ddfa26bb = n5365312f3a16ea17d2b92ef20fa77d84.CustomPropertyManager("")
b4e5c40e176840ea491c358d7b078c013 = b75677af635f8cc2aab0c89a1ddfa26bb.Add3("Length", swCustomInfoText, b81b400e697d89d3663cb93382ab0dda5, swCustomPropertyReplaceValue)
b4e5c40e176840ea491c358d7b078c013 = b75677af635f8cc2aab0c89a1ddfa26bb.Add3("Width", swCustomInfoText, Width, swCustomPropertyReplaceValue)
b4e5c40e176840ea491c358d7b078c013 = b75677af635f8cc2aab0c89a1ddfa26bb.Add3("Thickness", swCustomInfoText, bd62504a4f63646d093adb0d6e30b4a48, swCustomPropertyReplaceValue)
r089f2e75bc0c2785f3502a8c94ce64fb.SetSaveFlag
End Sub
Public Function b57cd87b65610e085b9ea334d181b4d3f(w33548d5cda925bb321073eb295ef6a62 As SldWorks.Body2) As SldWorks.Entity
Dim wf0ca045f9be481927e0f67dfe17d1b77           As SldWorks.SelectData
Dim rffc81c76fd2991e2439e5b991443b4b6              As SldWorks.Face
Dim b9d7565a43d54c4febd1406ce1d3da1b6       As SldWorks.Face
Dim b001a9eacbc89185fbe28bf716296ff0d            As Double
Dim bdd3ea8c4f00f380eb0b97065e334c42c        As Double
Dim be5e7faa24660dc04fcc1fd18080be146              As SldWorks.Surface
Set rffc81c76fd2991e2439e5b991443b4b6 = w33548d5cda925bb321073eb295ef6a62.GetFirstFace
r089f2e75bc0c2785f3502a8c94ce64fb.ClearSelection2 True
Do While Not rffc81c76fd2991e2439e5b991443b4b6 Is Nothing
b001a9eacbc89185fbe28bf716296ff0d = rffc81c76fd2991e2439e5b991443b4b6.GetArea
Set be5e7faa24660dc04fcc1fd18080be146 = rffc81c76fd2991e2439e5b991443b4b6.GetSurface
If be5e7faa24660dc04fcc1fd18080be146.Identity = 4001 Then
If b001a9eacbc89185fbe28bf716296ff0d > bdd3ea8c4f00f380eb0b97065e334c42c Then
Set b9d7565a43d54c4febd1406ce1d3da1b6 = rffc81c76fd2991e2439e5b991443b4b6
bdd3ea8c4f00f380eb0b97065e334c42c = b001a9eacbc89185fbe28bf716296ff0d
End If
End If
Set rffc81c76fd2991e2439e5b991443b4b6 = rffc81c76fd2991e2439e5b991443b4b6.GetNextFace
Loop
Set b57cd87b65610e085b9ea334d181b4d3f = b9d7565a43d54c4febd1406ce1d3da1b6
End Function
Function b8f9b81f33696150cc970d5b4ee7b1cf4(Array1 As Variant, Array2 As Variant) As Double
Dim z57fbbe9a55b7e76e8772bb12c27d0537 As Integer
Dim z6c3eff9e569fb98b27c3c6cb92f9c224 As Double
z6c3eff9e569fb98b27c3c6cb92f9c224 = 0
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(Array1)
z6c3eff9e569fb98b27c3c6cb92f9c224 = z6c3eff9e569fb98b27c3c6cb92f9c224 + Array1(z57fbbe9a55b7e76e8772bb12c27d0537) * Array2(z57fbbe9a55b7e76e8772bb12c27d0537)
Next z57fbbe9a55b7e76e8772bb12c27d0537
b8f9b81f33696150cc970d5b4ee7b1cf4 = z6c3eff9e569fb98b27c3c6cb92f9c224
End Function
Function ebb15456c1bc4f43896e5a35698434455(Array1 As Variant) As Double
Dim z57fbbe9a55b7e76e8772bb12c27d0537 As Integer
Dim z6c3eff9e569fb98b27c3c6cb92f9c224 As Double
z6c3eff9e569fb98b27c3c6cb92f9c224 = 0
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(Array1)
z6c3eff9e569fb98b27c3c6cb92f9c224 = z6c3eff9e569fb98b27c3c6cb92f9c224 + Array1(z57fbbe9a55b7e76e8772bb12c27d0537) ^ 2
Next z57fbbe9a55b7e76e8772bb12c27d0537
ebb15456c1bc4f43896e5a35698434455 = Sqr(z6c3eff9e569fb98b27c3c6cb92f9c224)
End Function
Function bb387760d7b3294fb83f59422f1f3c1bb(Face As Variant) As Variant
Dim bed14e7e027efcd0a307fd5dac3aabb38                  As Variant
Dim b013924ecd7b07b8b6a200a382e6eecec                 As SldWorks.Curve
Dim nd681de2593e14d2b4ad6d070a0a4cb67          As SldWorks.Curve
Dim ne071642f162122658e020324b3a4e3b6            As SldWorks.CurveParamData
Dim z57fbbe9a55b7e76e8772bb12c27d0537                       As Integer
Dim m2506ab39ff9e92c98563e1a3904df1bd                       As Long
Dim zca0cbf9d814a99e312766daf4dadfba1             As Variant
Dim t30981585944aebd2983a66ffb909fa33      As Variant
Dim b0e8529dea2db3219bcad294bc3058fa4              As Double
Dim z16cbd29c8ce70fb17b7666c215ead0dd          As Double
Dim bcadca90c0e65359ec167b5314f1e851d   As Double
Dim e03dae960a9b73193a014489fce4d2a23             As SldWorks.Edge
Dim qb44f2644340d6c4a48d345a2f830d356            As SldWorks.Edge
Dim b06835aff6a749bb70de9b38cda505595(2)           As Double
Dim n9440df21957b9ea404b36ffaf6532f35                 As Boolean
Dim b948a74a5c1b6ef4528c70d292c25d974                 As Boolean
Dim r14a94b7676b0720c1a13720e82e558b9               As Variant
Dim r0315c175c208c4b35efdd858b3573297             As Variant
bed14e7e027efcd0a307fd5dac3aabb38 = Face.GetEdges
For m2506ab39ff9e92c98563e1a3904df1bd = 0 To UBound(bed14e7e027efcd0a307fd5dac3aabb38)
Set b013924ecd7b07b8b6a200a382e6eecec = bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd).GetCurve
If b013924ecd7b07b8b6a200a382e6eecec.Identity = 3001 Then
If m2506ab39ff9e92c98563e1a3904df1bd = 0 Then
n9440df21957b9ea404b36ffaf6532f35 = bc4c7c2f522f75d166d7b156003fcad0e(bed14e7e027efcd0a307fd5dac3aabb38(0), bed14e7e027efcd0a307fd5dac3aabb38(UBound(bed14e7e027efcd0a307fd5dac3aabb38)))
b948a74a5c1b6ef4528c70d292c25d974 = bc4c7c2f522f75d166d7b156003fcad0e(bed14e7e027efcd0a307fd5dac3aabb38(0), bed14e7e027efcd0a307fd5dac3aabb38(1))
ElseIf m2506ab39ff9e92c98563e1a3904df1bd = UBound(bed14e7e027efcd0a307fd5dac3aabb38) Then
n9440df21957b9ea404b36ffaf6532f35 = bc4c7c2f522f75d166d7b156003fcad0e(bed14e7e027efcd0a307fd5dac3aabb38(UBound(bed14e7e027efcd0a307fd5dac3aabb38)), bed14e7e027efcd0a307fd5dac3aabb38(0))
b948a74a5c1b6ef4528c70d292c25d974 = bc4c7c2f522f75d166d7b156003fcad0e(bed14e7e027efcd0a307fd5dac3aabb38(UBound(bed14e7e027efcd0a307fd5dac3aabb38)), bed14e7e027efcd0a307fd5dac3aabb38(UBound(bed14e7e027efcd0a307fd5dac3aabb38) - 1))
Else
n9440df21957b9ea404b36ffaf6532f35 = bc4c7c2f522f75d166d7b156003fcad0e(bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd), bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd + 1))
b948a74a5c1b6ef4528c70d292c25d974 = bc4c7c2f522f75d166d7b156003fcad0e(bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd), bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd - 1))
End If
Set b013924ecd7b07b8b6a200a382e6eecec = bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd).GetCurve
zca0cbf9d814a99e312766daf4dadfba1 = bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd).GetCurveParams2
b0e8529dea2db3219bcad294bc3058fa4 = b013924ecd7b07b8b6a200a382e6eecec.GetLength3(zca0cbf9d814a99e312766daf4dadfba1(6), zca0cbf9d814a99e312766daf4dadfba1(7))
If n9440df21957b9ea404b36ffaf6532f35 Or b948a74a5c1b6ef4528c70d292c25d974 Then
If b0e8529dea2db3219bcad294bc3058fa4 > z16cbd29c8ce70fb17b7666c215ead0dd Then
z16cbd29c8ce70fb17b7666c215ead0dd = b0e8529dea2db3219bcad294bc3058fa4
Set e03dae960a9b73193a014489fce4d2a23 = bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd)
End If
Else
If b0e8529dea2db3219bcad294bc3058fa4 > bcadca90c0e65359ec167b5314f1e851d Then
bcadca90c0e65359ec167b5314f1e851d = b0e8529dea2db3219bcad294bc3058fa4
Set qb44f2644340d6c4a48d345a2f830d356 = bed14e7e027efcd0a307fd5dac3aabb38(m2506ab39ff9e92c98563e1a3904df1bd)
End If
End If
End If
Next m2506ab39ff9e92c98563e1a3904df1bd
If Not e03dae960a9b73193a014489fce4d2a23 Is Nothing Then
Set b013924ecd7b07b8b6a200a382e6eecec = e03dae960a9b73193a014489fce4d2a23.GetCurve
zca0cbf9d814a99e312766daf4dadfba1 = e03dae960a9b73193a014489fce4d2a23.GetCurveParams2
b0e8529dea2db3219bcad294bc3058fa4 = z16cbd29c8ce70fb17b7666c215ead0dd
Set bb387760d7b3294fb83f59422f1f3c1bb = e03dae960a9b73193a014489fce4d2a23
ElseIf Not qb44f2644340d6c4a48d345a2f830d356 Is Nothing Then
Set b013924ecd7b07b8b6a200a382e6eecec = qb44f2644340d6c4a48d345a2f830d356.GetCurve
zca0cbf9d814a99e312766daf4dadfba1 = qb44f2644340d6c4a48d345a2f830d356.GetCurveParams2
b0e8529dea2db3219bcad294bc3058fa4 = bcadca90c0e65359ec167b5314f1e851d
Set bb387760d7b3294fb83f59422f1f3c1bb = qb44f2644340d6c4a48d345a2f830d356
Else
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "Could not find straight edge in: " & n88c5392f2cfe40f6063575c761ab6c7d
Set bb387760d7b3294fb83f59422f1f3c1bb = Nothing
Exit Function
End If
End Function
Function bc4c7c2f522f75d166d7b156003fcad0e(Edge1 As Variant, Edge2 As Variant) As Boolean
Dim b013924ecd7b07b8b6a200a382e6eecec                 As SldWorks.Curve
Dim ba98c02c96961399ef1cd0e16a240b7c8             As Variant
Dim bf7fa53e62fae26fd56ac33fa7316f878             As Variant
Dim z57fbbe9a55b7e76e8772bb12c27d0537                       As Integer
Dim Vector1(2)              As Double
Dim Vector2(2)              As Double
Dim b77b80bb9a97e531c0e4d7afe4cb1fcaf           As Double
Dim b5bb4a8f96b73347aa569c8bd8b359903           As Double
Dim b9b4c91e56a7252144a2745fe17913d02                   As Double
Dim bdcc132ec89e911e347bcd7c6e05d44ad                 As Double
Set b013924ecd7b07b8b6a200a382e6eecec = Edge2.GetCurve
If Not b013924ecd7b07b8b6a200a382e6eecec.Identity = 3001 Then
bc4c7c2f522f75d166d7b156003fcad0e = False
Exit Function
End If
ba98c02c96961399ef1cd0e16a240b7c8 = Edge1.GetCurveParams2
bf7fa53e62fae26fd56ac33fa7316f878 = Edge2.GetCurveParams2
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To 2
Vector1(z57fbbe9a55b7e76e8772bb12c27d0537) = ba98c02c96961399ef1cd0e16a240b7c8(z57fbbe9a55b7e76e8772bb12c27d0537 + 3) - ba98c02c96961399ef1cd0e16a240b7c8(z57fbbe9a55b7e76e8772bb12c27d0537)
Vector2(z57fbbe9a55b7e76e8772bb12c27d0537) = -bf7fa53e62fae26fd56ac33fa7316f878(z57fbbe9a55b7e76e8772bb12c27d0537 + 3) + bf7fa53e62fae26fd56ac33fa7316f878(z57fbbe9a55b7e76e8772bb12c27d0537)
Next z57fbbe9a55b7e76e8772bb12c27d0537
bdcc132ec89e911e347bcd7c6e05d44ad = b8f9b81f33696150cc970d5b4ee7b1cf4(Vector1, Vector2)
b77b80bb9a97e531c0e4d7afe4cb1fcaf = ebb15456c1bc4f43896e5a35698434455(Vector1)
b5bb4a8f96b73347aa569c8bd8b359903 = ebb15456c1bc4f43896e5a35698434455(Vector2)
b9b4c91e56a7252144a2745fe17913d02 = (bdcc132ec89e911e347bcd7c6e05d44ad / (b77b80bb9a97e531c0e4d7afe4cb1fcaf * b5bb4a8f96b73347aa569c8bd8b359903))
b9b4c91e56a7252144a2745fe17913d02 = Arccos(b9b4c91e56a7252144a2745fe17913d02)
b9b4c91e56a7252144a2745fe17913d02 = b9b4c91e56a7252144a2745fe17913d02 * 180 / pi
b9b4c91e56a7252144a2745fe17913d02 = Round(b9b4c91e56a7252144a2745fe17913d02, Precision)
If b9b4c91e56a7252144a2745fe17913d02 = 90 Then
bc4c7c2f522f75d166d7b156003fcad0e = True
Else
bc4c7c2f522f75d166d7b156003fcad0e = False
End If
End Function
Function m1e98ad36788ac2ca18eb096d3a858f8e(Assembly As SldWorks.AssemblyDoc) As Variant
Dim rdb9aa470bf04fe7ac33cf99555c952a3              As Object
Dim bf00650a4e09a527cf4d146fe07aebc73                  As Variant
Dim bf79075ea2e17d0bf2eb4ec624b45fd6f                  As SldWorks.Component2
Dim z57fbbe9a55b7e76e8772bb12c27d0537                       As Integer
Dim e4d103366b2c0dcd16b291e93d75211ad                    As String
Dim b0a6494a16a644c0bc61f2d010cd2d3b2                    As String
Set rdb9aa470bf04fe7ac33cf99555c952a3 = CreateObject("Scripting.Dictionary")
bf00650a4e09a527cf4d146fe07aebc73 = Assembly.GetComponents(False)
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(bf00650a4e09a527cf4d146fe07aebc73)
Set bf79075ea2e17d0bf2eb4ec624b45fd6f = bf00650a4e09a527cf4d146fe07aebc73(z57fbbe9a55b7e76e8772bb12c27d0537)
If bf79075ea2e17d0bf2eb4ec624b45fd6f.IsSuppressed = False Then
e4d103366b2c0dcd16b291e93d75211ad = bf79075ea2e17d0bf2eb4ec624b45fd6f.GetPathName
If rdb9aa470bf04fe7ac33cf99555c952a3.exists(e4d103366b2c0dcd16b291e93d75211ad) Then
rdb9aa470bf04fe7ac33cf99555c952a3.Item(e4d103366b2c0dcd16b291e93d75211ad) = rdb9aa470bf04fe7ac33cf99555c952a3.Item(e4d103366b2c0dcd16b291e93d75211ad) + 1
Else
rdb9aa470bf04fe7ac33cf99555c952a3.Add e4d103366b2c0dcd16b291e93d75211ad, 1
End If
End If
Next
Dim wae800353f68b1d379a5e185c0932b738 As Variant
wae800353f68b1d379a5e185c0932b738 = rdb9aa470bf04fe7ac33cf99555c952a3.Keys
m1e98ad36788ac2ca18eb096d3a858f8e = wae800353f68b1d379a5e185c0932b738
End Function
Public Function q8b61244ae4822e03b47ad99c4fc71e0c(Maxiter As Integer, ByVal Caption As String)
Dim t382b9d8066debe4514dc179247019dcf     As Double
Dim r2e00027eabe0b722f19a23ac90779560            As Double
Dim qddcbd44bfca946d6461895d06978f141  As Double
Dim a                   As Integer
If Caption = "Rebuilding..." Then
qea414b44a6ee13e50a6ec2eb47ffefa7 = qea414b44a6ee13e50a6ec2eb47ffefa7 - 1
ProgressBar.Bar.BackColor = &HC000&
End If
qea414b44a6ee13e50a6ec2eb47ffefa7 = qea414b44a6ee13e50a6ec2eb47ffefa7 + 1
Maxiter = Maxiter + 1
t382b9d8066debe4514dc179247019dcf = qea414b44a6ee13e50a6ec2eb47ffefa7 / Maxiter
r2e00027eabe0b722f19a23ac90779560 = ProgressBar.Frame.Width * t382b9d8066debe4514dc179247019dcf
qddcbd44bfca946d6461895d06978f141 = Round(t382b9d8066debe4514dc179247019dcf * 100, 0)
ProgressBar.Bar.Width = r2e00027eabe0b722f19a23ac90779560 - 0.015 * r2e00027eabe0b722f19a23ac90779560
ProgressBar.Text2.Caption = qddcbd44bfca946d6461895d06978f141 & "% Complete"
ProgressBar.Text.Caption = qea414b44a6ee13e50a6ec2eb47ffefa7 & " of " & Maxiter
ProgressBar.Text3.Caption = Caption
DoEvents
If (qddcbd44bfca946d6461895d06978f141 / 10) Mod 2 = 0 Then
ProgressBar.Image1.Visible = False
ProgressBar.Image2.Visible = True
Else
ProgressBar.Image1.Visible = True
ProgressBar.Image2.Visible = False
End If
End Function
Function bc18d56f2833a99b219a49c707e4c0cac(swThisBody As SldWorks.Body2, swOtherBody As SldWorks.Body2) As SldWorks.MathTransform
Dim b8181efabee10d40dd9053ee11954b909      As Object
Dim nfae4202f6209b448fb8de3d7d4daa3c7         As SldWorks.MathTransform
Dim r81643b7f624a34b70d8ee428e5759613                 As Variant
Set b8181efabee10d40dd9053ee11954b909 = CreateObject("Scripting.Dictionary")
If swThisBody.GetCoincidenceTransform2(swOtherBody, nfae4202f6209b448fb8de3d7d4daa3c7) Then
If Not nfae4202f6209b448fb8de3d7d4daa3c7 Is Nothing Then
Dim t977aa682f63b1d4828b3c033c1b9a015 As Variant
t977aa682f63b1d4828b3c033c1b9a015 = nfae4202f6209b448fb8de3d7d4daa3c7.ArrayData
Dim r6baac1704054683f74efbedda182bf92 As Boolean
r6baac1704054683f74efbedda182bf92 = False
For Each r81643b7f624a34b70d8ee428e5759613 In b8181efabee10d40dd9053ee11954b909.Keys
If Not r81643b7f624a34b70d8ee428e5759613 Is Nothing Then
Dim tde97cb863eeb1602f1c54410c9525f0c As SldWorks.MathTransform
Set tde97cb863eeb1602f1c54410c9525f0c = r81643b7f624a34b70d8ee428e5759613
If b6b1295bcdd5794e47830fc4d56daf133(nfae4202f6209b448fb8de3d7d4daa3c7, tde97cb863eeb1602f1c54410c9525f0c) Then
b8181efabee10d40dd9053ee11954b909(tde97cb863eeb1602f1c54410c9525f0c) = b8181efabee10d40dd9053ee11954b909(tde97cb863eeb1602f1c54410c9525f0c) + 1
r6baac1704054683f74efbedda182bf92 = True
Exit For
End If
End If
Next
If Not r6baac1704054683f74efbedda182bf92 Then
b8181efabee10d40dd9053ee11954b909.Add nfae4202f6209b448fb8de3d7d4daa3c7, 1
End If
End If
Else
Debug.Print "CANNOT COINCIDE " & n88c5392f2cfe40f6063575c761ab6c7d
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "Could not reset location of: " & n88c5392f2cfe40f6063575c761ab6c7d
End If
Dim e001a7f66ef9b10ae3f319c0db7c623f8 As Integer
e001a7f66ef9b10ae3f319c0db7c623f8 = 0
For Each r81643b7f624a34b70d8ee428e5759613 In b8181efabee10d40dd9053ee11954b909.Keys
If Not r81643b7f624a34b70d8ee428e5759613 Is Nothing Then
Dim b5f68dfb2fb4e3332a536ce1a3a5cd3b1 As SldWorks.MathTransform
Set b5f68dfb2fb4e3332a536ce1a3a5cd3b1 = r81643b7f624a34b70d8ee428e5759613
If b8181efabee10d40dd9053ee11954b909(b5f68dfb2fb4e3332a536ce1a3a5cd3b1) > e001a7f66ef9b10ae3f319c0db7c623f8 Then
e001a7f66ef9b10ae3f319c0db7c623f8 = b8181efabee10d40dd9053ee11954b909(b5f68dfb2fb4e3332a536ce1a3a5cd3b1)
Set bc18d56f2833a99b219a49c707e4c0cac = b5f68dfb2fb4e3332a536ce1a3a5cd3b1
End If
End If
Next
End Function
Function b6b1295bcdd5794e47830fc4d56daf133(firstTransform As SldWorks.MathTransform, secondTransform As SldWorks.MathTransform) As Boolean
Dim td1044d8ed9c39e6bdfd376d537651509 As Variant
td1044d8ed9c39e6bdfd376d537651509 = firstTransform.ArrayData
Dim b60513bac05634fae746323acb993cad8 As Variant
b60513bac05634fae746323acb993cad8 = secondTransform.ArrayData
Dim z57fbbe9a55b7e76e8772bb12c27d0537 As Integer
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(td1044d8ed9c39e6bdfd376d537651509)
If Not t02e74b8aa1a7e04ae8e7998dfeff1933(CDbl(td1044d8ed9c39e6bdfd376d537651509(z57fbbe9a55b7e76e8772bb12c27d0537)), CDbl(b60513bac05634fae746323acb993cad8(z57fbbe9a55b7e76e8772bb12c27d0537))) Then
b6b1295bcdd5794e47830fc4d56daf133 = False
Exit Function
End If
Next
b6b1295bcdd5794e47830fc4d56daf133 = True
End Function
Function t02e74b8aa1a7e04ae8e7998dfeff1933(firstValue As Double, secondValue As Double, Optional tol As Double = 0.00000001) As Boolean
t02e74b8aa1a7e04ae8e7998dfeff1933 = Abs(secondValue - firstValue) <= tol
End Function
Function b5ebaf9eab0956be7d743ef1391203e5e(File As String) As String
Dim b93ff3abd396e4ca22447f6d4a50ae604 As String
b93ff3abd396e4ca22447f6d4a50ae604 = Right(File, Len(File) - InStrRev(File, "\"))
If Right(b93ff3abd396e4ca22447f6d4a50ae604, 7) = ".sldasm" Or Right(b93ff3abd396e4ca22447f6d4a50ae604, 7) = ".sldprt" Then
b5ebaf9eab0956be7d743ef1391203e5e = Left(b93ff3abd396e4ca22447f6d4a50ae604, Len(b93ff3abd396e4ca22447f6d4a50ae604) - 7)
Else
b5ebaf9eab0956be7d743ef1391203e5e = b93ff3abd396e4ca22447f6d4a50ae604
End If
End Function
Function b695664185dfefc4c1ad40b10646963d0(Name As String) As String
b695664185dfefc4c1ad40b10646963d0 = Replace(Name, "/", "@")
End Function
Function z1c324e9bb56f7d832776568176458a38(model As SldWorks.ModelDoc2, planeType As Integer) As String
Dim b5398c809e609a062040c7b919339789c As Integer
Dim b1596bdb5b54bd50c7bd07ea6d9da6316 As SldWorks.Feature
Set b1596bdb5b54bd50c7bd07ea6d9da6316 = model.FirstFeature
Do While Not b1596bdb5b54bd50c7bd07ea6d9da6316 Is Nothing
If b1596bdb5b54bd50c7bd07ea6d9da6316.GetTypeName = "RefPlane" Then
Debug.Print b1596bdb5b54bd50c7bd07ea6d9da6316.Description
b5398c809e609a062040c7b919339789c = b5398c809e609a062040c7b919339789c + 1
If CInt(planeType) = b5398c809e609a062040c7b919339789c Then
z1c324e9bb56f7d832776568176458a38 = b1596bdb5b54bd50c7bd07ea6d9da6316.Description
Exit Function
End If
End If
Set b1596bdb5b54bd50c7bd07ea6d9da6316 = b1596bdb5b54bd50c7bd07ea6d9da6316.GetNextFeature
Loop
End Function
