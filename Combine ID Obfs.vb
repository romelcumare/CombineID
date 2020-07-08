'VBA code protection using: www.excel-pratique.com/en/vba_tricks/vba-obfuscator
Option Explicit
Option bcb8c6c165563621a142e2b497b3c8f14 Text
Dim b11208ee7b1b6ffdc4d54001bf42aeb1a                   As SldWorks.SldWorks
Dim r089f2e75bc0c2785f3502a8c94ce64fb                 As SldWorks.ModelDoc2
Dim b6e7808d6dcb172b1a01afa6eaecae213                 As SldWorks.PartDoc
Dim nfeef4826d4eef9f613bf47c1a398604b                 As SldWorks.PartDoc
Dim z4f76b89c98a78e0601bd278fdabfe3ab                  As SldWorks.AssemblyDoc
Dim bf79075ea2e17d0bf2eb4ec624b45fd6f                  As SldWorks.Component2
Dim bf49699a88461a2a540df64abb48b2f98              As Long
Dim b4e5c40e176840ea491c358d7b078c013              As Boolean
Dim b94b7ebfc3d693c01e015c80ebfdd0725            As String
Dim q02ed784cf0fc1dc30d0c7933763a602c            As String
Dim z56989495ec6c0363b02cdbc53cb1808d                 As String
Dim b383222a9bea4084b4a55b3485db67e60                 As String
Dim rab2bd3733535f62d8ead8ad301e13be2                     As FileSystemObject
Dim bfc5d43b8304fa915bdf038e967c4d02a              As TextStream
Dim e5311fd6755b95d262631b01ddeb6773f              As String
Dim z31d5ee96103ab4d839c6808b87667355              As String
Dim b6dc70dfad7c15a723643c241dff64c41               As Integer
Dim r910c0775bffa167dbee861a2869f037e                As Variant
Dim w17ae4354e08c7a9a455c3a52c7620587             As Boolean
Dim r2e26fd7a88815e76612796762f046a59                As String
Dim be23c27a34feae679476e4e9ba4fd2352           As Integer
Dim bea5244a27ad17c66238c31e141ee170d           As String
Dim b253afc06d74e2ba0ae1e171327d89783            As String
Dim t382b9d8066debe4514dc179247019dcf         As Double
Dim r2e00027eabe0b722f19a23ac90779560                As Long
Dim qddcbd44bfca946d6461895d06978f141      As Double
Dim qea414b44a6ee13e50a6ec2eb47ffefa7                As Integer
Public CancelButton         As Boolean
Dim w5a54905c48b2789bf04dc498478c4272          As Boolean
Dim b00bb11a073ecc16d662c97410cee82de            As Boolean
Dim qfa9a074829425d28be1af4bebc0d480e             As Boolean
Dim rd70fd074b9edb535344cb5d4dfc857e5        As Boolean
Dim m77fc1e371470054ae753603ff082b6de               As Variant
Sub Main()
be23c27a34feae679476e4e9ba4fd2352 = 0
b253afc06d74e2ba0ae1e171327d89783 = ""
bea5244a27ad17c66238c31e141ee170d = ""
m77fc1e371470054ae753603ff082b6de = Array(1, 0, 0, 0, 1, 0, 0, 0, 1)
Debug.Print ""
Debug.Print "----------- Macro Started -----------"
Dim z57fbbe9a55b7e76e8772bb12c27d0537                       As Integer
Dim m2506ab39ff9e92c98563e1a3904df1bd                       As Integer
Dim bc575ece3d90a5a40e7a96b63c303961a              As Integer
Dim qc5842b4dc363ab44e784b14fea7b2c31             As Integer
Set b11208ee7b1b6ffdc4d54001bf42aeb1a = Application.SldWorks
Set r089f2e75bc0c2785f3502a8c94ce64fb = b11208ee7b1b6ffdc4d54001bf42aeb1a.ActiveDoc
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
r2e26fd7a88815e76612796762f046a59 = Environ("USERNAME")
b383222a9bea4084b4a55b3485db67e60 = Left(r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName, InStrRev(r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName, "\") - 1)
b383222a9bea4084b4a55b3485db67e60 = b383222a9bea4084b4a55b3485db67e60 + "\" + z56989495ec6c0363b02cdbc53cb1808d + " - Combine Log.txt"
Set rab2bd3733535f62d8ead8ad301e13be2 = New FileSystemObject
Set bfc5d43b8304fa915bdf038e967c4d02a = rab2bd3733535f62d8ead8ad301e13be2.OpenTextFile(b383222a9bea4084b4a55b3485db67e60, ForAppending, True)
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine ""
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "----------- " & r2e26fd7a88815e76612796762f046a59 & " - " & Now & " -----------"
Set z4f76b89c98a78e0601bd278fdabfe3ab = r089f2e75bc0c2785f3502a8c94ce64fb
b94b7ebfc3d693c01e015c80ebfdd0725 = r089f2e75bc0c2785f3502a8c94ce64fb.GetPathName
Dim wbdc9c654710dd90a84915649176e5f29 As ModelView
Set wbdc9c654710dd90a84915649176e5f29 = r089f2e75bc0c2785f3502a8c94ce64fb.ActiveView
wbdc9c654710dd90a84915649176e5f29.EnableGraphicsUpdate = False
z4f76b89c98a78e0601bd278fdabfe3ab.FeatureManager.EnableFeatureTree = False
qea414b44a6ee13e50a6ec2eb47ffefa7 = 0
With ProgressBar
.Bar.Width = 0
.Text.caption = "Getting Parts..."
.Text2.caption = "0% Complete"
.Text3.caption = "Processing..."
.Show vbModeless
End With
CancelButton = False
r910c0775bffa167dbee861a2869f037e = m1e98ad36788ac2ca18eb096d3a858f8e(z4f76b89c98a78e0601bd278fdabfe3ab)
bc575ece3d90a5a40e7a96b63c303961a = UBound(r910c0775bffa167dbee861a2869f037e) + 1
qc5842b4dc363ab44e784b14fea7b2c31 = (bc575ece3d90a5a40e7a96b63c303961a * (bc575ece3d90a5a40e7a96b63c303961a - 1)) / 2
b6dc70dfad7c15a723643c241dff64c41 = 0
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(r910c0775bffa167dbee861a2869f037e) - 1
Call q8b61244ae4822e03b47ad99c4fc71e0c(UBound(r910c0775bffa167dbee861a2869f037e), "Processing...")
For m2506ab39ff9e92c98563e1a3904df1bd = z57fbbe9a55b7e76e8772bb12c27d0537 To UBound(r910c0775bffa167dbee861a2869f037e)
If Not z57fbbe9a55b7e76e8772bb12c27d0537 = m2506ab39ff9e92c98563e1a3904df1bd Then
n3050a628b1fdb13ec677edd821562368 (r910c0775bffa167dbee861a2869f037e(z57fbbe9a55b7e76e8772bb12c27d0537)), (r910c0775bffa167dbee861a2869f037e(m2506ab39ff9e92c98563e1a3904df1bd))
End If
If CancelButton = True Then
Debug.Print "CANCEL IS TRUE"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "*----------- User Cancelled Macro -----------*"
GoTo ExitCode
End If
Next m2506ab39ff9e92c98563e1a3904df1bd
Next z57fbbe9a55b7e76e8772bb12c27d0537
Call q8b61244ae4822e03b47ad99c4fc71e0c(UBound(r910c0775bffa167dbee861a2869f037e), "Processing...")
ExitCode:
Call q8b61244ae4822e03b47ad99c4fc71e0c(UBound(r910c0775bffa167dbee861a2869f037e), "Rebuilding...")
Set r089f2e75bc0c2785f3502a8c94ce64fb = b11208ee7b1b6ffdc4d54001bf42aeb1a.ActivateDoc3(b94b7ebfc3d693c01e015c80ebfdd0725, False, 1, 0)
b4e5c40e176840ea491c358d7b078c013 = r089f2e75bc0c2785f3502a8c94ce64fb.EditRebuild3()
r089f2e75bc0c2785f3502a8c94ce64fb.ClearSelection2 (True)
wbdc9c654710dd90a84915649176e5f29.EnableGraphicsUpdate = True
z4f76b89c98a78e0601bd278fdabfe3ab.FeatureManager.EnableFeatureTree = True
Unload ProgressBar
Debug.Print ""
Debug.Print "----------- Macro Finished -----------"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "-----------       Macro Finished       -----------"
bfc5d43b8304fa915bdf038e967c4d02a.Close
If be23c27a34feae679476e4e9ba4fd2352 = 1 Then
MsgBox be23c27a34feae679476e4e9ba4fd2352 & " error occurred. Review log file", vbExclamation, "Error"
GoTo EndCode
ElseIf be23c27a34feae679476e4e9ba4fd2352 > 1 Then
MsgBox be23c27a34feae679476e4e9ba4fd2352 & " errors occurred. Review log file", vbExclamation, "Error"
GoTo EndCode
End If
If b253afc06d74e2ba0ae1e171327d89783 <> "" Or bea5244a27ad17c66238c31e141ee170d <> "" Then
If bea5244a27ad17c66238c31e141ee170d = "" Then
MsgBox "Excluded parts due to wrong view orientation:      " + vbCrLf + b253afc06d74e2ba0ae1e171327d89783, vbExclamation, "Error"
ElseIf b253afc06d74e2ba0ae1e171327d89783 = "" Then
MsgBox "Excluded parts due to 'Grain is Vertical' checked: " + vbCrLf + bea5244a27ad17c66238c31e141ee170d, vbExclamation, "Error"
Else
MsgBox "Excluded parts due to wrong view orientation: " + vbCrLf + b253afc06d74e2ba0ae1e171327d89783 + vbCrLf + "Excluded parts due to 'Grain is Vertical' checked: " + vbCrLf + bea5244a27ad17c66238c31e141ee170d, vbExclamation, "Excluded Parts"
End If
End If
GoTo EndCode
ErrorHandler:
Call q8b61244ae4822e03b47ad99c4fc71e0c(UBound(r910c0775bffa167dbee861a2869f037e), "Rebuilding...")
Set r089f2e75bc0c2785f3502a8c94ce64fb = b11208ee7b1b6ffdc4d54001bf42aeb1a.ActivateDoc3(b94b7ebfc3d693c01e015c80ebfdd0725, False, 1, 0)
b4e5c40e176840ea491c358d7b078c013 = r089f2e75bc0c2785f3502a8c94ce64fb.EditRebuild3()
r089f2e75bc0c2785f3502a8c94ce64fb.ClearSelection2 (True)
wbdc9c654710dd90a84915649176e5f29.EnableGraphicsUpdate = True
z4f76b89c98a78e0601bd278fdabfe3ab.FeatureManager.EnableFeatureTree = True
Unload ProgressBar
Debug.Print ""
Debug.Print "----------- ERROR while processing " & e5311fd6755b95d262631b01ddeb6773f & " and " & z31d5ee96103ab4d839c6808b87667355 & "-----------"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "----------- ERROR while processing " & e5311fd6755b95d262631b01ddeb6773f & " and " & z31d5ee96103ab4d839c6808b87667355 & "-----------"
bfc5d43b8304fa915bdf038e967c4d02a.Close
MsgBox "Error while processing: " & e5311fd6755b95d262631b01ddeb6773f & " and " & z31d5ee96103ab4d839c6808b87667355, vbExclamation, "Error"
EndCode:
End Sub
Sub n3050a628b1fdb13ec677edd821562368(Model1 As String, Model2 As String)
Dim b504941eb71f3c550d2bd388544e54c02                        As SldWorks.ModelDoc2
Dim n73400a23842fcd631ada707de45bc700                        As SldWorks.ModelDoc2
Dim b17a25919d25c2fef2bbf8a0ad44b3beb                  As ModelDocExtension
Dim b6bae2b13b21303140fcd4170d68e5995                  As ModelDocExtension
Dim ed955a66b173653f492821712e76b895d                     As SldWorks.ModelDoc2
Dim bf12210a0cb154f978ecf89039f5d5bc2                     As SldWorks.ModelDoc2
Dim b4508e8ada015179fb23d19b03762a87c                     As CustomPropertyManager
Dim wb3b45d5383d72e5b14cb45f825483dc7                     As CustomPropertyManager
Dim r510d3b707ac86b6e69098c25c51c9d8b                       As String
Dim q5ef7e34a6a5c539cfcd08096b9be4fc5                       As String
Dim bb0bc51dbe868424bcecdab271b82777b                        As Variant
Dim b5e25ddcac86b735c0eb3c51235b8cc14                        As Variant
Dim bee85bfbf3d2f21d752586b99bbf13dc9                         As SldWorks.Body2
Dim b349b060de24fda28b8abe9b7cdb8b583                         As SldWorks.Body2
Dim rffc81c76fd2991e2439e5b991443b4b6                          As SldWorks.Entity
Dim b13f198bc8fe09ccf6b9daeb837b6ae30(15)                  As Variant
Dim bf49699a88461a2a540df64abb48b2f98                      As Long
Dim b95f845acaeb1bed280585702474b8385                    As Long
Dim b4e5c40e176840ea491c358d7b078c013                      As Boolean
Dim za10e9e9c62738832edd9ee590b7d64c9                       As SldWorks.MathTransform
Dim z57fbbe9a55b7e76e8772bb12c27d0537                               As Integer
Dim m2506ab39ff9e92c98563e1a3904df1bd                               As Integer
Dim bd788acd2ab98604acf912471d7b2c47d                               As Integer
Dim b3c9ab9ec8af6cf7e866c57e77b47aa2e                     As Boolean
Dim bf363bee250a6c82e1f1d4a17a681778a                          As Long
w17ae4354e08c7a9a455c3a52c7620587 = False
e5311fd6755b95d262631b01ddeb6773f = b5ebaf9eab0956be7d743ef1391203e5e(Model1)
Set b504941eb71f3c550d2bd388544e54c02 = b11208ee7b1b6ffdc4d54001bf42aeb1a.OpenDoc6(Model1, 1, 0, "", bf49699a88461a2a540df64abb48b2f98, b95f845acaeb1bed280585702474b8385)
Set b6e7808d6dcb172b1a01afa6eaecae213 = b504941eb71f3c550d2bd388544e54c02
Set ed955a66b173653f492821712e76b895d = b504941eb71f3c550d2bd388544e54c02
Set b17a25919d25c2fef2bbf8a0ad44b3beb = ed955a66b173653f492821712e76b895d.Extension
r510d3b707ac86b6e69098c25c51c9d8b = ed955a66b173653f492821712e76b895d.MaterialIdName
Set b4508e8ada015179fb23d19b03762a87c = b17a25919d25c2fef2bbf8a0ad44b3beb.CustomPropertyManager("")
z31d5ee96103ab4d839c6808b87667355 = b5ebaf9eab0956be7d743ef1391203e5e(Model2)
Set n73400a23842fcd631ada707de45bc700 = b11208ee7b1b6ffdc4d54001bf42aeb1a.OpenDoc6(Model2, 1, 0, "", bf49699a88461a2a540df64abb48b2f98, b95f845acaeb1bed280585702474b8385)
Set nfeef4826d4eef9f613bf47c1a398604b = n73400a23842fcd631ada707de45bc700
Set bf12210a0cb154f978ecf89039f5d5bc2 = n73400a23842fcd631ada707de45bc700
Set b6bae2b13b21303140fcd4170d68e5995 = bf12210a0cb154f978ecf89039f5d5bc2.Extension
q5ef7e34a6a5c539cfcd08096b9be4fc5 = bf12210a0cb154f978ecf89039f5d5bc2.MaterialIdName
Set wb3b45d5383d72e5b14cb45f825483dc7 = b6bae2b13b21303140fcd4170d68e5995.CustomPropertyManager("")
If wb3b45d5383d72e5b14cb45f825483dc7.Get("CombineID") <> "" Then
Exit Sub
End If
If r510d3b707ac86b6e69098c25c51c9d8b <> q5ef7e34a6a5c539cfcd08096b9be4fc5 Then
Exit Sub
End If
Set bee85bfbf3d2f21d752586b99bbf13dc9 = qb28c7f492495b46037a65fb9fb0ee3a5(b6e7808d6dcb172b1a01afa6eaecae213)
Set b349b060de24fda28b8abe9b7cdb8b583 = qb28c7f492495b46037a65fb9fb0ee3a5(nfeef4826d4eef9f613bf47c1a398604b)
Set za10e9e9c62738832edd9ee590b7d64c9 = n0cf4df21da935765363783994bcc3f90(b349b060de24fda28b8abe9b7cdb8b583, bee85bfbf3d2f21d752586b99bbf13dc9)
If w17ae4354e08c7a9a455c3a52c7620587 Then
b3c9ab9ec8af6cf7e866c57e77b47aa2e = qb9424d4760062ad8fbdf4ef08ecb8674(b504941eb71f3c550d2bd388544e54c02, n73400a23842fcd631ada707de45bc700, za10e9e9c62738832edd9ee590b7d64c9, bee85bfbf3d2f21d752586b99bbf13dc9)
If w5a54905c48b2789bf04dc498478c4272 Then
Debug.Print "  Different Panel Grain"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Different Panel Grain: " + vbCrLf + "     " + e5311fd6755b95d262631b01ddeb6773f + vbCrLf + "     " + z31d5ee96103ab4d839c6808b87667355
End If
If b00bb11a073ecc16d662c97410cee82de Then
Debug.Print "  Different Laminates"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Different Laminates: " + vbCrLf + "     " + e5311fd6755b95d262631b01ddeb6773f + vbCrLf + "     " + z31d5ee96103ab4d839c6808b87667355
End If
If qfa9a074829425d28be1af4bebc0d480e Then
Debug.Print "  Different Edgebands"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Different Edgebands: " + vbCrLf + "     " + e5311fd6755b95d262631b01ddeb6773f + vbCrLf + "     " + z31d5ee96103ab4d839c6808b87667355
End If
If rd70fd074b9edb535344cb5d4dfc857e5 Then
Debug.Print "  Different Corners"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Different Corners: " + vbCrLf + "     " + e5311fd6755b95d262631b01ddeb6773f + vbCrLf + "     " + z31d5ee96103ab4d839c6808b87667355
End If
End If
If (w17ae4354e08c7a9a455c3a52c7620587 And b3c9ab9ec8af6cf7e866c57e77b47aa2e) Then
Set wb3b45d5383d72e5b14cb45f825483dc7 = b6bae2b13b21303140fcd4170d68e5995.CustomPropertyManager("")
If b4508e8ada015179fb23d19b03762a87c.Get("CombineID") = "" Then
b6dc70dfad7c15a723643c241dff64c41 = b6dc70dfad7c15a723643c241dff64c41 + 1
bf363bee250a6c82e1f1d4a17a681778a = b4508e8ada015179fb23d19b03762a87c.Add3("CombineID", swCustomInfoText, b6dc70dfad7c15a723643c241dff64c41, swCustomPropertyReplaceValue)
bf363bee250a6c82e1f1d4a17a681778a = wb3b45d5383d72e5b14cb45f825483dc7.Add3("CombineID", swCustomInfoText, b6dc70dfad7c15a723643c241dff64c41, swCustomPropertyReplaceValue)
Else
bf363bee250a6c82e1f1d4a17a681778a = wb3b45d5383d72e5b14cb45f825483dc7.Add3("CombineID", swCustomInfoText, b6dc70dfad7c15a723643c241dff64c41, swCustomPropertyReplaceValue)
End If
b6e7808d6dcb172b1a01afa6eaecae213.SetSaveFlag
nfeef4826d4eef9f613bf47c1a398604b.SetSaveFlag
Debug.Print " CombineID " & b6dc70dfad7c15a723643c241dff64c41 & ": TRUE"
Debug.Print "------------------------------------"
ElseIf w17ae4354e08c7a9a455c3a52c7620587 And b3c9ab9ec8af6cf7e866c57e77b47aa2e = False Then
Debug.Print "------------------------------------"
End If
End Sub
Function qb9424d4760062ad8fbdf4ef08ecb8674(b504941eb71f3c550d2bd388544e54c02 As SldWorks.ModelDoc2, n73400a23842fcd631ada707de45bc700 As SldWorks.ModelDoc2, za10e9e9c62738832edd9ee590b7d64c9 As SldWorks.MathTransform, bee85bfbf3d2f21d752586b99bbf13dc9 As SldWorks.Body2) As Boolean
Dim qfff1b0443a3099d302af74a74c4e84ab        As String
Dim b5e52fabdef5935f212f9becaf6f5ff65             As SldWorks.Configuration
Dim r42180a7915e4ba50380f64dff6a32cf3             As SldWorks.Configuration
Dim m48499e3196e98a71bfdc71cb71af683c         As SldWorks.CustomPropertyManager
Dim w4f359a5dade06cec87441fc6d26308e9         As SldWorks.CustomPropertyManager
Dim z57fbbe9a55b7e76e8772bb12c27d0537                   As Integer
Dim bb2bb50031f00f57a0f7f5dbf71256fc9     As Integer
Dim e7c391568aacfd1c1d91931f828d1fc59     As Integer
Dim bb5938d18987845a206e457617f62b152     As Integer
Dim w5b2aff5004ee6e2fc04bd8dbc98f9c1e     As Integer
Dim t09b12c89d051dba358beef6985f41001         As Variant
Dim b9df6f3a0bdc01553f9c94c01ec997f2a                As Variant
Dim bf89e76f2843223b073d4010f5619adfb             As Variant
Dim e46bc396232a89750402eed821be2dab7              As Variant
Dim b742b2c03ac082278faa886a731c9a386               As Variant
Dim qe52bc64dff8bd9a97018ac25f21e9e4c               As Variant
Dim b3b8baba30b0fafd4184de79823067c9e              As Variant
Dim ebf25e2690d69ccf8140d2905efae91c7         As Variant
Dim q2a7cab4307b775927748073cbf1fb115                As Variant
Dim t8cd98e81c512cebccf46e72924467fab             As Variant
Dim b080849b49c7bfb3ff8b387e7fc055623              As Variant
Dim qc99e114dcd24b059a28416f9084d6f64               As Variant
Dim be39429c24af09b18c9f5fb46ee6dce92               As Variant
Dim n702e422db41740cae94fba6d9cbe0188              As Variant
Dim td9c29382fca5518f9d8d3cdaaaaaa028            As Variant
Dim ba455d6037a9cd7ae7bce314a1a69b0ab            As Variant
Set b5e52fabdef5935f212f9becaf6f5ff65 = b504941eb71f3c550d2bd388544e54c02.GetActiveConfiguration
Set r42180a7915e4ba50380f64dff6a32cf3 = n73400a23842fcd631ada707de45bc700.GetActiveConfiguration
Set m48499e3196e98a71bfdc71cb71af683c = b5e52fabdef5935f212f9becaf6f5ff65.CustomPropertyManager
Set w4f359a5dade06cec87441fc6d26308e9 = r42180a7915e4ba50380f64dff6a32cf3.CustomPropertyManager
Debug.Print " Grain angle1: " & m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_PanelGrainAngleInFrontView")
If Not IsNumeric(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_PanelGrainAngleInFrontView")) Then
Debug.Print "Unable to get grain angle for " & e5311fd6755b95d262631b01ddeb6773f
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Unable to get grain angle for " & e5311fd6755b95d262631b01ddeb6773f & ". Check if Front View if normal to top face"
be23c27a34feae679476e4e9ba4fd2352 = be23c27a34feae679476e4e9ba4fd2352 + 1
Exit Function
Else
t09b12c89d051dba358beef6985f41001 = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_PanelGrainAngleInFrontView"), CBool(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_Ext_Core_MyGrain")))
End If
If m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_Ext_Top_MyGrain") = "" Then
b9df6f3a0bdc01553f9c94c01ec997f2a = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_TopStockMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_TopStockThickness"), "", False)
Else
b9df6f3a0bdc01553f9c94c01ec997f2a = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_TopStockMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_TopStockThickness"), CDbl(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_TopStockGrainAngleInFrontView")), CBool(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_Ext_Top_MyGrain")))
End If
If m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_Ext_Bottom_MyGrain") = "" Then
bf89e76f2843223b073d4010f5619adfb = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_BottomStockMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_BottomStockThickness"), "", False)
Else
bf89e76f2843223b073d4010f5619adfb = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_BottomStockMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_BottomStockThickness"), CDbl(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_BottomStockGrainAngleInFrontView")), CBool(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_Ext_Bottom_MyGrain")))
End If
e46bc396232a89750402eed821be2dab7 = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeFrontMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeFrontThickness"))
b3b8baba30b0fafd4184de79823067c9e = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeRightMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeRightThickness"))
b742b2c03ac082278faa886a731c9a386 = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeBackMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeBackThickness"))
qe52bc64dff8bd9a97018ac25f21e9e4c = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeLeftMaterial"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeLeftThickness"))
td9c29382fca5518f9d8d3cdaaaaaa028 = Array(m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeCornerFR"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeCornerRB"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeCornerBL"), m48499e3196e98a71bfdc71cb71af683c.Get("SWOODCP_EdgeCornerLF"))
Debug.Print " Grain angle2: " & w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_PanelGrainAngleInFrontView")
If Not IsNumeric(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_PanelGrainAngleInFrontView")) Then
Debug.Print "Unable to get grain angle for " & z31d5ee96103ab4d839c6808b87667355
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Unable to get grain angle for " & z31d5ee96103ab4d839c6808b87667355 & ". Check if Front View if normal to top face"
be23c27a34feae679476e4e9ba4fd2352 = be23c27a34feae679476e4e9ba4fd2352 + 1
Exit Function
Else
ebf25e2690d69ccf8140d2905efae91c7 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_PanelGrainAngleInFrontView"), CBool(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_Ext_Core_MyGrain")))
End If
If w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_Ext_Top_MyGrain") = "" Then
q2a7cab4307b775927748073cbf1fb115 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_TopStockMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_TopStockThickness"), "", False)
Else
q2a7cab4307b775927748073cbf1fb115 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_TopStockMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_TopStockThickness"), CDbl(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_TopStockGrainAngleInFrontView")), CBool(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_Ext_Top_MyGrain")))
End If
If w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_Ext_Bottom_MyGrain") = "" Then
t8cd98e81c512cebccf46e72924467fab = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_BottomStockMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_BottomStockThickness"), "", False)
Else
t8cd98e81c512cebccf46e72924467fab = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_BottomStockMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_BottomStockThickness"), CDbl(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_BottomStockGrainAngleInFrontView")), CBool(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_Ext_Bottom_MyGrain")))
End If
b080849b49c7bfb3ff8b387e7fc055623 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeFrontMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeFrontThickness"))
n702e422db41740cae94fba6d9cbe0188 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeRightMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeRightThickness"))
qc99e114dcd24b059a28416f9084d6f64 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeBackMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeBackThickness"))
be39429c24af09b18c9f5fb46ee6dce92 = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeLeftMaterial"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeLeftThickness"))
ba455d6037a9cd7ae7bce314a1a69b0ab = Array(w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeCornerFR"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeCornerRB"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeCornerBL"), w4f359a5dade06cec87441fc6d26308e9.Get("SWOODCP_EdgeCornerLF"))
If e46bc396232a89750402eed821be2dab7(0) <> "" Then bb2bb50031f00f57a0f7f5dbf71256fc9 = bb2bb50031f00f57a0f7f5dbf71256fc9 + 1
If b742b2c03ac082278faa886a731c9a386(0) <> "" Then bb2bb50031f00f57a0f7f5dbf71256fc9 = bb2bb50031f00f57a0f7f5dbf71256fc9 + 1
If qe52bc64dff8bd9a97018ac25f21e9e4c(0) <> "" Then bb2bb50031f00f57a0f7f5dbf71256fc9 = bb2bb50031f00f57a0f7f5dbf71256fc9 + 1
If b3b8baba30b0fafd4184de79823067c9e(0) <> "" Then bb2bb50031f00f57a0f7f5dbf71256fc9 = bb2bb50031f00f57a0f7f5dbf71256fc9 + 1
If b080849b49c7bfb3ff8b387e7fc055623(0) <> "" Then e7c391568aacfd1c1d91931f828d1fc59 = e7c391568aacfd1c1d91931f828d1fc59 + 1
If qc99e114dcd24b059a28416f9084d6f64(0) <> "" Then e7c391568aacfd1c1d91931f828d1fc59 = e7c391568aacfd1c1d91931f828d1fc59 + 1
If be39429c24af09b18c9f5fb46ee6dce92(0) <> "" Then e7c391568aacfd1c1d91931f828d1fc59 = e7c391568aacfd1c1d91931f828d1fc59 + 1
If n702e422db41740cae94fba6d9cbe0188(0) <> "" Then e7c391568aacfd1c1d91931f828d1fc59 = e7c391568aacfd1c1d91931f828d1fc59 + 1
If b9df6f3a0bdc01553f9c94c01ec997f2a(0) <> "" Then bb5938d18987845a206e457617f62b152 = bb5938d18987845a206e457617f62b152 + 1
If bf89e76f2843223b073d4010f5619adfb(0) <> "" Then bb5938d18987845a206e457617f62b152 = bb5938d18987845a206e457617f62b152 + 1
If q2a7cab4307b775927748073cbf1fb115(0) <> "" Then w5b2aff5004ee6e2fc04bd8dbc98f9c1e = w5b2aff5004ee6e2fc04bd8dbc98f9c1e + 1
If t8cd98e81c512cebccf46e72924467fab(0) <> "" Then w5b2aff5004ee6e2fc04bd8dbc98f9c1e = w5b2aff5004ee6e2fc04bd8dbc98f9c1e + 1
If bb2bb50031f00f57a0f7f5dbf71256fc9 <> e7c391568aacfd1c1d91931f828d1fc59 Then
qb9424d4760062ad8fbdf4ef08ecb8674 = False
Debug.Print " Different number of Edgebands"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Different number of Edgebands: " + vbCrLf + "     " + e5311fd6755b95d262631b01ddeb6773f + vbCrLf + "     " + z31d5ee96103ab4d839c6808b87667355
Exit Function
End If
If bb5938d18987845a206e457617f62b152 <> w5b2aff5004ee6e2fc04bd8dbc98f9c1e Then
qb9424d4760062ad8fbdf4ef08ecb8674 = False
Debug.Print " Different number of  Laminates"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Different number of  Laminates: " + vbCrLf + "     " + e5311fd6755b95d262631b01ddeb6773f + vbCrLf + "     " + z31d5ee96103ab4d839c6808b87667355
Exit Function
End If
If t09b12c89d051dba358beef6985f41001(0) = 90 Then
Dim q306397456f5470cd57651ada43df6212        As Variant
Dim b008cc4c77c5624a19986e5083c5595ec          As Variant
q306397456f5470cd57651ada43df6212 = qe52bc64dff8bd9a97018ac25f21e9e4c
qe52bc64dff8bd9a97018ac25f21e9e4c = b742b2c03ac082278faa886a731c9a386
b742b2c03ac082278faa886a731c9a386 = b3b8baba30b0fafd4184de79823067c9e
b3b8baba30b0fafd4184de79823067c9e = e46bc396232a89750402eed821be2dab7
e46bc396232a89750402eed821be2dab7 = q306397456f5470cd57651ada43df6212
b008cc4c77c5624a19986e5083c5595ec = td9c29382fca5518f9d8d3cdaaaaaa028(3)
td9c29382fca5518f9d8d3cdaaaaaa028(3) = td9c29382fca5518f9d8d3cdaaaaaa028(2)
td9c29382fca5518f9d8d3cdaaaaaa028(2) = td9c29382fca5518f9d8d3cdaaaaaa028(1)
td9c29382fca5518f9d8d3cdaaaaaa028(1) = td9c29382fca5518f9d8d3cdaaaaaa028(0)
td9c29382fca5518f9d8d3cdaaaaaa028(0) = b008cc4c77c5624a19986e5083c5595ec
End If
If ebf25e2690d69ccf8140d2905efae91c7(0) = 90 Then
q306397456f5470cd57651ada43df6212 = be39429c24af09b18c9f5fb46ee6dce92
be39429c24af09b18c9f5fb46ee6dce92 = qc99e114dcd24b059a28416f9084d6f64
qc99e114dcd24b059a28416f9084d6f64 = n702e422db41740cae94fba6d9cbe0188
n702e422db41740cae94fba6d9cbe0188 = b080849b49c7bfb3ff8b387e7fc055623
b080849b49c7bfb3ff8b387e7fc055623 = q306397456f5470cd57651ada43df6212
b008cc4c77c5624a19986e5083c5595ec = ba455d6037a9cd7ae7bce314a1a69b0ab(3)
ba455d6037a9cd7ae7bce314a1a69b0ab(3) = ba455d6037a9cd7ae7bce314a1a69b0ab(2)
ba455d6037a9cd7ae7bce314a1a69b0ab(2) = ba455d6037a9cd7ae7bce314a1a69b0ab(1)
ba455d6037a9cd7ae7bce314a1a69b0ab(1) = ba455d6037a9cd7ae7bce314a1a69b0ab(0)
ba455d6037a9cd7ae7bce314a1a69b0ab(0) = b008cc4c77c5624a19986e5083c5595ec
End If
qfff1b0443a3099d302af74a74c4e84ab = b8ec1e6e28940df08b70ad4d5cc5b689c(bee85bfbf3d2f21d752586b99bbf13dc9)
Debug.Print " Panel Type : " & qfff1b0443a3099d302af74a74c4e84ab
Dim r3a464671e81d8c379ce53a814ac5bffa            As SldWorks.mathUtility
Dim b10901ad311a7e4ff1878f4e8a0a6b82c(15)              As Double
Dim b2ae956b0de2b64efadbc40c5eedd6063       As SldWorks.MathTransform
Set r3a464671e81d8c379ce53a814ac5bffa = b11208ee7b1b6ffdc4d54001bf42aeb1a.GetMathUtility
If qfff1b0443a3099d302af74a74c4e84ab = "Unique" Then
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(za10e9e9c62738832edd9ee590b7d64c9, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb2bb50031f00f57a0f7f5dbf71256fc9, bb5938d18987845a206e457617f62b152, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
Exit Function
End If
If qfff1b0443a3099d302af74a74c4e84ab = "Rotatable" Then
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(za10e9e9c62738832edd9ee590b7d64c9, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb2bb50031f00f57a0f7f5dbf71256fc9, bb5938d18987845a206e457617f62b152, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = 1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
End If
If qfff1b0443a3099d302af74a74c4e84ab = "Rotatable and Flippable" Then
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(za10e9e9c62738832edd9ee590b7d64c9, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb2bb50031f00f57a0f7f5dbf71256fc9, bb5938d18987845a206e457617f62b152, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = 1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = -1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = -1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
End If
If qfff1b0443a3099d302af74a74c4e84ab = "Fully Symmetric" Then
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(za10e9e9c62738832edd9ee590b7d64c9, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb2bb50031f00f57a0f7f5dbf71256fc9, bb5938d18987845a206e457617f62b152, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = 1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = 1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = 1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = -1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = -1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = -1: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = -1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
b10901ad311a7e4ff1878f4e8a0a6b82c(0) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(1) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(2) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(3) = 1: b10901ad311a7e4ff1878f4e8a0a6b82c(4) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(5) = 0:
b10901ad311a7e4ff1878f4e8a0a6b82c(6) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(7) = 0: b10901ad311a7e4ff1878f4e8a0a6b82c(8) = -1:
Set b2ae956b0de2b64efadbc40c5eedd6063 = r3a464671e81d8c379ce53a814ac5bffa.CreateTransform(b10901ad311a7e4ff1878f4e8a0a6b82c)
Set b2ae956b0de2b64efadbc40c5eedd6063 = za10e9e9c62738832edd9ee590b7d64c9.Multiply(b2ae956b0de2b64efadbc40c5eedd6063)
qb9424d4760062ad8fbdf4ef08ecb8674 = w7e23a258f71b52e6440f61446637290f(b2ae956b0de2b64efadbc40c5eedd6063, t09b12c89d051dba358beef6985f41001, ebf25e2690d69ccf8140d2905efae91c7, bb5938d18987845a206e457617f62b152, bb2bb50031f00f57a0f7f5dbf71256fc9, b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115, bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab, e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623, b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64, qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92, b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188, td9c29382fca5518f9d8d3cdaaaaaa028, ba455d6037a9cd7ae7bce314a1a69b0ab)
If qb9424d4760062ad8fbdf4ef08ecb8674 Then Exit Function
End If
End Function
Function w7e23a258f71b52e6440f61446637290f(za10e9e9c62738832edd9ee590b7d64c9 As SldWorks.MathTransform, t09b12c89d051dba358beef6985f41001 As Variant, ebf25e2690d69ccf8140d2905efae91c7 As Variant, bb5938d18987845a206e457617f62b152 As Integer, bb2bb50031f00f57a0f7f5dbf71256fc9 As Integer, b9df6f3a0bdc01553f9c94c01ec997f2a As Variant, q2a7cab4307b775927748073cbf1fb115 As Variant, bf89e76f2843223b073d4010f5619adfb As Variant, t8cd98e81c512cebccf46e72924467fab As Variant, e46bc396232a89750402eed821be2dab7 As Variant, b080849b49c7bfb3ff8b387e7fc055623 As Variant, b742b2c03ac082278faa886a731c9a386 As Variant, qc99e114dcd24b059a28416f9084d6f64 As Variant, qe52bc64dff8bd9a97018ac25f21e9e4c As Variant, be39429c24af09b18c9f5fb46ee6dce92 As Variant, b3b8baba30b0fafd4184de79823067c9e As Variant, n702e422db41740cae94fba6d9cbe0188 As Variant, td9c29382fca5518f9d8d3cdaaaaaa028 As Variant, ba455d6037a9cd7ae7bce314a1a69b0ab As Variant) As Boolean
Dim r8cf92a7f551cee0a5b92640f7fa9e176      As Double
Dim b67d0b9127a68904ceb42f67ae509e202             As Variant
Dim n97c0ae247a434d657b60b634db90a18a          As Variant
w5a54905c48b2789bf04dc498478c4272 = False
b00bb11a073ecc16d662c97410cee82de = False
qfa9a074829425d28be1af4bebc0d480e = False
rd70fd074b9edb535344cb5d4dfc857e5 = False
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(8), 5) = 1 Then
If za10e9e9c62738832edd9ee590b7d64c9.ArrayData(0) = 1 And za10e9e9c62738832edd9ee590b7d64c9.ArrayData(4) = 1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
If Not (t09b12c89d051dba358beef6985f41001(0) = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
If Not (bcb8c6c165563621a142e2b497b3c8f14(b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115) And bcb8c6c165563621a142e2b497b3c8f14(bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92) Then
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> ba455d6037a9cd7ae7bce314a1a69b0ab(0) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> ba455d6037a9cd7ae7bce314a1a69b0ab(1) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> ba455d6037a9cd7ae7bce314a1a69b0ab(2) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> ba455d6037a9cd7ae7bce314a1a69b0ab(3) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(0), 5) = -1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(4), 5) = -1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
If Not (t09b12c89d051dba358beef6985f41001(0) = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
If Not (bcb8c6c165563621a142e2b497b3c8f14(b9df6f3a0bdc01553f9c94c01ec997f2a, q2a7cab4307b775927748073cbf1fb115) And bcb8c6c165563621a142e2b497b3c8f14(bf89e76f2843223b073d4010f5619adfb, t8cd98e81c512cebccf46e72924467fab)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, qc99e114dcd24b059a28416f9084d6f64) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, b080849b49c7bfb3ff8b387e7fc055623) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, be39429c24af09b18c9f5fb46ee6dce92) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, n702e422db41740cae94fba6d9cbe0188) Then
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> ba455d6037a9cd7ae7bce314a1a69b0ab(2) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> ba455d6037a9cd7ae7bce314a1a69b0ab(3) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> ba455d6037a9cd7ae7bce314a1a69b0ab(0) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> ba455d6037a9cd7ae7bce314a1a69b0ab(1) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(1), 5) = 1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(3), 5) = -1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
r8cf92a7f551cee0a5b92640f7fa9e176 = t09b12c89d051dba358beef6985f41001(0) + 90
If r8cf92a7f551cee0a5b92640f7fa9e176 = 180 Then
r8cf92a7f551cee0a5b92640f7fa9e176 = 0
End If
If Not (r8cf92a7f551cee0a5b92640f7fa9e176 = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
b67d0b9127a68904ceb42f67ae509e202 = b9df6f3a0bdc01553f9c94c01ec997f2a
n97c0ae247a434d657b60b634db90a18a = bf89e76f2843223b073d4010f5619adfb
b67d0b9127a68904ceb42f67ae509e202(2) = Abs(b9df6f3a0bdc01553f9c94c01ec997f2a(2) + 90)
n97c0ae247a434d657b60b634db90a18a(2) = Abs(bf89e76f2843223b073d4010f5619adfb(2) + 90)
If b67d0b9127a68904ceb42f67ae509e202(2) = 180 Then
b67d0b9127a68904ceb42f67ae509e202(2) = 0
End If
If n97c0ae247a434d657b60b634db90a18a(2) = 180 Then
n97c0ae247a434d657b60b634db90a18a(2) = 0
End If
If Not (bcb8c6c165563621a142e2b497b3c8f14(b67d0b9127a68904ceb42f67ae509e202, q2a7cab4307b775927748073cbf1fb115) And bcb8c6c165563621a142e2b497b3c8f14(n97c0ae247a434d657b60b634db90a18a, t8cd98e81c512cebccf46e72924467fab)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, n702e422db41740cae94fba6d9cbe0188) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, be39429c24af09b18c9f5fb46ee6dce92) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, qc99e114dcd24b059a28416f9084d6f64) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, b080849b49c7bfb3ff8b387e7fc055623) Then
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> ba455d6037a9cd7ae7bce314a1a69b0ab(1) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> ba455d6037a9cd7ae7bce314a1a69b0ab(2) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> ba455d6037a9cd7ae7bce314a1a69b0ab(3) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> ba455d6037a9cd7ae7bce314a1a69b0ab(0) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(1), 5) = -1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(3), 5) = 1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
r8cf92a7f551cee0a5b92640f7fa9e176 = Abs(t09b12c89d051dba358beef6985f41001(0) - 90)
If Not (r8cf92a7f551cee0a5b92640f7fa9e176 = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
b67d0b9127a68904ceb42f67ae509e202 = b9df6f3a0bdc01553f9c94c01ec997f2a
n97c0ae247a434d657b60b634db90a18a = bf89e76f2843223b073d4010f5619adfb
b67d0b9127a68904ceb42f67ae509e202(2) = Abs(b9df6f3a0bdc01553f9c94c01ec997f2a(2) - 90)
n97c0ae247a434d657b60b634db90a18a(2) = Abs(bf89e76f2843223b073d4010f5619adfb(2) - 90)
If b67d0b9127a68904ceb42f67ae509e202(2) = 180 Then
b67d0b9127a68904ceb42f67ae509e202(2) = 0
End If
If n97c0ae247a434d657b60b634db90a18a(2) = 180 Then
n97c0ae247a434d657b60b634db90a18a(2) = 0
End If
If Not (bcb8c6c165563621a142e2b497b3c8f14(b67d0b9127a68904ceb42f67ae509e202, q2a7cab4307b775927748073cbf1fb115) And bcb8c6c165563621a142e2b497b3c8f14(n97c0ae247a434d657b60b634db90a18a, t8cd98e81c512cebccf46e72924467fab)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, be39429c24af09b18c9f5fb46ee6dce92) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, n702e422db41740cae94fba6d9cbe0188) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, b080849b49c7bfb3ff8b387e7fc055623) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, qc99e114dcd24b059a28416f9084d6f64) Then
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> ba455d6037a9cd7ae7bce314a1a69b0ab(3) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> ba455d6037a9cd7ae7bce314a1a69b0ab(0) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> ba455d6037a9cd7ae7bce314a1a69b0ab(1) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> ba455d6037a9cd7ae7bce314a1a69b0ab(2) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
ElseIf Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(8), 5) = -1 Then
Dim bd9af0e1ae47155b61bb1e6d92b4573e8          As Variant
bd9af0e1ae47155b61bb1e6d92b4573e8 = ba455d6037a9cd7ae7bce314a1a69b0ab
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(0), 5) = -1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(4), 5) = 1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
If Not (t09b12c89d051dba358beef6985f41001(0) = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w7e23a258f71b52e6440f61446637290f = False
w5a54905c48b2789bf04dc498478c4272 = True
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
If Not (bcb8c6c165563621a142e2b497b3c8f14(b9df6f3a0bdc01553f9c94c01ec997f2a, t8cd98e81c512cebccf46e72924467fab) And bcb8c6c165563621a142e2b497b3c8f14(bf89e76f2843223b073d4010f5619adfb, q2a7cab4307b775927748073cbf1fb115)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, b080849b49c7bfb3ff8b387e7fc055623) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, qc99e114dcd24b059a28416f9084d6f64) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, be39429c24af09b18c9f5fb46ee6dce92) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, n702e422db41740cae94fba6d9cbe0188) Then
md4de103e6023a643680ab67abc733350 bd9af0e1ae47155b61bb1e6d92b4573e8
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> bd9af0e1ae47155b61bb1e6d92b4573e8(3) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> bd9af0e1ae47155b61bb1e6d92b4573e8(2) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> bd9af0e1ae47155b61bb1e6d92b4573e8(1) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> bd9af0e1ae47155b61bb1e6d92b4573e8(0) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(1), 5) = -1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(3), 5) = -1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
r8cf92a7f551cee0a5b92640f7fa9e176 = Abs(t09b12c89d051dba358beef6985f41001(0) - 90)
If Not (r8cf92a7f551cee0a5b92640f7fa9e176 = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
b67d0b9127a68904ceb42f67ae509e202 = b9df6f3a0bdc01553f9c94c01ec997f2a
n97c0ae247a434d657b60b634db90a18a = bf89e76f2843223b073d4010f5619adfb
b67d0b9127a68904ceb42f67ae509e202(2) = Abs(b9df6f3a0bdc01553f9c94c01ec997f2a(2) - 90)
n97c0ae247a434d657b60b634db90a18a(2) = Abs(bf89e76f2843223b073d4010f5619adfb(2) - 90)
If Not (bcb8c6c165563621a142e2b497b3c8f14(b67d0b9127a68904ceb42f67ae509e202, t8cd98e81c512cebccf46e72924467fab) And bcb8c6c165563621a142e2b497b3c8f14(n97c0ae247a434d657b60b634db90a18a, q2a7cab4307b775927748073cbf1fb115)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, n702e422db41740cae94fba6d9cbe0188) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, be39429c24af09b18c9f5fb46ee6dce92) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, b080849b49c7bfb3ff8b387e7fc055623) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, qc99e114dcd24b059a28416f9084d6f64) Then
md4de103e6023a643680ab67abc733350 bd9af0e1ae47155b61bb1e6d92b4573e8
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> bd9af0e1ae47155b61bb1e6d92b4573e8(0) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> bd9af0e1ae47155b61bb1e6d92b4573e8(3) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> bd9af0e1ae47155b61bb1e6d92b4573e8(2) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> bd9af0e1ae47155b61bb1e6d92b4573e8(1) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(1), 5) = 1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(3), 5) = 1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
r8cf92a7f551cee0a5b92640f7fa9e176 = t09b12c89d051dba358beef6985f41001(0) + 90
If r8cf92a7f551cee0a5b92640f7fa9e176 = 180 Then
r8cf92a7f551cee0a5b92640f7fa9e176 = 0
End If
If Not r8cf92a7f551cee0a5b92640f7fa9e176 = ebf25e2690d69ccf8140d2905efae91c7(0) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
b67d0b9127a68904ceb42f67ae509e202 = b9df6f3a0bdc01553f9c94c01ec997f2a
n97c0ae247a434d657b60b634db90a18a = bf89e76f2843223b073d4010f5619adfb
b67d0b9127a68904ceb42f67ae509e202(2) = Abs(b9df6f3a0bdc01553f9c94c01ec997f2a(2) + 90)
n97c0ae247a434d657b60b634db90a18a(2) = Abs(bf89e76f2843223b073d4010f5619adfb(2) + 90)
If b67d0b9127a68904ceb42f67ae509e202(2) = 180 Then
b67d0b9127a68904ceb42f67ae509e202(2) = 0
End If
If n97c0ae247a434d657b60b634db90a18a(2) = 180 Then
n97c0ae247a434d657b60b634db90a18a(2) = 0
End If
If Not (bcb8c6c165563621a142e2b497b3c8f14(b67d0b9127a68904ceb42f67ae509e202, t8cd98e81c512cebccf46e72924467fab) And bcb8c6c165563621a142e2b497b3c8f14(n97c0ae247a434d657b60b634db90a18a, q2a7cab4307b775927748073cbf1fb115)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, be39429c24af09b18c9f5fb46ee6dce92) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, n702e422db41740cae94fba6d9cbe0188) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, qc99e114dcd24b059a28416f9084d6f64) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, b080849b49c7bfb3ff8b387e7fc055623) Then
md4de103e6023a643680ab67abc733350 bd9af0e1ae47155b61bb1e6d92b4573e8
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> bd9af0e1ae47155b61bb1e6d92b4573e8(2) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> bd9af0e1ae47155b61bb1e6d92b4573e8(1) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> bd9af0e1ae47155b61bb1e6d92b4573e8(0) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> bd9af0e1ae47155b61bb1e6d92b4573e8(3) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
If Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(0), 5) = 1 And Round(za10e9e9c62738832edd9ee590b7d64c9.ArrayData(4), 5) = -1 Then
If (t09b12c89d051dba358beef6985f41001(1) And ebf25e2690d69ccf8140d2905efae91c7(1)) Then
If Not (t09b12c89d051dba358beef6985f41001(0) = ebf25e2690d69ccf8140d2905efae91c7(0)) Then
w5a54905c48b2789bf04dc498478c4272 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb5938d18987845a206e457617f62b152 > 0 Then
If Not (bcb8c6c165563621a142e2b497b3c8f14(b9df6f3a0bdc01553f9c94c01ec997f2a, t8cd98e81c512cebccf46e72924467fab) And bcb8c6c165563621a142e2b497b3c8f14(bf89e76f2843223b073d4010f5619adfb, q2a7cab4307b775927748073cbf1fb115)) Then
b00bb11a073ecc16d662c97410cee82de = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
End If
If bb2bb50031f00f57a0f7f5dbf71256fc9 > 0 Then
If bcb8c6c165563621a142e2b497b3c8f14(e46bc396232a89750402eed821be2dab7, qc99e114dcd24b059a28416f9084d6f64) And bcb8c6c165563621a142e2b497b3c8f14(b742b2c03ac082278faa886a731c9a386, b080849b49c7bfb3ff8b387e7fc055623) And bcb8c6c165563621a142e2b497b3c8f14(b3b8baba30b0fafd4184de79823067c9e, n702e422db41740cae94fba6d9cbe0188) And bcb8c6c165563621a142e2b497b3c8f14(qe52bc64dff8bd9a97018ac25f21e9e4c, be39429c24af09b18c9f5fb46ee6dce92) Then
md4de103e6023a643680ab67abc733350 bd9af0e1ae47155b61bb1e6d92b4573e8
If td9c29382fca5518f9d8d3cdaaaaaa028(0) <> bd9af0e1ae47155b61bb1e6d92b4573e8(1) Or td9c29382fca5518f9d8d3cdaaaaaa028(1) <> bd9af0e1ae47155b61bb1e6d92b4573e8(0) Or td9c29382fca5518f9d8d3cdaaaaaa028(2) <> bd9af0e1ae47155b61bb1e6d92b4573e8(3) Or td9c29382fca5518f9d8d3cdaaaaaa028(3) <> bd9af0e1ae47155b61bb1e6d92b4573e8(2) Then
rd70fd074b9edb535344cb5d4dfc857e5 = True
w7e23a258f71b52e6440f61446637290f = False
Exit Function
End If
w7e23a258f71b52e6440f61446637290f = True
Exit Function
Else
qfa9a074829425d28be1af4bebc0d480e = True
End If
Else
w7e23a258f71b52e6440f61446637290f = True
End If
End If
End If
End Function
Function md4de103e6023a643680ab67abc733350(Corner As Variant) As Variant
Dim b087a69b2abd16456fad7ad7c0ce23407 As Integer
For b087a69b2abd16456fad7ad7c0ce23407 = 0 To UBound(Corner)
If Corner(b087a69b2abd16456fad7ad7c0ce23407) = "1" Then
Corner(b087a69b2abd16456fad7ad7c0ce23407) = "2"
ElseIf Corner(b087a69b2abd16456fad7ad7c0ce23407) = "2" Then
Corner(b087a69b2abd16456fad7ad7c0ce23407) = "1"
ElseIf Corner(b087a69b2abd16456fad7ad7c0ce23407) = "3" Then
Corner(b087a69b2abd16456fad7ad7c0ce23407) = "4"
ElseIf Corner(b087a69b2abd16456fad7ad7c0ce23407) = "4" Then
Corner(b087a69b2abd16456fad7ad7c0ce23407) = "3"
End If
Next b087a69b2abd16456fad7ad7c0ce23407
md4de103e6023a643680ab67abc733350 = Corner
End Function
Function m1e98ad36788ac2ca18eb096d3a858f8e(Assembly As SldWorks.AssemblyDoc) As Variant
Dim rdb9aa470bf04fe7ac33cf99555c952a3                      As Object
Dim bccb524dabc70974b6774d4ea371f9aa0                   As Object
Dim bf00650a4e09a527cf4d146fe07aebc73                          As Variant
Dim bf79075ea2e17d0bf2eb4ec624b45fd6f                          As SldWorks.Component2
Dim z57fbbe9a55b7e76e8772bb12c27d0537                               As Integer
Dim m2506ab39ff9e92c98563e1a3904df1bd                               As Integer
Dim e4d103366b2c0dcd16b291e93d75211ad                            As String
Dim m24beff3e0c5c6f9d0fc822a626d5d6dd                      As SldWorks.ModelDoc2
Dim n5365312f3a16ea17d2b92ef20fa77d84                   As ModelDocExtension
Dim b75677af635f8cc2aab0c89a1ddfa26bb                      As CustomPropertyManager
Dim b678655fe005c20a0b99ebd51523838ca                   As SldWorks.CustomPropertyManager
Dim b0163f709bd9330c19683ca6a927f0628                        As String
Set rdb9aa470bf04fe7ac33cf99555c952a3 = CreateObject("Scripting.Dictionary")
Set bccb524dabc70974b6774d4ea371f9aa0 = CreateObject("Scripting.Dictionary")
bf00650a4e09a527cf4d146fe07aebc73 = Assembly.GetComponents(False)
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(bf00650a4e09a527cf4d146fe07aebc73)
Set bf79075ea2e17d0bf2eb4ec624b45fd6f = bf00650a4e09a527cf4d146fe07aebc73(z57fbbe9a55b7e76e8772bb12c27d0537)
If bf79075ea2e17d0bf2eb4ec624b45fd6f.IsSuppressed = False Then
Set m24beff3e0c5c6f9d0fc822a626d5d6dd = bf79075ea2e17d0bf2eb4ec624b45fd6f.GetModelDoc2
If m24beff3e0c5c6f9d0fc822a626d5d6dd.GetType = 1 Then
Set n5365312f3a16ea17d2b92ef20fa77d84 = m24beff3e0c5c6f9d0fc822a626d5d6dd.Extension
Set b75677af635f8cc2aab0c89a1ddfa26bb = n5365312f3a16ea17d2b92ef20fa77d84.CustomPropertyManager("")
Set b678655fe005c20a0b99ebd51523838ca = m24beff3e0c5c6f9d0fc822a626d5d6dd.GetActiveConfiguration.CustomPropertyManager
e4d103366b2c0dcd16b291e93d75211ad = bf79075ea2e17d0bf2eb4ec624b45fd6f.GetPathName
b0163f709bd9330c19683ca6a927f0628 = b5ebaf9eab0956be7d743ef1391203e5e(e4d103366b2c0dcd16b291e93d75211ad)
b4e5c40e176840ea491c358d7b078c013 = b75677af635f8cc2aab0c89a1ddfa26bb.Delete2("CombineID")
If Not (b75677af635f8cc2aab0c89a1ddfa26bb.Get("Length") = "" Or b75677af635f8cc2aab0c89a1ddfa26bb.Get("Width") = "" Or b75677af635f8cc2aab0c89a1ddfa26bb.Get("Thickness") = "") Then
If Not (b75677af635f8cc2aab0c89a1ddfa26bb.Get("IS_HARDWARE") = "Yes" Or b75677af635f8cc2aab0c89a1ddfa26bb.Get("Combine") = "No" Or InStr(e4d103366b2c0dcd16b291e93d75211ad, "\Hardwares\") <> 0) Then
If rdb9aa470bf04fe7ac33cf99555c952a3.exists(e4d103366b2c0dcd16b291e93d75211ad) Then
rdb9aa470bf04fe7ac33cf99555c952a3.Item(e4d103366b2c0dcd16b291e93d75211ad) = rdb9aa470bf04fe7ac33cf99555c952a3.Item(e4d103366b2c0dcd16b291e93d75211ad) + 1
Else
If bccb524dabc70974b6774d4ea371f9aa0.exists(e4d103366b2c0dcd16b291e93d75211ad) Then
GoTo NextIteration
End If
Dim bab2a11ad6cb633ad744e1cf599d72b34 As Variant
bab2a11ad6cb633ad744e1cf599d72b34 = m24beff3e0c5c6f9d0fc822a626d5d6dd.GetStandardViewRotation(1)
If Not bcb8c6c165563621a142e2b497b3c8f14(m77fc1e371470054ae753603ff082b6de, bab2a11ad6cb633ad744e1cf599d72b34) Then
Debug.Print b0163f709bd9330c19683ca6a927f0628 & " has wrong view orientation"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Excluded due to wrong view orientation:      " & b0163f709bd9330c19683ca6a927f0628
b253afc06d74e2ba0ae1e171327d89783 = b253afc06d74e2ba0ae1e171327d89783 + "  " + b0163f709bd9330c19683ca6a927f0628 + vbCrLf
bccb524dabc70974b6774d4ea371f9aa0.Add e4d103366b2c0dcd16b291e93d75211ad, 1
GoTo NextIteration
End If
If b678655fe005c20a0b99ebd51523838ca.Get("SWOODCP_PanelGrainVertical") = "1" Then
Debug.Print b0163f709bd9330c19683ca6a927f0628 & " has 'Grain is Vertical' checked"
bfc5d43b8304fa915bdf038e967c4d02a.WriteLine "  Excluded due to 'Grain is Vertical' checked: " & b0163f709bd9330c19683ca6a927f0628
bea5244a27ad17c66238c31e141ee170d = bea5244a27ad17c66238c31e141ee170d + "  " + b0163f709bd9330c19683ca6a927f0628 + vbCrLf
bccb524dabc70974b6774d4ea371f9aa0.Add e4d103366b2c0dcd16b291e93d75211ad, 1
GoTo NextIteration
End If
rdb9aa470bf04fe7ac33cf99555c952a3.Add e4d103366b2c0dcd16b291e93d75211ad, 1
End If
Else
End If
Else
End If
End If
End If
NextIteration:
Next
Dim wae800353f68b1d379a5e185c0932b738 As Variant
wae800353f68b1d379a5e185c0932b738 = rdb9aa470bf04fe7ac33cf99555c952a3.Keys
m1e98ad36788ac2ca18eb096d3a858f8e = wae800353f68b1d379a5e185c0932b738
End Function
Function n0cf4df21da935765363783994bcc3f90(swThisBody As SldWorks.Body2, swOtherBody As SldWorks.Body2) As SldWorks.MathTransform
Dim nfae4202f6209b448fb8de3d7d4daa3c7         As SldWorks.MathTransform
If swThisBody.GetCoincidenceTransform2(swOtherBody, nfae4202f6209b448fb8de3d7d4daa3c7) Then
Set n0cf4df21da935765363783994bcc3f90 = nfae4202f6209b448fb8de3d7d4daa3c7
If Not nfae4202f6209b448fb8de3d7d4daa3c7 Is Nothing Then
Debug.Print ""
Debug.Print "------------------------------------"
Debug.Print " Model 1    : " & e5311fd6755b95d262631b01ddeb6773f
Debug.Print " Model 2    : " & z31d5ee96103ab4d839c6808b87667355
Debug.Print " Matrix     : " & True
Dim t506c74819deed62c289a276709a5f2a5            As Variant
Dim b676159c2ac9506877e6290ba8e768f58     As Long
t506c74819deed62c289a276709a5f2a5 = nfae4202f6209b448fb8de3d7d4daa3c7.ArrayData
b676159c2ac9506877e6290ba8e768f58 = t506c74819deed62c289a276709a5f2a5(0) * (t506c74819deed62c289a276709a5f2a5(4) * t506c74819deed62c289a276709a5f2a5(8) - t506c74819deed62c289a276709a5f2a5(5) * t506c74819deed62c289a276709a5f2a5(7)) - t506c74819deed62c289a276709a5f2a5(1) * (t506c74819deed62c289a276709a5f2a5(3) * t506c74819deed62c289a276709a5f2a5(8) - t506c74819deed62c289a276709a5f2a5(5) * t506c74819deed62c289a276709a5f2a5(6)) + t506c74819deed62c289a276709a5f2a5(2) * (t506c74819deed62c289a276709a5f2a5(3) * t506c74819deed62c289a276709a5f2a5(7) - t506c74819deed62c289a276709a5f2a5(4) * t506c74819deed62c289a276709a5f2a5(6))
If b676159c2ac9506877e6290ba8e768f58 = -1 Then
Debug.Print " Mirror     : " & e5311fd6755b95d262631b01ddeb6773f
Debug.Print "------------------------------------"
ElseIf t506c74819deed62c289a276709a5f2a5(12) = 1 Then
Debug.Print " CombineID " & b6dc70dfad7c15a723643c241dff64c41 & ": " & e5311fd6755b95d262631b01ddeb6773f
w17ae4354e08c7a9a455c3a52c7620587 = True
End If
End If
Else
End If
End Function
Function qb28c7f492495b46037a65fb9fb0ee3a5(Part As SldWorks.PartDoc) As SldWorks.Body2
Dim w33548d5cda925bb321073eb295ef6a62                          As SldWorks.Body2
Dim bf1492617213aa582c0775df45383c89a                      As Double
Dim bb323a9df80869e4e0b85a937f96c1909                As Integer
Dim b42c949cef725fec3e619f59152e39db5                       As Variant
Dim bfd8c89eae9803e5f705c3dd395249328                         As Variant
Dim bd788acd2ab98604acf912471d7b2c47d                               As Integer
Part.ClearSelection2 True
bfd8c89eae9803e5f705c3dd395249328 = Part.GetBodies2(swAllBodies, True)
bf1492617213aa582c0775df45383c89a = 0
For bd788acd2ab98604acf912471d7b2c47d = 0 To UBound(bfd8c89eae9803e5f705c3dd395249328)
Set w33548d5cda925bb321073eb295ef6a62 = bfd8c89eae9803e5f705c3dd395249328(bd788acd2ab98604acf912471d7b2c47d)
b42c949cef725fec3e619f59152e39db5 = w33548d5cda925bb321073eb295ef6a62.GetMassProperties(1)
If b42c949cef725fec3e619f59152e39db5(3) > bf1492617213aa582c0775df45383c89a Then
bf1492617213aa582c0775df45383c89a = b42c949cef725fec3e619f59152e39db5(3)
bb323a9df80869e4e0b85a937f96c1909 = bd788acd2ab98604acf912471d7b2c47d
End If
Next bd788acd2ab98604acf912471d7b2c47d
Set qb28c7f492495b46037a65fb9fb0ee3a5 = bfd8c89eae9803e5f705c3dd395249328(bb323a9df80869e4e0b85a937f96c1909)
End Function
Function b5ebaf9eab0956be7d743ef1391203e5e(file As String) As String
Dim b93ff3abd396e4ca22447f6d4a50ae604 As String
b93ff3abd396e4ca22447f6d4a50ae604 = Right(file, Len(file) - InStrRev(file, "\"))
If Right(b93ff3abd396e4ca22447f6d4a50ae604, 7) = ".sldasm" Or Right(b93ff3abd396e4ca22447f6d4a50ae604, 7) = ".sldprt" Then
b5ebaf9eab0956be7d743ef1391203e5e = Left(b93ff3abd396e4ca22447f6d4a50ae604, Len(b93ff3abd396e4ca22447f6d4a50ae604) - 7)
Else
b5ebaf9eab0956be7d743ef1391203e5e = b93ff3abd396e4ca22447f6d4a50ae604
End If
End Function
Function b695664185dfefc4c1ad40b10646963d0(Name As String) As String
b695664185dfefc4c1ad40b10646963d0 = Replace(Name, "/", "@")
End Function
Function q8b61244ae4822e03b47ad99c4fc71e0c(Maxiter As Integer, ByVal caption As String)
Dim t382b9d8066debe4514dc179247019dcf     As Double
Dim r2e00027eabe0b722f19a23ac90779560            As Double
Dim qddcbd44bfca946d6461895d06978f141  As Double
If caption = "Rebuilding..." Then
qea414b44a6ee13e50a6ec2eb47ffefa7 = qea414b44a6ee13e50a6ec2eb47ffefa7 - 1
ProgressBar.Bar.BackColor = &HC000&
End If
qea414b44a6ee13e50a6ec2eb47ffefa7 = qea414b44a6ee13e50a6ec2eb47ffefa7 + 1
Maxiter = Maxiter + 1
t382b9d8066debe4514dc179247019dcf = qea414b44a6ee13e50a6ec2eb47ffefa7 / Maxiter
r2e00027eabe0b722f19a23ac90779560 = ProgressBar.Frame.Width * t382b9d8066debe4514dc179247019dcf
qddcbd44bfca946d6461895d06978f141 = Round(t382b9d8066debe4514dc179247019dcf * 100, 0)
ProgressBar.Bar.Width = r2e00027eabe0b722f19a23ac90779560 - 0.015 * r2e00027eabe0b722f19a23ac90779560
ProgressBar.Text2.caption = qddcbd44bfca946d6461895d06978f141 & "% Complete"
ProgressBar.Text.caption = qea414b44a6ee13e50a6ec2eb47ffefa7 & " of " & Maxiter
ProgressBar.Text3.caption = caption
DoEvents
If (qddcbd44bfca946d6461895d06978f141 / 10) Mod 2 = 0 Then
ProgressBar.Image1.Visible = False
ProgressBar.Image2.Visible = True
Else
ProgressBar.Image1.Visible = True
ProgressBar.Image2.Visible = False
End If
End Function
Function bcb8c6c165563621a142e2b497b3c8f14(A As Variant, B As Variant) As Boolean
Dim z57fbbe9a55b7e76e8772bb12c27d0537 As Integer
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(A)
If A(z57fbbe9a55b7e76e8772bb12c27d0537) <> B(z57fbbe9a55b7e76e8772bb12c27d0537) Then
bcb8c6c165563621a142e2b497b3c8f14 = False
Exit Function
End If
Next z57fbbe9a55b7e76e8772bb12c27d0537
bcb8c6c165563621a142e2b497b3c8f14 = True
End Function
Function b8ec1e6e28940df08b70ad4d5cc5b689c(Body As SldWorks.Body2) As String
Dim w33548d5cda925bb321073eb295ef6a62                  As SldWorks.Body2
Dim b42c949cef725fec3e619f59152e39db5               As Variant
Dim b6220cae036ff69014b7204ed3a63a737(2)                  As Variant
Dim b13a6f2ed85f6f02d594fc6d79ebadf3b                     As Variant
Dim z7c43b6942b30fa7fc8e1888c5d2b8bfa(2)               As Double
Dim m46684438e4ede88e9c6df9c610ac6c30(2)              As Double
Dim z57fbbe9a55b7e76e8772bb12c27d0537                       As Integer
b42c949cef725fec3e619f59152e39db5 = Body.GetMassProperties(1)
b13a6f2ed85f6f02d594fc6d79ebadf3b = Body.GetBodyBox
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To 2
b6220cae036ff69014b7204ed3a63a737(z57fbbe9a55b7e76e8772bb12c27d0537) = b42c949cef725fec3e619f59152e39db5(z57fbbe9a55b7e76e8772bb12c27d0537)
z7c43b6942b30fa7fc8e1888c5d2b8bfa(z57fbbe9a55b7e76e8772bb12c27d0537) = (b13a6f2ed85f6f02d594fc6d79ebadf3b(z57fbbe9a55b7e76e8772bb12c27d0537 + 3) + b13a6f2ed85f6f02d594fc6d79ebadf3b(z57fbbe9a55b7e76e8772bb12c27d0537)) / 2
m46684438e4ede88e9c6df9c610ac6c30(z57fbbe9a55b7e76e8772bb12c27d0537) = Round(z7c43b6942b30fa7fc8e1888c5d2b8bfa(z57fbbe9a55b7e76e8772bb12c27d0537) - b6220cae036ff69014b7204ed3a63a737(z57fbbe9a55b7e76e8772bb12c27d0537), 8)
Next z57fbbe9a55b7e76e8772bb12c27d0537
If m46684438e4ede88e9c6df9c610ac6c30(0) = 0 And m46684438e4ede88e9c6df9c610ac6c30(1) = 0 And m46684438e4ede88e9c6df9c610ac6c30(2) = 0 And b13a6f2ed85f6f02d594fc6d79ebadf3b(3) - b13a6f2ed85f6f02d594fc6d79ebadf3b(0) = b13a6f2ed85f6f02d594fc6d79ebadf3b(4) - b13a6f2ed85f6f02d594fc6d79ebadf3b(1) Then
b8ec1e6e28940df08b70ad4d5cc5b689c = "Fully Symmetric"
Exit Function
End If
If m46684438e4ede88e9c6df9c610ac6c30(0) = 0 And m46684438e4ede88e9c6df9c610ac6c30(1) = 0 And m46684438e4ede88e9c6df9c610ac6c30(2) = 0 Then
b8ec1e6e28940df08b70ad4d5cc5b689c = "Rotatable and Flippable"
Exit Function
End If
If m46684438e4ede88e9c6df9c610ac6c30(0) = 0 And m46684438e4ede88e9c6df9c610ac6c30(1) = 0 And m46684438e4ede88e9c6df9c610ac6c30(2) <> 0 Then
b8ec1e6e28940df08b70ad4d5cc5b689c = "Rotatable"
Exit Function
End If
b8ec1e6e28940df08b70ad4d5cc5b689c = "Unique"
End Function
Function n35225f25210d92934e50a137bd4f1d09(List As Variant)
Dim z57fbbe9a55b7e76e8772bb12c27d0537               As Integer
Dim w8274db519c9ea1ef37a1a70714c4034a          As Variant
Dim q40d97d45ce0c91ae51dd3a0c556fc4f2          As Long
Dim tc9b3778edfb8354db20c422b93d2ea8d        As Long
Dim r089f2e75bc0c2785f3502a8c94ce64fb         As ModelDoc2
Dim n5365312f3a16ea17d2b92ef20fa77d84   As ModelDocExtension
Dim b75677af635f8cc2aab0c89a1ddfa26bb      As CustomPropertyManager
Dim b313252fe79dfccec0556f4a44bf0f68c       As String
Dim babc569aab41bd8f74288db6ae3ca18de       As Boolean
Dim b258239c0412d75b373e25b71f0a3e626          As String
Dim becf62aedf74c3f44c716f44bc1dac6d1  As String
Dim e80cdfe64535bd6a1dd5c09a20598d72b     As Boolean
Dim b32680ce79717419affe531c1dec41527  As Boolean
Dim w21cc3f518f9954fb6e2a373fbcc0794d          As String
Dim bf2fe97c16dae5d14b78fab741a410bb8           As String
Dim e67249c62a1b4d653b1f3c8cbb80d7772       As String
For z57fbbe9a55b7e76e8772bb12c27d0537 = 0 To UBound(List)
Set r089f2e75bc0c2785f3502a8c94ce64fb = b11208ee7b1b6ffdc4d54001bf42aeb1a.OpenDoc6(List(z57fbbe9a55b7e76e8772bb12c27d0537), swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", q40d97d45ce0c91ae51dd3a0c556fc4f2, tc9b3778edfb8354db20c422b93d2ea8d)
Set n5365312f3a16ea17d2b92ef20fa77d84 = r089f2e75bc0c2785f3502a8c94ce64fb.Extension
Set b75677af635f8cc2aab0c89a1ddfa26bb = n5365312f3a16ea17d2b92ef20fa77d84.CustomPropertyManager("")
b75677af635f8cc2aab0c89a1ddfa26bb.Get6 "Length", babc569aab41bd8f74288db6ae3ca18de, b258239c0412d75b373e25b71f0a3e626, w21cc3f518f9954fb6e2a373fbcc0794d, e80cdfe64535bd6a1dd5c09a20598d72b, b32680ce79717419affe531c1dec41527
b75677af635f8cc2aab0c89a1ddfa26bb.Get6 "Width", babc569aab41bd8f74288db6ae3ca18de, b258239c0412d75b373e25b71f0a3e626, bf2fe97c16dae5d14b78fab741a410bb8, e80cdfe64535bd6a1dd5c09a20598d72b, b32680ce79717419affe531c1dec41527
b75677af635f8cc2aab0c89a1ddfa26bb.Get6 "Thickness", babc569aab41bd8f74288db6ae3ca18de, b258239c0412d75b373e25b71f0a3e626, e67249c62a1b4d653b1f3c8cbb80d7772, e80cdfe64535bd6a1dd5c09a20598d72b, b32680ce79717419affe531c1dec41527
Debug.Print w21cc3f518f9954fb6e2a373fbcc0794d & " x " & bf2fe97c16dae5d14b78fab741a410bb8 & " x " & e67249c62a1b4d653b1f3c8cbb80d7772
w8274db519c9ea1ef37a1a70714c4034a = w21cc3f518f9954fb6e2a373fbcc0794d
Next z57fbbe9a55b7e76e8772bb12c27d0537
End Function