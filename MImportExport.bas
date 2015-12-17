Attribute VB_Name = "MImportExport"
Option Explicit


Public Sub ExportSourceFiles(destPath As String)
 
Dim component As VBComponent
For Each component In Application.VBE.ActiveVBProject.VBComponents
If component.Type = vbext_ct_ClassModule Or component.Type = vbext_ct_StdModule Then
component.Export destPath & component.Name & ToFileExtension(component.Type)
End If
Next
 
End Sub
 
Private Function ToFileExtension(vbeComponentType As vbext_ComponentType) As String
Select Case vbeComponentType
Case vbext_ComponentType.vbext_ct_ClassModule
ToFileExtension = ".cls"
Case vbext_ComponentType.vbext_ct_StdModule
ToFileExtension = ".bas"
Case vbext_ComponentType.vbext_ct_MSForm
ToFileExtension = ".frm"
Case vbext_ComponentType.vbext_ct_ActiveXDesigner
Case vbext_ComponentType.vbext_ct_Document
Case Else
ToFileExtension = vbNullString
End Select
 
End Function

