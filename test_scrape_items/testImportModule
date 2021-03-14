Option Explicit


Sub Calling_Procedure()
    InsertVBComponent ActiveWorkbook, "TestImport.bas", "C:\Users\" & Environ("UserName") & "\Desktop\"
End Sub


Sub InsertVBComponent(ByVal wb As Workbook, ByVal CompFileName As String, path As String)
    Dim vbcomp As VBComponent, addModule As Boolean, file As Variant
    addModule = False
    file = Split(CompFileName, ".")
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        If ((vbcomp.Type = vbext_ct_StdModule) Or (vbcomp.Type = vbext_ct_ClassModule) Or (vbcomp.Type = vbext_ct_MSForm)) Then
            If vbcomp.Name = file(0) Then
                addModule = True
                GoTo en:
            End If
        End If
    Next vbcomp
    If addModule = False Then
        If CompFileName <> "" Then
            On Error Resume Next
            Application.ActiveWorkbook.VBProject.VBComponents.import path & CompFileName
            On Error GoTo 0
        End If
    End If
en:
End Sub
