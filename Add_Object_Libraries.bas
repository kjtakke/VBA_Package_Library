'https://support.pcmiler.com/en/support/solutions/articles/19000047036-what-does-error-1004-programmatic-access-to-visual-basic-project-is-not-trusted-mean-

Sub addObjectLibrary()
    'File Scripting Runtime Library
    Application.VBE.ActiveVBProject.References.AddFromFile ("C:\Windows\SysWOW64\scrrun.dll")
End Sub

'COMMON GRUD REFERENCES
'Microsoft Excel        {00020813-0000-0000-C000-000000000046}
'Microsoft Word         {00020905-0000-0000-C000-000000000046}
'Microsoft PowerPoint   {91493440-5A91-11CF-8700-00AA0060263B}
'Microsoft Access       {4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}
'Microsoft Outlook      {00062FFF-0000-0000-C000-000000000046}

Function F_isReferenceAdded(referenceGUID As String) As Boolean

    Dim varRef As Variant

    'Loop through VBProject references if input GUID found return TRUE otherwise FALSE
    For Each varRef In ThisWorkbook.VBProject.References
        
        If varRef.GUID = referenceGUID Then
            F_isReferenceAdded = True
            Exit For
        End If
        
    Next varRef

End Function

Sub addObjectLibraryReference()
    Dim strGUID As String

    'Microsoft Visual Basic For Application Extensibility GUID
    strGUID = "{0002E157-0000-0000-C000-000000000046}"

    'Check if reference is already added to the project, if not add it
    If F_isReferenceAdded(strGUID) = False Then
        ThisWorkbook.VBProject.References.AddFromGuid strGUID, 0, 0
    End If
End Sub
