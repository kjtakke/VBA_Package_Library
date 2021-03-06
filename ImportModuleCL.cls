'Sub ScrapeData()
'    Dim WD As New ImportModuleCL
'    WD.HTMLScrape("Test.bas") = "https://raw.githubusercontent.com/kjtakke/Excel-VBA/master/ScrapeDel"
'End Sub

Const LineDel = "{%;%}"
Private HTML_String As String

Public Property Let HTMLScrape(fileName As String, URL As String)
    Dim HTML_Array As Variant, HTML_Text As String
    Call Get_HTML_Body(URL)
    HTML_Array = Split(HTML_String, LineDel)
    HTML_Text = arrayToText(HTML_Array)
    Call writeToTextFile(fileName, HTML_Text)
    Call InsertVBComponent(fileName)
    Call DeleteFile(fileName)
End Property

Private Sub Get_HTML_Body(URL As String)
    Dim http As Object, html As New HTMLDocument
    Dim HTMLText As String, HTMLArray As Variant
    Dim i As Integer
    On Error GoTo Error_Message:
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", URL, False
    http.send
    html.body.innerHTML = http.responseText
    HTML_String = html.body.innerHTML
    GoTo Scraped_Data:
Error_Message:
    MsgBox ("Can not access the library server at this time")
Scraped_Data:
End Sub

Private Function arrayToText(ary As Variant) As String
    On Error Resume Next
    arrayToText = ""
    For i = 0 To UBound(ary)
        arrayToText = arrayToText & ary(i) & vbNewLine
    Next i
End Function

Private Sub writeToTextFile(ByVal fileName As String, text As String)
    On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\" & Environ("UserName") & "\Desktop\" & fileName, True)
    a.WriteLine text
    a.Close
End Sub

Private Sub InsertVBComponent(ByVal CompFileName As String)
    Dim vbcomp As VBComponent, addModule As Boolean, file As Variant
    Dim modules As Collection, model As Variant
    Set modules = New Collection
    
    addModule = False
    file = Split(CompFileName, ".")
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        If ((vbcomp.Type = vbext_ct_StdModule) Or (vbcomp.Type = vbext_ct_ClassModule) Or (vbcomp.Type = vbext_ct_MSForm)) Then
           modules.Add vbcomp.Name
            If vbcomp.Name = file(0) Then
                addModule = True
                GoTo en:
            End If
        End If
    Next vbcomp
    If addModule = False Then
        If CompFileName <> "" Then
            On Error Resume Next
            Application.ActiveWorkbook.VBProject.VBComponents.Import "C:\Users\" & Environ("UserName") & "\Desktop\" & CompFileName
            On Error GoTo 0
        End If
    End If
    file = Split(CompFileName, ".")
    For Each vbcomp In ThisWorkbook.VBProject.VBComponents
        addModule = True
        If ((vbcomp.Type = vbext_ct_StdModule) Or (vbcomp.Type = vbext_ct_ClassModule) Or (vbcomp.Type = vbext_ct_MSForm)) Then
            For Each model In modules
                If vbcomp.Name = model Then
                    addModule = False
                    GoTo nxt:
                End If
            Next model
nxt:
            If addModule = True Then
                vbcomp.Name = file(0)
            End If
        End If
    Next vbcomp
en:
End Sub

Private Sub DeleteFile(ByVal FileToDelete As String)
    FileToDelete = "C:\Users\" & Environ("UserName") & "\Desktop\" & FileToDelete
    If FileExists(FileToDelete) Then
        SetAttr FileToDelete, vbNormal
        Kill FileToDelete
    End If
End Sub

Private Function FileExists(ByVal FileToTest As String) As Boolean
    FileExists = (Dir(FileToTest) <> "")
End Function


