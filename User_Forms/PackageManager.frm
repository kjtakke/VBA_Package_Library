VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PackageManager 
   Caption         =   "VBA External Package Manager"
   ClientHeight    =   8328
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14100
   OleObjectBlob   =   "PackageManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PackageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const package_list_url = "https://raw.githubusercontent.com/kjtakke/VBA_Package_Library/main/test_scrape_items/import__package_list"
Const package_versions_url = "https://raw.githubusercontent.com/kjtakke/VBA_Package_Library/main/test_scrape_items/"

Public Package_Array As Variant
Public Package_Versions As Variant
Public i As Single, j As Single, k As Single, h As Single

Private Sub UserForm_Initialize()
    Call get_package_list
End Sub

Private Sub Button_Get_A_List_of_Packages_Click()
    Call get_package_list
    Me.TextBox_Search_Packages.Value = ""
End Sub

Private Sub ListBox_Packages_Click()
    Call review_version_from_list
End Sub

Private Sub Button_Review_Version_Click()
    Call review_version_from_list
End Sub

Private Sub ListBox_Versions_Click()
    On Error Resume Next



    Dim http As Object, html As New HTMLDocument
    Dim HTMLText As String, HTMLArray As Variant
    Dim HTML_String As String
    Dim i As Integer
    On Error GoTo Error_Message:
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", package_versions_url & Me.ListBox_Packages.Value, False
    http.send
    html.body.innerHTML = http.responseText
    HTML_String = html.body.innerHTML
    GoTo Scraped_Data:
Error_Message:
    MsgBox ("Can not access the library server at this time")
Scraped_Data:
  
    Dim Package_Array As Variant, package_version As Variant
    Package_Array = Split(HTML_String, "{%;%}")
    
    Me.TextBox_Developer_Notes.Value = ""
        For i = 0 To UBound(Package_Array)
            package_version = Split(Package_Array(i), "{%-%}")
            If package_version(0) = Me.ListBox_Versions.Value Then
                Me.TextBox_Developer_Notes = package_version(1)
            End If
            
        Next i
End Sub

Private Sub TextBox_Search_Packages_Change()
    Dim http As Object, html As New HTMLDocument
    Dim HTMLText As String, HTMLArray As Variant
    Dim HTML_String As String, string_compare As Integer
    Dim i As Integer
    On Error GoTo Error_Message:
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", package_list_url, False
    http.send
    html.body.innerHTML = http.responseText
    HTML_String = html.body.innerHTML
    GoTo Scraped_Data:
Error_Message:
    MsgBox ("Can not access the library server at this time")
Scraped_Data:
  
    Dim Package_Array As Variant
    Package_Array = Split(HTML_String, "{%;%}")
    Me.ListBox_Packages.Clear
    Me.ListBox_Versions.Clear
    Me.TextBox_Developer_Notes.Value = ""
    For i = 0 To UBound(Package_Array)
        string_compare = InStr(1, Package_Array(i), Me.TextBox_Search_Packages.Value)
        If string_compare > 0 Then
            Me.ListBox_Packages.AddItem (Package_Array(i))
        End If
    Next i
End Sub

Sub get_package_list()
    Dim http As Object, html As New HTMLDocument
    Dim HTMLText As String, HTMLArray As Variant
    Dim HTML_String As String
    Dim i As Integer
    On Error GoTo Error_Message:
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", package_list_url, False
    http.send
    html.body.innerHTML = http.responseText
    HTML_String = html.body.innerHTML
    GoTo Scraped_Data:
Error_Message:
    MsgBox ("Can not access the library server at this time")
Scraped_Data:
  
    Dim Package_Array As Variant
    Package_Array = Split(HTML_String, "{%;%}")
    
    Me.ListBox_Packages.Clear
    
    For i = 0 To UBound(Package_Array)
        Me.ListBox_Packages.AddItem (Package_Array(i))
    Next i
End Sub

Sub review_version_from_list()
    Dim http As Object, html As New HTMLDocument
    Dim HTMLText As String, HTMLArray As Variant
    Dim HTML_String As String
    Dim i As Integer
    On Error GoTo Error_Message:
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", package_versions_url & Me.ListBox_Packages.Value, False
    http.send
    html.body.innerHTML = http.responseText
    HTML_String = html.body.innerHTML
    GoTo Scraped_Data:
Error_Message:
    MsgBox ("Can not access the library server at this time")
Scraped_Data:
  
    Dim Package_Array As Variant, package_version As Variant
    Package_Array = Split(HTML_String, "{%;%}")
    
    Me.ListBox_Versions.Clear
        For i = 0 To UBound(Package_Array)
            package_version = Split(Package_Array(i), "{%-%}")
            Me.ListBox_Versions.AddItem package_version(0)
        Next i
End Sub


Sub InsertVBComponentExample()
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
            Application.ActiveWorkbook.VBProject.VBComponents.Import path & CompFileName
            On Error GoTo 0
        End If
    End If
en:
End Sub
