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
End Sub


Private Sub ListBox_Packages_Click()
    Call review_version_from_list
End Sub

Private Sub Button_Review_Version_Click()
    Call review_version_from_list
End Sub

Private Sub ListBox_Versions_Click()
    Dim html As WebScrapeCL
    Set html = New WebScrapeCL
 
    Dim package_version As String, package_version_array As Variant
    html.HTMLScrape = package_versions_url & Me.ListBox_Packages.Value
    package_version = html.HTMLScrape

    Package_Versions = Split(package_version, "{%;%}")

    For i = 0 To UBound(Package_Versions)
        
        package_version_array = Split(Package_Versions(i), "{%-%}")
        
        If Me.ListBox_Versions.Value = package_version_array(0) Then
            Me.TextBox_Developer_Notes.Value = package_version_array(1)
        End If

    Next i
End Sub

Private Sub TextBox_Search_Packages_Change()

    Dim html As WebScrapeCL
    Set html = New WebScrapeCL
    Dim string_compare As Single
    
    Dim package_list As String
    html.HTMLScrape = package_list_url
    package_list = html.HTMLScrape
    
    Dim Package_Array As Variant
    Package_Array = Split(package_list, "{%;%}")
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
    Dim html As WebScrapeCL
    Set html = New WebScrapeCL
    
    Dim package_list As String
    html.HTMLScrape = package_list_url
    package_list = html.HTMLScrape
    
    Dim Package_Array As Variant
    Package_Array = Split(package_list, "{%;%}")
    
    Me.ListBox_Packages.Clear
    
    For i = 0 To UBound(Package_Array)
        Me.ListBox_Packages.AddItem (Package_Array(i))
    Next i
End Sub

Sub review_version_from_list()
    Dim html As WebScrapeCL
    Set html = New WebScrapeCL
 
    Dim package_version As Variant, package_version_array As Variant
    html.HTMLScrape = package_versions_url & Me.ListBox_Packages.Value
    package_version = html.HTMLScrape
    
    package_version_array = Split(package_version, "{%;%}")
    
    Me.ListBox_Versions.Clear
    For i = 0 To UBound(package_version_array)
        package_version = Split(package_version_array(i), "{%-%}")
        Me.ListBox_Versions.AddItem (package_version(0))
    Next i
End Sub

