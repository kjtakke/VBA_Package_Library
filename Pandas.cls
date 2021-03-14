'###################################################
'Variables #########################################
'###################################################

Const VBA_Array_Help_Text = ""
Const Worksheet_Help_Text = ""
Const SQL_Connection_String_Help_Text = ""
Const CSV_Help_Text = ""
Const HTML_Table_Help_Text = ""
Const JSON_Help_Text = ""
Const XML_Help_Text = ""
Const JavaSctipt_Array_Help_Text = ""
Const SharePoint_Online_List_Help_Text = ""

Public DataFrame As Variant
Public RowCount As Single
Public ColumnCount As Single
Public Headers As Variant
Public RowCountA As Variant
Public RowSum As Variant
Public RowAverage As Variant
Public RowMax As Variant
Public RowMin As Variant
Public RowMedian As Variant
Public RowBitSize As Variant
Public DataFrameBitSize As Single

'###################################################
'Enumerations ######################################
'###################################################

Public Enum DataTypes
    VBA_Array
    Worksheet
    SQL_Connection_String
    CSV
    HTML_Table
    JSON
    XML
    JavaSctipt_Array
    SharePoint_Online_List
End Enum

Public Enum Joins
    Inner
    Outer
    Full
End Enum

Public Enum DataFrame_Import_Help_Text
    VBA_Array
    Worksheet
    SQL_Connection_String
    CSV
    HTML_Table
    JSON
    XML
    JavaSctipt_Array
    SharePoint_Online_List
End Enum

'###################################################
'DataFrame Import ##################################
'###################################################
 
Property Let DataFrame_Import(header As Boolean, Import_Method As DataTypes, data As Variant)
    
    'Import Type
    Select Case True
        Case Import_Method = CSV
        
        Case Import_Method = HTML_Table
        
        Case Import_Method = JavaSctipt_Array
        
        Case Import_Method = JSON
        
        Case Import_Method = SharePoint_Online_List
        
        Case Import_Method = SQL_Connection_String
        
        Case Import_Method = VBA_Array
        
        Case Import_Method = Worksheet
        
        Case Import_Method = XML
    
    End Select
    
    'Add Method Values
    DataFrame = ""
    
End Property

'###################################################
'Head ##############################################
'###################################################

Public Property Let Head(Rows As Integer)
    Debug.Print ("")
End Property

'###################################################
'Joins #############################################
'###################################################

Property Let DataFrame_Join(JoinType As Joins, DataFrameColumns As String, NewDataFrameColumns As String, NewDataFrame As Variant)
    Select Case True
        Case JoinType = Full
        
        Case JoinType = Inner
        
        Case JoinType = Outer
    
    End Select
End Property

'###################################################
'Pivot #############################################
'###################################################






'###################################################
'Plots #############################################
'###################################################






'###################################################
'Headers ###########################################
'###################################################






'###################################################
'Columns ###########################################
'###################################################






'###################################################
'Rows ##############################################
'###################################################






'###################################################
'Sort ##############################################
'###################################################






'###################################################
'Filter ############################################
'###################################################






'###################################################
'Export ############################################
'###################################################

Property Let DataFrame_Export(header As Boolean, Import_Method As DataTypes, data As Variant)
    
    'Import Type
    Select Case True
        Case Import_Method = CSV
        
        Case Import_Method = HTML_Table
        
        Case Import_Method = JavaSctipt_Array
        
        Case Import_Method = JSON
        
        Case Import_Method = SharePoint_Online_List
        
        Case Import_Method = SQL_Connection_String
        
        Case Import_Method = VBA_Array
        
        Case Import_Method = Worksheet
        
        Case Import_Method = XML
    
    End Select
    
    'Add Method Values
    DataFrame = ""
    
End Property





'###################################################
'Help ##############################################
'###################################################

Property Let DataFrame_Import_Help(Import_Method As DataFrame_Import_Help_Text)
    
    'Import Type
    Select Case True
        Case Import_Method = CSV
            
        Case Import_Method = HTML_Table
        
        Case Import_Method = JavaSctipt_Array
        
        Case Import_Method = JSON
