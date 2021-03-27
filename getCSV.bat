Sub LoadCSVtoArray()

strPath = ThisWorkbook.Path & "\"

Set cn = CreateObject("ADODB.Connection")
strcon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
cn.Open strcon
strSQL = "SELECT * FROM SAMPLE.csv;"

Dim rs As Recordset
Dim rsARR() As Variant

Set rs = cn.Execute(strSQL)
rsARR = WorksheetFunction.Transpose(rs.GetRows)
rs.Close
Set cn = Nothing

[a1].Resize(UBound(rsARR), UBound(Application.Transpose(rsARR))) = rsARR

End Sub