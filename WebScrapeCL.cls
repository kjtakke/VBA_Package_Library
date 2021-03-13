'###################################################
'HTML SCRAPE AND RETURN AS STRING ##################
'###################################################

Private HTML_String As String 'HTML String Variable

Public Property Let HTMLScrape(URL As String)
    'Pass the URL into the Scrape Sub
    Call Get_HTML_Body(URL)
End Property

Public Property Get HTMLScrape() As String
    'Return the scraped data
    HTMLScrape = HTML_String
End Property

Private Sub Get_HTML_Body(URL As String)
    'Get the HTML Body from a URL
    
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
    msgbox ("Can not access the library server at this time")
    
Scraped_Data:

End Sub

'###################################################
'WRITE TO FILE #####################################
'###################################################






'###################################################
'DATA TO ARRAY #####################################
'###################################################






'###################################################
'EXTRACT VERSION ###################################
'###################################################






'###################################################
'GO TO DOCUMENTATION ###############################
'###################################################






'###################################################
'IMPORT MODULE #####################################
'###################################################






'###################################################
'GET MODULE LIST ###################################
'###################################################






'###################################################
'GET VERSIONS ######################################
'###################################################






'###################################################
'GET GITHUB DATA ###################################
'###################################################






'###################################################
'IDENTIFY IMPORTED VERSION #########################
'###################################################






'###################################################
'DELETE FILE #######################################
'###################################################






'###################################################
'IMPORT MODULE #####################################
'###################################################

