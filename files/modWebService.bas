Option Explicit

' URL to call
Const URL = "http://ec.europa.eu/taxation_customs/vies/services/checkVatService"

' XML to send to the web service method
Const InputXmlFile = "C:\temp\checkVat.xml"

' *************************************************************
'
' Entry point
'
'    - Call the web service checkVAT method
'    - Upload XML data (country and VAT number)
'    - Get XML response
'    - Open the response as a workbook
'
' *************************************************************
Sub run()

    Dim sData As String
    Dim sResponseFileName As String

    ' Get the input manifest
    sData = openCheckVatXml(InputXmlFile)

    ' Consume the web service and get a filename with the response
    If (sData = "") Then
        MsgBox "Failure, the " & InputXmlFile & " file didn't exists", vbExclamation + vbOKOnly
        Exit Sub
    End If

    sResponseFileName = consumeWebService(URL, sData)

    ' Open the response as a workbook
    Call Application.Workbooks.OpenXML(Filename:=sResponseFileName)

End Sub

' *************************************************************
'
' Open the checkVat.xml input and replace variables
'
' *************************************************************
Private Function openCheckVatXml(ByVal sFileName As String) As String

    Dim sData As String

    sData = readFile(sFileName)

    If (sData <> "") Then
        sData = Replace(sData, "%COUNTRY%", "BE")
        sData = Replace(sData, "%VATNUMBER%", "0403170701") ' ENGIE Electrabel Belgique
    End If

    openCheckVatXml = sData

End Function

' *************************************************************
'
' Generic file reader. Return the content of the text file
'
' *************************************************************
Private Function readFile(ByVal sFileName As String) As String

    Dim objFso As Object
    Dim objFile As Object
    Dim sContent As String

    Set objFso = CreateObject("Scripting.FileSystemObject")

    If Not (objFso.FileExists(sFileName)) Then
        ' The file didn't exists
        readFile = ""
        Exit Function
    End If

    Set objFile = objFso.OpenTextFile(sFileName, 1)

    sContent = objFile.readAll

    objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing

    readFile = sContent

End Function

' *************************************************************
'
' Return a filename with the response of the web service method
'
' *************************************************************
Private Function consumeWebService(ByVal sURL As String, ByVal sData As String) As String

    Dim xmlhttp As Object
    Dim sResponseFileName As String

    Set xmlhttp = New MSXML2.ServerXMLHTTP60  ' Requires Microsoft XML, v6.0

    xmlhttp.Open "POST", sURL, True
    xmlhttp.send sData
    xmlhttp.waitForResponse

    sResponseFileName = createXmlTempFile(xmlhttp.responseText)

    Set xmlhttp = Nothing

    consumeWebService = sResponseFileName

End Function

' *************************************************************
'
' Create a temporary file in the TEMP folder and write in that
' file the XML response received by the web service.
'
' Return the temporary filename as result of this function
'
' *************************************************************
Private Function createXmlTempFile(ByVal sContent As String) As String

    Dim objFso As Object
    Dim objFile As Object
    Dim objFolder As Object
    Dim sFileName As String

    Set objFso = CreateObject("Scripting.FileSystemObject")

    ' 2 = temporary folder
    Set objFolder = objFso.GetSpecialFolder(2)
    sFileName = objFolder & "\"
    Set objFolder = Nothing

    sFileName = sFileName & objFso.GetTempName()
    sFileName = Replace(sFileName, ".tmp", ".xml")

    Set objFile = objFso.CreateTextFile(sFileName)

    objFile.Write sContent

    objFile.Close

    Set objFile = Nothing
    Set objFso = Nothing

    createXmlTempFile = sFileName

End Function
