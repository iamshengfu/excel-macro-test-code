Sub getxml()


Dim xmlhttp As New MSXML2.XMLHTTP60, myurl As String

myurl = "http://www.google.com.sg"

xmlhttp.Open "GET", myurl, False

xmlhttp.send

Debug.Print (xmlhttp.responseText)
End Sub
