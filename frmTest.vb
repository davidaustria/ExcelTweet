Imports System.Text.RegularExpressions

Public Class frmTest

    Private Sub btnTest_Click(sender As System.Object, e As System.EventArgs) Handles btnTest.Click
        MsgBox(GetGeoCoords("Reforma 222", 1))
    End Sub


    Private Function GetGeoCoords(ByVal inString As String, ByVal inType As Integer) As String
        ' Explanation of function:
        ' Use inType=0 and feed in a specific Google Maps URL to parse out the GeoCoords from the URL
        ' e.g. http://maps.google.com/maps?f=q&source=s_q&hl=en&geocode=&q=53154&sll=37.0625,-95.677068&sspn=52.505328,80.507812&ie=UTF8&ll=42.858224,-88.000832&spn=0.047943,0.078621&t=h&z=14
        ' Function returns a string of geocoords (e.g. "-87.9010610,42.8864960")
        '
        ' Use inType=1 and feed in a zip code, address, or business name
        ' Function returns a string of geocoords (e.g. "-87.9010610,42.8864960")
        ' If an invalid address, zip code or location was entered, the function will return "0,0"

        Dim Chunks As String()
        Dim outString As String = ""

        If inType = 0 Then
            Chunks = Regex.Split(inString, "&")
            For Each s As String In Chunks
                If InStr(s, "ll") > 0 Then outString = s
            Next
            outString = Replace(Replace(outString, "sll=", ""), "ll=", "")
        Else
            Dim xmlString As String = GetHTML("http://maps.google.com/maps/geo?output=xml&key=abcdefg&q=" & inString, 1)
            Chunks = Regex.Split(xmlString, "coordinates>", RegexOptions.Multiline)
            If Chunks.Count > 1 Then
                outString = Replace(Chunks(1), ",0</", "")
            Else
                outString = "0,0"
            End If

        End If
        Return outString
    End Function

    Private Function GetHTML(ByVal sURL As String, ByVal e As Integer) As String
        Dim oHttpWebRequest As System.Net.HttpWebRequest
        Dim oStream As System.IO.Stream
        Dim sChunk As String
        oHttpWebRequest = (System.Net.HttpWebRequest.Create(sURL))
        Dim oHttpWebResponse As System.Net.WebResponse = oHttpWebRequest.GetResponse()
        oStream = oHttpWebResponse.GetResponseStream
        sChunk = New System.IO.StreamReader(oStream).ReadToEnd()
        oStream.Close()
        oHttpWebResponse.Close()
        If e = 0 Then
            'Return HttpUtility.HtmlEncode(sChunk)
        Else
            'Return HttpUtility.HtmlDecode(sChunk)
        End If
        ' Revisar, DAG 12.11.12
        GetHTML = sChunk
    End Function

End Class