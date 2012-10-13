'---------------------------------------
'Programmed by: David Austria
'Website: http://www.exceltweet.com
'---------------------------------------

Imports System.Data.OleDb
'Imports System.Data

Module gModule
    Public Enum FormState
        adStateAddMode = 0
        adStateEditMode = 1
    End Enum

    'Fill ListView control with data
    Public Sub FillListView(ByRef lvList As ListView, ByRef myData As OleDbDataReader)
        Dim itmListItem As ListViewItem
        Dim strValue As String
        Try
            Do While myData.Read
                itmListItem = New ListViewItem()
                strValue = IIf(myData.IsDBNull(0), "", myData.GetValue(0))
                itmListItem.Text = strValue

                For shtCntr = 1 To myData.FieldCount() - 1
                    If myData.IsDBNull(shtCntr) Then
                        itmListItem.SubItems.Add("")
                    Else
                        If IsDate(myData.GetValue(shtCntr)) Then
                            itmListItem.SubItems.Add(myData.GetValue(shtCntr).ToString)
                        Else
                            itmListItem.SubItems.Add(myData.GetString(shtCntr))
                        End If
                    End If
                Next shtCntr

                lvList.Items.Add(itmListItem)
            Loop
        Catch ex As OleDbException
            MsgBox(ex.ToString)
        End Try
    End Sub

    'Fill ListView control with data
    Public Sub FillComboBox(ByRef cbList As ComboBox, ByRef myData As OleDbDataReader)
        Try
            Dim strValue As String
            Do While myData.Read
                strValue = IIf(myData.IsDBNull(0), "", myData.GetValue(0))
                cbList.Items.Add(strValue & "")
            Loop
        Catch ex As OleDbException
            MsgBox(ex.ToString)
        End Try
    End Sub



    'Execute Non Query
    Public Function ExecNonQuery(ByVal strSQL As String)
        Dim cnHotel As OleDbConnection
        cnHotel = New OleDbConnection

        Try
            With cnHotel
                If .State = ConnectionState.Open Then .Close()

                .ConnectionString = FuncConexion()
                .Open()
            End With

            Dim cmd As OleDbCommand = New OleDbCommand(strSQL, cnHotel)

            cmd.ExecuteNonQuery()

            Return True
        Catch ex As OleDbException
            Return ex
        Finally
            cnHotel.Close()
        End Try
    End Function

    ' Referencia: http://www.elguille.info/NET/ADONET/cuando_usar_ExecuteNonQuery_o_ExecuteScalar.htm#ExecuteScalar_VB
    'ExecuteScalar  Non Query
    Public Function ExecScalar(ByVal strSQL As String)
        Dim cnHotel As OleDbConnection
        Dim n As Integer
        cnHotel = New OleDbConnection

        Try
            With cnHotel
                If .State = ConnectionState.Open Then .Close()

                .ConnectionString = FuncConexion()
                .Open()
            End With

            Dim cmd As OleDbCommand = New OleDbCommand(strSQL, cnHotel)

            'cmd.ExecuteNonQuery()

            n = CInt(cmd.ExecuteScalar())

            Return n
        Catch ex As OleDbException
            'Return ex
            Return 0
        Finally
            cnHotel.Close()
        End Try
    End Function

    Public Function GetData(ByVal sSQL As String) As OleDbDataReader
        On Error GoTo C
        Dim cnCustomers As OleDbConnection
        Dim sqlCmd As OleDbCommand = New OleDbCommand(sSQL)
        Dim myData As OleDbDataReader
        'Dim b As System.Data.CommandBehavior

        cnCustomers = New OleDbConnection(FuncConexion)

        'Try
        cnCustomers.Open()

        sqlCmd.Connection = cnCustomers

        'b = CommandBehavior.CloseConnection

        myData = sqlCmd.ExecuteReader()

        Return myData
        'Catch ex As Exception
        'Return ex
        'End Try
S:
        Exit Function
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Function


    Public Function HTMLClean(ByVal strText) As String
        On Error Resume Next
        Dim strContent As String, mString As String
        Dim mStartPos As Long, mEndPos As Long
        Dim i, j
        strContent = strText & ""
        ' Start process
        mStartPos = InStr(strContent, "<")
        mEndPos = InStr(strContent, ">")
        Do While mStartPos <> 0 And mEndPos <> 0 And mEndPos > mStartPos
            mString = Mid(strContent, mStartPos, mEndPos - mStartPos + 1)
            strContent = Replace(strContent, mString, "")
            mStartPos = InStr(strContent, "<")
            mEndPos = InStr(strContent, ">")
        Loop
        ' Translate common escape sequence chars
        ' ''strContent = Replace(strContent, "&nbsp;", " ")
        ' ''strContent = Replace(strContent, "&amp;", "&")
        ' ''strContent = Replace(strContent, "&quot;", "'")
        ' ''strContent = Replace(strContent, "&#", "#")
        ' ''strContent = Replace(strContent, "&lt;", "<")
        ' ''strContent = Replace(strContent, "&gt;", ">")
        ' ''strContent = Replace(strContent, "%20", " ")
        strContent = LTrim(Trim(strContent))
        Do While Left(strContent, 1) = Chr(13) Or Left(strContent, 1) = Chr(10)
            strContent = Mid(strContent, 2)
        Loop
        HTMLClean = strContent
    End Function

    Public Function fnTextoCelda(ByVal strText1) As String
        Dim strTexto2 As String
        Dim strInicio As String

        strTexto2 = strText1 & ""
        strInicio = Left(strTexto2, 1)
        If strInicio = "=" Or strInicio = "+" Or strInicio = "*" Then
            strTexto2 = "'" & strTexto2
        End If

        fnTextoCelda = strTexto2
    End Function


    'Function stripHTML(ByVal strHTML)
    '    'Strips the HTML tags from strHTML

    '    Dim strOutput
    '    Dim objRegExp = System.Text.RegularExpressions.Regex

    '    objRegExp.IgnoreCase = True
    '    objRegExp.Global = True
    '    objRegExp.Pattern = "<(.|\n)+?>"

    '    'Replace all HTML tag matches with the empty string
    '    strOutput = objRegExp.Replace(strHTML, "")

    '    'Replace all < and > with < and >
    '    strOutput = Replace(strOutput, "<", "<")
    '    strOutput = Replace(strOutput, ">", ">")

    '    stripHTML = strOutput 'Return the value of strOutput

    '    objRegExp = Nothing
    'End Function

End Module
