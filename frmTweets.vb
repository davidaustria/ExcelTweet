'---------------------------------------
'Programmed by: David Austria
'Website: http://www.exceltweet.com
'---------------------------------------

Imports System.Data.OleDb

Public Class frmTweets
    Dim daTweets As New OleDbDataAdapter
    Dim dsTweets As New DataSet
    Public ID As Integer
    Public state As Integer    'Public State As gModule.FormState

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cnTweets As OleDbConnection
        cnTweets = New OleDbConnection

        Call FillCombos()
        Me.cmbStatus.Text = "Programmed"

        Try
            With cnTweets
                If .State = ConnectionState.Open Then .Close()

                .ConnectionString = FuncConexion()
                .Open()
            End With
        Catch ex As OleDbException
            MsgBox(ex.ToString)
        End Try

        'DTP_time.Value = System.DateTime.Now

        If state = gModule.FormState.adStateAddMode Then
            Try
                Dim qryTweets As String = "SELECT * FROM Tweets"

                daTweets.SelectCommand = New OleDbCommand(qryTweets, cnTweets)

                Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(daTweets)

                daTweets.Fill(dsTweets, "Tweets")
            Catch ex As OleDbException
                MsgBox(ex.ToString)
            Finally
                cnTweets.Close()
            End Try
        ElseIf state = gModule.FormState.adStateEditMode Then
            Dim qryTweets As String = "SELECT * FROM Tweets WHERE (ID = " & ID & ")"

            daTweets.SelectCommand = New OleDbCommand(qryTweets, cnTweets)

            Dim cb As OleDbCommandBuilder = New OleDbCommandBuilder(daTweets)

            daTweets.Fill(dsTweets, "Tweets")

            Dim dt As DataTable = dsTweets.Tables("Tweets")

            Try
                txtID.Text = dt.Rows(0)("ID")
                txtTweet.Text = dt.Rows(0)("Tweet")
                If IsDate(dt.Rows(0)("Time")) = True Then DTP_time.Value = dt.Rows(0)("Time")
                cmbUserName.Text = IIf(IsDBNull(dt.Rows(0)("UserName")), "", dt.Rows(0)("UserName"))
                cmbStatus.Text = IIf(IsDBNull(dt.Rows(0)("Status")), "", dt.Rows(0)("Status"))
            Catch ex As OleDbException
                MsgBox(ex.ToString)
            Finally
                cnTweets.Close()
            End Try
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim dt As DataTable = dsTweets.Tables("Tweets")
        Dim strSQL As String

        'If txtTweet.Text = "" Or txtDate.Text = "" Then
        If txtTweet.Text = "" Then
            MsgBox("Please fill up Tweet information.", MsgBoxStyle.Critical)
            Exit Sub
        End If

        If cmbUserName.Text = "" Then
            MsgBox("Select a user.", MsgBoxStyle.Critical)
            Exit Sub
        End If

        Dim cnTweets As OleDbConnection
        cnTweets = New OleDbConnection
        'Dim cmTweets As OleDbCommand


        Try
            With cnTweets
                If .State = ConnectionState.Open Then .Close()

                .ConnectionString = FuncConexion()
                .Open()
            End With
        Catch ex As OleDbException
            MsgBox(ex.ToString)
        End Try




        Try
            If state = gModule.FormState.adStateAddMode Then
                ' add a row


                Dim strInsert As String
                strSQL = "INSERT INTO Tweets ([Tweet],[Status],[Time],[UserName]) " & _
                    "VALUES ('" & Me.txtTweet.Text & "', '" & Me.cmbStatus.Text & "', '" & _
                    DTP_time.Value & "' , '" & cmbUserName.Text & "')"

                strInsert = ExecNonQuery(strSQL)

                If strInsert <> "True" Then
                    MsgBox(strInsert)
                End If



            Else
                'With dt
                '    '.Rows(0)("ID") = txtID.Text
                '    .Rows(0)("Tweet") = txtTweet.Text
                '    .Rows(0)("Date") = DTP_date.Value
                '    .Rows(0)("Time") = DTP_time.Value
                '    .Rows(0)("UserName") = IIf(cmbUserName.Text = "", System.DBNull.Value, cmbUserName.Text)
                '    .Rows(0)("Status") = IIf(cmbStatus.Text = "", System.DBNull.Value, cmbStatus.Text)
                'End With

                'daTweets.Update(dsTweets, "Tweets")

                Dim strInsert As String
                strSQL = "UPDATE Tweets SET [Tweet]='" & Me.txtTweet.Text & "',[Status]='" & _
                    Me.cmbStatus.Text & "', [Time]='" & DTP_time.Value & "', [UserName]='" & cmbUserName.Text & "' " & _
                    "WHERE (ID=" & txtID.Text & ") "

                '"VALUES ('" & Me.txtTweet.Text & "', '" & Me.cmbStatus.Text & "', '" & _
                'DTP_date.Value & "' , '" & DTP_time.Value & "' , '" & cmbUserName.Text & "')"

                strInsert = ExecNonQuery(strSQL)

                If strInsert <> "True" Then
                    MsgBox(strInsert)
                End If


            End If




            MsgBox("Record successfully saved.", MsgBoxStyle.Information)
        Catch ex As OleDbException
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub FillCombos()
        On Error GoTo C
        Dim sSql As String

        sSql = "SELECT DISTINCT Status FROM Tweets ORDER BY Status"

        With Me.cmbStatus.Items
            .Clear()
            FillComboBox(cmbStatus, GetData(sSql))
        End With

        sSql = "SELECT UserName FROM UserNames ORDER BY UserName ASC"

        With Me.cmbUserName.Items
            .Clear()
            FillComboBox(cmbUserName, GetData(sSql))
        End With
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

End Class