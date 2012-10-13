'---------------------------------------
'Programmed by: David Austria
'Website: http://www.exceltweet.com
'---------------------------------------

'--------------------------------------------------------------------------------
'This source code has a tutorial at my website for the step by step explanation
'on how to connect to a database and make changes like add/update/delete.
'--------------------------------------------------------------------------------

Imports System.Data.OleDb

Public Class frmTweetsList
    Dim sSql As String

    Private Sub frmTweetsList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        On Error GoTo C
        Call FillCombos()
        Me.cmbStatus.Text = "Programmed"
        Call FillList()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub lvList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvList.DoubleClick
        On Error GoTo C
        'Dim ID As String
        Dim ID As String = ""

        For Each sItem As ListViewItem In lvList.SelectedItems
            ID = sItem.Text
        Next

        With frmTweets
            .state = gModule.FormState.adStateEditMode
            .ID = ID

            .ShowDialog()

            Call FillList()
        End With

        frmTweets = Nothing
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub frmTweetsList_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        On Error GoTo C
        If Me.WindowState <> FormWindowState.Minimized Then
            'If Me.Width < 550 Then Me.Width = 550
            If Me.Height < 250 Then Me.Height = 150
            lvList.Height = Me.Height - 150
            'lvList.Width = Me.Width - 10
        End If
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        On Error GoTo C
        'Dim ID As String
        Dim Tweets As New frmTweets

        Tweets.state = gModule.FormState.adStateAddMode

        'For Each sItem As ListViewItem In lvList.SelectedItems
        '    ID = sItem.Text
        'Next

        'frmTweets.ID = ID
        Tweets.ShowDialog()

        Call FillList()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        On Error GoTo C
        Dim ID As String = ""

        For Each sItem As ListViewItem In lvList.SelectedItems
            ID = sItem.Text
        Next

        If ID <> "" Then
            'Delete the selected record
            Dim strDeleted As String

            strDeleted = ExecNonQuery("DELETE Tweets.ID FROM Tweets WHERE ID= " & ID & "")

            If strDeleted = "True" Then
                MsgBox("Record's deleted.", MsgBoxStyle.Information)

                Call FillList()
            Else
                MsgBox(strDeleted)
            End If
        Else
            MsgBox("Please select record to delete.", MsgBoxStyle.Critical)
        End If
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub FillList()
        On Error GoTo C
        sSql = "SELECT ID, Tweet, Time, UserName, Status FROM Tweets WHERE (1=1) "

        If Me.cmbStatus.Text <> "" Then
            sSql = sSql & " AND Status = '" & Me.cmbStatus.Text & "' "
        End If

        If Me.cmbUserName.Text <> "" Then
            sSql = sSql & " AND UserName = '" & Me.cmbUserName.Text & "' "
        End If

        sSql = sSql & " ORDER BY ID ASC"

        With lvList
            .Clear()

            .View = View.Details
            .FullRowSelect = True
            .GridLines = True
            .Columns.Add("ID", 50)
            .Columns.Add("Tweet", 200)
            .Columns.Add("Time", 150)
            .Columns.Add("UserName", 100)
            .Columns.Add("Status", 100)

            FillListView(lvList, GetData(sSql))
        End With
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub lvList_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lvList.SelectedIndexChanged

    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbStatus.SelectedIndexChanged
        Call FillList()
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

    Private Sub cmbUserName_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbUserName.SelectedIndexChanged
        Call FillList()
    End Sub

    Private Sub btnSearch_Click(sender As System.Object, e As System.EventArgs) Handles btnSearch.Click
        Call FillList()
    End Sub
End Class