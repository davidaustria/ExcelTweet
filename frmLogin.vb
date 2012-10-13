Public Class frmLogin

    Private Sub cmdLogin_Click(sender As System.Object, e As System.EventArgs) Handles cmdLogin.Click
        On Error GoTo C
        frmWebBrowser.Show()
        Me.Hide()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub PictureBox1_Click(sender As System.Object, e As System.EventArgs) Handles PictureBox1.Click
        Call OpenET()
    End Sub

    Private Sub frmLogin_Click(sender As Object, e As System.EventArgs) Handles Me.Click
        Call OpenET()
    End Sub

    Private Sub frmLogin_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        On Error GoTo C
        'Find UserNames
        ' if hasuserName --> frmTweets
        'else login

        System.Threading.Thread.Sleep(5000) ' Sleep for 2 seconds

        If HasUserName() = True Then
            Me.Timer1.Enabled = True
        Else
            Me.cmdLogin.Visible = True
        End If
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub


    Private Function HasUserName() As Boolean
        On Error GoTo C
        Dim sSql As String
        Dim intUserName As Integer

        sSql = "SELECT COUNT(*) FROM UserNames GROUP BY UserName"
        intUserName = ExecScalar(sSql)

        'With Me.cmbUserName.Items
        '    .Clear()
        '    FillComboBox(cmbUserName, GetData(sSql))

        '    If Me.cmbUserName.Items.Count > 0 Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'End With

        If intUserName > 0 Then
            Return True
        Else
            Return False
        End If

S:      Exit Function
C:      'MsgBox(Err.Description)
        Return False
        Resume S
    End Function


    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Call OpenET()
    End Sub


    Private Sub OpenET()
        On Error GoTo C
        Me.Timer1.Enabled = False
        frmTwitter.Show()
        Me.Hide()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub
End Class