'---------------------------------------
'Programmed by: David Austria
'Website: http://www.exceltweet.com
'---------------------------------------

Imports System.Data.OleDb

Public Class frmWebBrowser

    Private Sub cmdEnter_Click(sender As System.Object, e As System.EventArgs) Handles cmdEnter.Click
        On Error GoTo C
        Dim strPin As String
        Dim strSQL As String
        Dim strUserName As String

        strPin = Me.txtPIN.Text
        Dim IsValid As Boolean = tw.ValidatePIN(strPin)

        If IsValid Then
            strOAuthToken = tw.OAuth_Token()
            strOAuthTokenSecret = tw.OAuth_TokenSecret()
            strUserName = tw.AccountInformation.ScreenName

            subSaveUserName(strUserName, strOAuthToken, strOAuthTokenSecret)

        End If

        Me.Close()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub frmWebBrowser_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo C
        'frmTwitter.Close()
        frmTwitter.Show()
        frmTwitter.FillCombos()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub frmWebBrowser_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        On Error GoTo C

        Dim Url As String = tw.GetAuthorizationLink(strConsumerKey, strConsumerKeySecret)

        Me.WebBrowser1.Navigate(Url)

S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub


    Private Sub subSaveUserName(strUserName, strOAuthToken, strOAuthTokenSecret)
        On Error GoTo C
        Dim strSQL As String
        Dim strDeleted As String
        Dim strInsert As String

        strDeleted = ExecNonQuery("DELETE UserNames.UserName FROM UserNames WHERE UserName= '" & strUserName & "'")

        If strDeleted <> "True" Then
            MsgBox(strDeleted)
        End If

        strSQL = "INSERT INTO UserNames (UserName,OAuthToken,OAuthTokenSecret) " & _
            "VALUES ('" & strUserName & "', '" & strOAuthToken & "', '" & strOAuthTokenSecret & "')"

        strInsert = ExecNonQuery(strSQL)

        If strInsert <> "True" Then
            MsgBox(strInsert)
        End If
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub


End Class