Imports TwitterVB2
Imports System.Data.OleDb

'Dim tw As New TwitterVB2.TwitterAPI
'http://www.brianpautsch.com/blog/2009/5/1/twitterizer-simplifies-net-integration-with-twitter/

Public Class frmTwitter


    Dim daTweets As New OleDbDataAdapter
    Dim dsTweets As New DataSet

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Close()
    End Sub

    Private Sub cmdSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSend.Click
        On Error GoTo C
        If Len(Me.txtMessage.Text) = 0 Then
            MsgBox("Write your  tweet.", MsgBoxStyle.Information, "ExcelTweet")
            Exit Sub
        End If
        If Len(Me.txtMessage.Text) > 140 Then
            MsgBox("Your tweet was over 140 characters. You'll have to be more clever.", MsgBoxStyle.Information, "ExcelTweet")
            Exit Sub
        End If

        'tw.AuthenticateWith(strConsumerKey, strConsumerKeySecret, strOAuthToken, strOAuthTokenSecret)
        'tw.Update("¡Hola Mundo!")
        FuncAuthenticateUser(Me.cmbUserName.Text)


        If Me.txtInReplayToID.Text = "" Then
            tw.Update(Me.txtMessage.Text)
            Me.txtMessage.Text = ""
        Else
            tw.ReplyToUpdate(Me.txtMessage.Text, Me.txtInReplayToID.Text)
            Me.txtMessage.Text = ""
        End If

        'Me.lblResultado.Text = "Mensaje enviado: " & Me.txtMensaje.Text & "[" & Now() & "]"

        Call TestAPI()
        Call subLastTweet()
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub TestAPI()
        On Error Resume Next
        Dim a = tw.RateLimit_RemainingHits
        Dim b = tw.RateLimit_HourlyLimit

        Me.lblAPI.Text = "API (" & b & "-" & b - a & "=" & a & ")"
    End Sub

    Private Sub subLastTweet()
        On Error GoTo C
        Me.lblResultado.Text = "Last Tweet: " & _
            tw.AccountInformation.Status.Text & " - " & tw.AccountInformation.Status.CreatedAtLocalTime
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub txtMessage_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMessage.TextChanged
        On Error Resume Next
        Me.lblCaracteres.Text = "140-" & Len(Me.txtMessage.Text) & "=" & 140 - Len(Me.txtMessage.Text)
    End Sub

    Private Sub frmTwitter_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error Resume Next
        frmLogin.Close()
    End Sub

    Private Sub frmTwitter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        On Error GoTo C
        Me.cmdSend.Enabled = False
        Me.cmdExcel.Enabled = False
        Me.btnAdd.Enabled = False

        'Me.lblMessage.Text = cnString

        Call FillCombos()

        If Me.cmbUserName.Items.Count > 0 Then
            Me.cmbUserName.Text = Me.cmbUserName.Items(0).ToString
            'Call SubSelectUserName()

        End If



S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub cmdExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExcel.Click
        On Error GoTo C
        Dim m_Excel As Object ' New excel.Application
        Dim objLibroExcel As Object ' excel.Workbook
        Dim objHojaExcel As Object 'excel.Worksheet
        Dim i As Integer = 2
        'Dim twtMention As TwitterVB2.TwitterStatus
        Dim twtSearch As TwitterVB2.TwitterSearchResult

        Dim intTweets As Integer
        Dim intPages As Integer
        Dim tsp As TwitterSearchParameters
        Dim intPage As Integer
        Dim tp As TwitterParameters
        Dim strUser As String
        Dim intSinceID As String

        strUser = tw.AccountInformation.ScreenName & ""
        intSinceID = ""

        'm_Excel = New Excel.Application
        m_Excel = CreateObject("Excel.Application") ' New excel.Application
        ' m_Excel.Visible = True 
        Me.lblProcesing.Text = "Create Excel"
        intTweets = Me.txtTweets.Text

        'Dim ci As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")

        'objLibroExcel.GetType().InvokeMember("Add", Reflection.BindingFlags.InvokeMethod, Nothing, objLibroExcel, Nothing, ci)

        Dim oldCI As System.Globalization.CultureInfo
        oldCI = System.Threading.Thread.CurrentThread.CurrentCulture

        'System.Threading.Thread.CurrentThread.CurrentCulture = _
        '    New System.Globalization.CultureInfo("en-US")
        Call subCambiarIdioma()

        objLibroExcel = m_Excel.Workbooks.Add()

        System.Threading.Thread.CurrentThread.CurrentCulture = oldCI

        'objHojaExcel.Visible = excel.XlSheetVisibility.xlSheetVisible
        'AuthorUrl, Content, CreatedAt, ProfileImageUrl
        'objHojaExcel.Visible = True
        objHojaExcel = objLibroExcel.Worksheets(1)

        Me.lblProcesing.Text = "Create Workbook"
        With objHojaExcel
            '.Name = "Search " & strUser
            .Cells(1, 1) = "Tweet"
            .Cells(1, 2) = "ID"
            .Cells(1, 3) = "AuthorName"
            .Cells(1, 4) = "CreatedAtLocalTime"
            .Cells(1, 5) = "Title"
            .Cells(1, 6) = "Source"
            .Cells(1, 7) = "StatusUrl"
        End With

        Me.lblProcesing.Text = "Procesig (0)"
        'Dim tp As TwitterParameters
        'tp.Add(TwitterParameterNames.Count, 50)

        '' ''http://twittervb.codeplex.com/discussions/223841
        '' ''{SOLVED}Get newest tweets since X

        intPages = CInt(intTweets / 100)

        If (intTweets Mod 100) > 0 Then
            intPages = intPages + 1
        End If

        Select Case Me.cmbImportaOptions.Text
            Case "Search"
                'intPage = 1
                For intPage = 1 To intPages
                    tsp = New TwitterSearchParameters
                    tsp.Add(TwitterSearchParameterNames.SearchTerm, Me.txtSeach.Text)
                    tsp.Add(TwitterSearchParameterNames.Rpp, 100)

                    'If intSinceID <> "" Then
                    'tsp.Add(TwitterSearchParameterNames.Since, intSinceID)
                    'End If

                    tsp.Add(TwitterSearchParameterNames.Page, intPage)

                    'tsp.Add(TwitterSearchParameterNames.Lang,  intPage)

                    If tw.Search(tsp).Count > 0 Then

                        For Each twtSearch In tw.Search(tsp)
                            objHojaExcel.Cells(i, 1) = i - 1
                            intSinceID = twtSearch.ID
                            objHojaExcel.Cells(i, 2) = intSinceID
                            objHojaExcel.Cells(i, 3) = twtSearch.AuthorName
                            objHojaExcel.Cells(i, 4) = twtSearch.CreatedAtLocalTime
                            objHojaExcel.Cells(i, 5) = fnTextoCelda(twtSearch.Title)
                            'objHojaExcel.Cells(i, 3) = twtSearch.AuthorUrl
                            'objHojaExcel.Cells(i, 4) = twtSearch.Content
                            'objHojaExcel.Cells(i, 5) = twtSearch.CreatedAt

                            'objHojaExcel.Cells(i, 7) = twtSearch.ProfileImageUrl
                            objHojaExcel.Cells(i, 6) = HTMLClean(twtSearch.Source)
                            objHojaExcel.Cells(i, 7) = twtSearch.StatusUrl

                            Me.lblProcesing.Text = "Procesig (" & i - 1 & ") tweets."
                            Call TestAPI()
                            i = i + 1
                        Next
                    Else
                        Exit For
                    End If
                Next
            Case "Mentions"
                intPages = intPages * 2
                For intPage = 1 To intPages
                    tp = New TwitterParameters
                    'ScreenName 
                    tp.Add(TwitterParameterNames.Count, 50)
                    tp.Add(TwitterParameterNames.Page, intPage)

                    If tw.Mentions(tp).Count > 0 Then
                        For Each twtMention In tw.Mentions(tp)
                            objHojaExcel.Cells(i, 1) = i
                            objHojaExcel.Cells(i, 2) = twtMention.ID
                            'twtMention.User.Name() twtMention.CreatedAt twtMention.GeoLat twtMention.GeoLong twtMention.InReplyToScreenName
                            objHojaExcel.Cells(i, 3) = twtMention.User.ScreenName
                            objHojaExcel.Cells(i, 4) = twtMention.CreatedAtLocalTime
                            objHojaExcel.Cells(i, 5) = fnTextoCelda(twtMention.Text)
                            objHojaExcel.Cells(i, 6) = HTMLClean(twtMention.Source)
                            Me.lblProcesing.Text = "Procesig (" & i - 1 & ") tweets."
                            Call TestAPI()
                            i = i + 1
                        Next
                    Else
                        Exit For
                    End If
                Next
        End Select

        'On Error GoTo C

        Me.lblProcesing.Text = "..."
S:
        On Error Resume Next
        m_Excel.Visible = True

        objLibroExcel = Nothing
        objHojaExcel = Nothing
        m_Excel = Nothing
        Call TestAPI()

        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub subCambiarIdioma()
        On Error Resume Next
        System.Threading.Thread.CurrentThread.CurrentCulture = _
    New System.Globalization.CultureInfo("en-US")
    End Sub


    Public Sub New()
        ' Llamada necesaria para el diseñador.
        InitializeComponent()
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
    End Sub

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        frmTweetsList.Show()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        On Error Resume Next
        'Dim strTime As String
        Dim date1 As Date
        'strTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
        date1 = DateTime.Now

        'Me.ToolStripStatusLabel1.Text = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")
        Me.ToolStripStatusLabel1.Text = date1.ToString("dd/MM/yyyy HH:mm:ss")

        'If strTime & "" <> "" Then

        If date1.Second = 0 Or date1.Second = 30 Then
            'If Microsoft.VisualBasic.Right(strTime, 2) = "00" Then
            Me.ToolStripStatusLabel2.Text = "Procesar: " & date1.ToString("dd/MM/yyyy HH:mm:ss")
            Call SubFindTweets(date1)
        End If
        'End If

    End Sub


    Private Sub SubFindTweets(dateTweet As Date)
        On Error GoTo C
        Dim strSQL As String
        'Dim cnTweets As OleDbConnection
        'Dim cmTweets As OleDbCommand
        Dim strID As String
        Dim strTweet As String
        Dim strUserName As String
        Dim strUpdate As String

        Dim rdTweets As OleDbDataReader

        'Programmed, Published
        'strSQL = "SELECT ID, UserName, Tweet, Date, Time, Status FROM Tweets WHERE ((Time<=#" & strTime & "#) AND (Tweets.Status <> 'Published'))"
        'strSQL = "SELECT * FROM Tweets WHERE ((Time<=#" & dateTweet.ToString("yyyy/MM/dd HH:mm:ss") & "#) " & _
        '    "AND NOT (Tweets.Status IN('Published','Error')))"
        strSQL = "SELECT * FROM Tweets WHERE ((Time<=#" & dateTweet.ToString("yyyy/MM/dd HH:mm:ss") & "#) " & _
            "AND NOT (Tweets.Status IN('Published')))"
        rdTweets = GetData(strSQL)

        If rdTweets.HasRows = True Then
            While rdTweets.Read
                strID = rdTweets.Item("ID") & ""
                strTweet = rdTweets.Item("Tweet") & ""
                strUserName = rdTweets.Item("UserName") & ""


                If FuncAuthenticateUser(strUserName) = True Then

                    On Error Resume Next
                    tw.Update(strTweet)

                    If Err.Number = 0 Then
                        strUpdate = ExecNonQuery("UPDATE Tweets SET Status = 'Published' WHERE ID = " & strID & "")
                    Else
                        strUpdate = ExecNonQuery("UPDATE Tweets SET Status = 'Error' WHERE ID = " & strID & "")
                    End If

                Else
                    strUpdate = ExecNonQuery("UPDATE Tweets SET Status = 'Without User' WHERE ID = " & strID & "")
                End If

                'strUpdate = ExecNonQuery("UPDATE Tweets SET Status = 'Published' WHERE ID=" & strID & "")

                'If strUpdate <> "True" Then
                '    MsgBox(strUpdate)
                'End If

                System.Threading.Thread.Sleep(5000) ' Sleep for 2 seconds
                Me.ToolStripStatusLabel2.Text = "Tweet: " & strID & " published."

            End While
        End If
        rdTweets.Close()

S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub


    Public Sub FillCombos()
        On Error GoTo C
        Dim sSql As String

        If Me.cmbImportaOptions.Items.Count = 0 Then
            Me.cmbImportaOptions.Items.Add("Search")
            Me.cmbImportaOptions.Items.Add("Mentions")
        End If
        'Me.cmdImportaOptions'Search()'My(Tweets)'My(Timeline)'Mentions()'Direct(Message)'Tweets(from)'Followers()'Friends()

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
        Call SubSelectUserName()
        'Me.cmdSend.Enabled = True
    End Sub

    Private Sub SubSelectUserName()
        Dim strUserName As String
        strUserName = Me.cmbUserName.Text
        Call subAuthenticate()
    End Sub

    Private Function FuncAuthenticateUser(strUserName) As Boolean
        On Error GoTo C
        Dim sSql As String
        Dim myData As OleDbDataReader

        sSql = "SELECT UserName, OAuthToken, OAuthTokenSecret FROM UserNames " & _
            "WHERE UserName = '" & strUserName & "'"

        myData = GetData(sSql)

        Dim strValue As String

        If myData.HasRows = True Then
            Do While myData.Read
                strValue = IIf(myData.IsDBNull(0), "", myData.GetValue(0))
                strOAuthToken = IIf(myData.IsDBNull(1), "", myData.GetValue(1))
                strOAuthTokenSecret = IIf(myData.IsDBNull(2), "", myData.GetValue(2))
            Loop
            tw.AuthenticateWith(strConsumerKey, strConsumerKeySecret, strOAuthToken, strOAuthTokenSecret)
        Else
            Return False
        End If

        Return True
S:      Exit Function
C:      Return False
        Resume S
    End Function

    Private Sub subAuthenticate()
        On Error GoTo C

        FuncAuthenticateUser(Me.cmbUserName.Text)

        Me.cmdSend.Enabled = True
        Me.cmdExcel.Enabled = True
        Me.btnAdd.Enabled = True

        Me.Text = "ExcelTweet - " & tw.AccountInformation.ScreenName
        Me.txtMessage.Enabled = True
        Me.cmbImportaOptions.Enabled = True
        Me.txtSeach.Enabled = True
        Me.txtTweets.Enabled = True
        Me.txtInReplayToID.Enabled = True
        Call subLastTweet()
        Call TestAPI()

        Me.Timer1.Enabled = True

        'End If
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub

    Private Sub cmdAddUser_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddUser.Click
        frmWebBrowser.Show()
    End Sub

    Private Sub cmdDeleteUser_Click(sender As System.Object, e As System.EventArgs) Handles cmdDeleteUser.Click
        On Error GoTo C
        Dim strDeleted As String
        Dim strUserName As String

        strUserName = Me.cmbUserName.Text

        strDeleted = ExecNonQuery("DELETE * FROM UserNames WHERE UserName= '" & strUserName & "'")

        If strDeleted <> "True" Then
            MsgBox(strDeleted)
        End If

        Me.cmbUserName.Text = ""
        Me.lblResultado.Text = ""

        Me.cmdSend.Enabled = False
        Me.cmdExcel.Enabled = False
        Me.btnAdd.Enabled = False

        Call FillCombos()

        If Me.cmbUserName.Items.Count > 0 Then
            Me.cmbUserName.Text = Me.cmbUserName.Items(0).ToString

        End If
S:
        Exit Sub
C:      MsgBox(Err.Description, MsgBoxStyle.Information, "ExcelTweet")
        Resume S
    End Sub
End Class
