Imports TwitterVB2

Module ModTwitterVB2
    Public tw As New TwitterVB2.TwitterAPI

    Sub Main()
        Dim strURL As String
        Dim strPIN As String

        strURL = GetLinkURL()
        Process.Start(strURL)
        Do
            Console.Write("Input PIN: ")
            strPIN = Console.ReadLine()
        Loop Until isValidPIN(strPIN)
        Console.WriteLine("This will never happen")
        Console.ReadKey()

    End Sub

    Function GetLinkURL() As String
        Dim tw1 As New TwitterAPI
        Return tw1.GetAuthorizationLink("consumerkey", "consumersecret")
    End Function

    Function isValidPIN(ByVal p_strPIN) As Boolean
        Dim tw2 As New TwitterAPI
        Return tw2.ValidatePIN(p_strPIN)
    End Function

End Module
