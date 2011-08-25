Imports System.Threading
Imports DownloaderLibs.Logger


''' pre post build 
'''F:\Projects\CareShare\CareShareService\CareShareService\bin\Debug\installutil.exe /u "F:\Projects\CareShare\CareShareService\CareShareService\bin\Debug\CareShareService.exe"
'''NET START CareShareService

Public Class DownloaderSvc

    Private timer1 As System.Threading.Timer

    ' Dim oOAuth As CSOAuth

    Protected Overrides Sub OnStart(ByVal args() As String)


        WriteToEventLog("Sleeping for 20 seconds", "CareShareService", EventLogEntryType.Error, "Application")
        System.Threading.Thread.Sleep(6000)
        Try

            ' Add code here to start your service. This method should set things
            ' in motion so your service can do its work.
            WriteToEventLog("Service Started", "CareShareService", EventLogEntryType.Error, "Application")


            'AppendLog("Service Started: " & Now.ToString)

            ''Dim objWriter As New System.IO.StreamWriter("c:\csservice.txt", True)
            ''objWriter.Write("Service Started: " & Now.ToString & vbCrLf)
            ''objWriter.Close()

            ''' READ Config File '''
            Dim cfg As New CSClientDAL.CSConfig
            CSConfig.Initialize(GetAppPath.Substring(0, _
                    GetAppPath.LastIndexOf("\")) & "\etc\servercfg")

            ''''see which CONFIG we are seeing
            '     WriteToErrorLog(CSConfig.ConfigFileName, GetAppPath.Substring(0, GetAppPath.LastIndexOf("\")) & "\etc\servercfg", "FileINFO")
            sCareShareUrl = CSConfig.GetOption("APIServer") '"http://127.0.0.1:3000"
            sGloDBIP = CSConfig.GetOption("GloDBIP")
            sGloDBName = CSConfig.GetOption("GloDBName")


            oOAuth = New CSOAuth

            With oOAuth
                .Access_Token = ""
                .Username = CSConfig.GetOption("oAuthUsername")
                .Client_Secret = CSConfig.GetOption("Client_Secret") '"55cbd097e3c65dc679f971ea0966daa800572ca8927b0275ecde940f9c1fd4a6&"
                .Password = CSConfig.GetOption("oAuthPass")
                .Grant_Type = CSConfig.GetOption("Grant_Type")
                .Client_ID = CSConfig.GetOption("Client_ID")
                .RequestAccessToken(sCareShareUrl)
            End With

            Dim oCallback As New TimerCallback(AddressOf OnTimedEvent)
            timer1 = New System.Threading.Timer(oCallback, Nothing, 5000, 60000 * 5)


        Catch ex As Exception
            ' WriteToEventLog(ex.StackTrace, "CareShareService", EventLogEntryType.Error, "Application")
            WriteToErrorLog(ex.Message, ex.StackTrace, "Error on Service Startup")
            Me.Stop()
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        timer1.Dispose()
    End Sub

    Private Sub OnTimedEvent(ByVal state As Object)

        Try

            Dim notices As String() = CSComm.CheckForNotices(sCareShareUrl, oOAuth.Access_Token)

            Dim i As Integer = 0
            If Not notices Is Nothing Then
                For i = 0 To notices.GetUpperBound(0)

                    Dim sPatientID As String = CSComm.HandleNotice(notices(i), sCareShareUrl, oOAuth.Access_Token, sGloDBIP, sGloDBName)

                    CSLogger.AppendLog("Glo PRI KEY: " & sPatientID)
                Next
            End If

            ' WriteToEventLog("TESTING" & Now.ToString, "CareShareService", EventLogEntryType.Information, "Application")
        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "Error on APIQueryEvent")
        End Try
    End Sub

End Class
