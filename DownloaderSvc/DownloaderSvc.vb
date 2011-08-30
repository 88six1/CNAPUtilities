Imports System.Threading
Imports DownloaderLibs.MMLogger
Imports DownloaderLibs.MMConfig
Imports DownloaderLibs
Imports System.Timers


''' pre post build 
'''F:\Projects\CareShare\CareShareService\CareShareService\bin\Debug\installutil.exe /u "F:\Projects\CareShare\CareShareService\CareShareService\bin\Debug\CareShareService.exe"
'''NET START CareShareService

Public Class DownloaderSvc

    Private timer1 As System.Threading.Timer

    ' Dim oOAuth As CSOAuth

    Protected Overrides Sub OnStart(ByVal args() As String)


        System.Threading.Thread.Sleep(6000)
        Try

            ' Add code here to start your service. This method should set things
            ' in motion so your service can do its work.
         

            'AppendLog("Service Started: " & Now.ToString)

            ''Dim objWriter As New System.IO.StreamWriter("c:\csservice.txt", True)
            ''objWriter.Write("Service Started: " & Now.ToString & vbCrLf)
            ''objWriter.Close()

            ''' READ Config File '''
            '     Dim cfg As New DownloaderLibs.MMConfig
            '    MMConfig.Initialize(GetAppPath.Substring(0, _
            '           GetAppPath.LastIndexOf("\")) & "\etc\servercfg")

            ''''see which CONFIG we are seeing
            '     WriteToErrorLog(CSConfig.ConfigFileName, GetAppPath.Substring(0, GetAppPath.LastIndexOf("\")) & "\etc\servercfg", "FileINFO")
            '   sCareShareUrl = Config.GetOption("APIServer")
            '  sGloDBIP = Config.GetOption("GloDBIP")
            ' sGloDBName = Config.GetOption("GloDBName")


            Dim oCallback As New TimerCallback(AddressOf OnTimedEvent)
            timer1 = New System.Threading.Timer(oCallback, Nothing, 5000, 60000)

            AppendLog("Service Started:::" & Now)

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "Error on Service Startup")
            Me.Stop()
        End Try
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        AppendLog("Service Stopped:::" & Now)
        timer1.Dispose()
    End Sub

    Private Sub OnTimedEvent(ByVal state As Object)

        Try
            Dim time As String = TimeString
            If time.Substring(0, time.LastIndexOf(":")) = "24:00" Then
                AppendLog("DownloaderStarted:::" & Now)
                MMDownloader.DownloadFromCNAP()
            End If

        Catch ex As Exception
            WriteToErrorLog(ex.Message, ex.StackTrace, "Error on APIQueryEvent")
        End Try
    End Sub

End Class
