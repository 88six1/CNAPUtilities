Imports System.IO

Public Class MMLogger

    Public Shared Sub AppendLog(ByVal sLogString As String)

        Try

            Dim sFileLoc As String = GetAppPath()
            sFileLoc = sFileLoc.Substring(0, sFileLoc.LastIndexOf("\"))
            If Not System.IO.Directory.Exists(sFileLoc & "\Logs\") Then
                System.IO.Directory.CreateDirectory(sFileLoc & "\Logs\")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(sFileLoc & "\Logs\applog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)
            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(sFileLoc & "\Logs\applog.txt", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            s1.Write(sLogString & vbCrLf)
            s1.Close()
            fs1.Close()

        Catch ex As Exception
            ''' write to event log
        End Try

    End Sub

    Public Shared Sub WriteToErrorLog(ByVal msg As String, _
       ByVal stkTrace As String, ByVal title As String)

        Try

            Dim sFileLoc As String = GetAppPath()
            sFileLoc = sFileLoc.Substring(0, sFileLoc.LastIndexOf("\"))

            If Not System.IO.Directory.Exists(sFileLoc & "\Logs\") Then
                System.IO.Directory.CreateDirectory(sFileLoc & "\Logs\")
            End If

            'check the file
            Dim fs As FileStream = New FileStream(sFileLoc & "\Logs\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
            Dim s As StreamWriter = New StreamWriter(fs)
            s.Close()
            fs.Close()

            'log it
            Dim fs1 As FileStream = New FileStream(sFileLoc & "\Logs\errlog.txt", FileMode.Append, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)
            s1.Write("Title: " & title & vbCrLf)
            s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
            s1.Write("Message: " & msg & vbCrLf)
            s1.Write("StackTrace: " & stkTrace & vbCrLf)
            s1.Write("================================================" & vbCrLf)
            s1.Close()
            fs1.Close()

        Catch ex As Exception
            ''' write to event log
        End Try
    End Sub


    ''' ' Future
    Public Shared Function WriteToEventLog(ByVal entry As String, _
                    Optional ByVal appName As String = "CompanyName", _
                    Optional ByVal eventType As  _
                    EventLogEntryType = EventLogEntryType.Information, _
                    Optional ByVal logName As String = "ProductName") As Boolean

        Dim objEventLog As New EventLog

        Try

            'Register the Application as an Event Source

            If Not EventLog.SourceExists(appName) Then
                EventLog.CreateEventSource(appName, logName)
            End If

            'log the entry

            objEventLog.Source = appName
            objEventLog.WriteEntry(entry, eventType)

            Return True

        Catch Ex As Exception

            Return False

        End Try

    End Function

    Public Shared Function GetAppPath()
        Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location())
    End Function

End Class
