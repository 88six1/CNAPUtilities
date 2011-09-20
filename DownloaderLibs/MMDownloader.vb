Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports DownloaderLibs
Imports System.Net.Mail
Imports CrystalDecisions.Shared


Public Class MMDownloader

    Private Shared connstr As String
    Public Shared Sub DownloadFromCNAP(ByVal argconnstr As String)

        connstr = argconnstr

        Dim req As WebRequest = WebRequest.Create("http://www.cis.state.mi.us/fhs/brs/txt/cdc.txt")
        Dim resp As WebResponse = req.GetResponse()
        Dim reader As New StreamReader(resp.GetResponseStream())
        Dim str As String = reader.ReadLine()
        Dim counter As Integer = 0
        Dim linecount As Integer = 0
        Dim providers As New List(Of String)

        While True
            str = reader.ReadLine()

            If str Is Nothing Then
                Exit While
            End If

            providers.Add(str)

            linecount += 1
            counter += str.Length

        End While

        'ImportData(filename)

        Dim oconn As New SqlConnection
        Dim ocomm As New SqlCommand

        oconn = New SqlConnection(connstr)
        ocomm = New SqlCommand("", oconn)


        ''' SECOND SQL connection to see if record exists'''
        '  Dim compareconn As New SqlConnection(Globals.connstr)
        '  Dim comparecomm As New SqlCommand("Select count(*) from Providers where ProviderID = @providerid", compareconn)
        '  compareconn.Open()
        Dim newrecordcounter = 0
        ''''''''''''''''''''''''''''''''''''''''''''
        oconn.Open()

        ocomm.CommandText = "Set NoCount ON; Insert into DownloadList (dtDownloaded,Merged) Values " & _
        "('" & Now & "'," & 0 & "); Select IDENT_CURRENT('DownloadList') as expr1; CREATE TABLE #tmpProviders(" & _
          "[DCKEY] [int] IDENTITY(1,1) PRIMARY KEY NOT NULL ," & _
         "[dlid] [int] NULL," & _
         "[dtDownloaded] [datetime] NULL," & _
         "[DCID] [varchar](14) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[LicenseBegin] [datetime] NULL," & _
         "[LicenseEnd] [datetime] NULL," & _
         "[Type] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[Name] [varchar](150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[Address] [varchar](150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[City] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[State] [varchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[Zip] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[County] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[Phone] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[Type2] [varchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," & _
         "[statusflags] [int] NULL)"

        Dim id As Integer = ocomm.ExecuteScalar


        Dim recordcount As Integer = 0

        Const DATA_QUOTE_CHAR = "_dataquote_"
        Const DATA_DELIM = "_datadelim_"

        Dim bAbortOnError As Boolean
        Dim iFirstQuote As Integer
        Dim iSecondQuote As Integer

        Dim sField As String
        Dim sNewField As String

        Dim line As String
        Dim newProviders As New List(Of String)

        '''LOOOPPPPP
        For Each line In providers
            line = Replace(line, "," & """", "," & DATA_QUOTE_CHAR) 'data quote as first char in quoted field
            line = Replace(line, """" & ",", DATA_QUOTE_CHAR & ",") 'data quote as last char in quoted field
            line = Replace(line, """", DATA_QUOTE_CHAR) 'other data quote within field

            bAbortOnError = False

            Do Until InStr(1, line, "_dataquote_") = 0 Or bAbortOnError
                iFirstQuote = InStr(1, line, "_dataquote_")
                iSecondQuote = InStr(iFirstQuote + 1, line, "_dataquote_")

                If iFirstQuote = 0 Or iSecondQuote = 0 Then
                    bAbortOnError = True
                Else
                    sField = Mid(line, iFirstQuote, (iSecondQuote - iFirstQuote) + 11)
                    sNewField = Replace(sField, ",", DATA_DELIM) 'replace commas with placeholders
                    sNewField = Replace(sNewField, "_dataquote_", "")
                    line = Replace(line, sField, sNewField, , 1) 'replace first instance with placeholder version
                End If

            Loop

            Dim item As String
            Dim items() As String = Split(line, ",")

            Dim count As Integer = 0
            For Each item In items
                items(count) = Replace(items(count), "_datadelim_", ",")

                If items(count) Is Nothing Then
                    items(count) = ""
                End If
                count += 1
            Next


            ''If InStr(items(0), "DC") = 0 Then
            ''    With ocomm
            ''        .CommandText = "Insert into #tmpProviders " & _
            ''        "(DLID,dtdownloaded,DCID,LicenseBegin,LicenseEnd,Type,Name,Address," & _
            ''        "City,State,Zip,county,Phone) Values (@dlid,@dtdownloaded,@dcid,@licensebegin,@licenseend," & _
            ''        "@type,@name,@address,@city,@state,@zip,@county,@phone)"
            ''        .Parameters.Clear()
            ''        .Parameters.AddWithValue("@dlid", id)
            ''        .Parameters.AddWithValue("@dtdownloaded", Now.ToString("d"))
            ''        .Parameters.AddWithValue("@dcid", items(0))
            ''        .Parameters.AddWithValue("@licensebegin", items(1))
            ''        .Parameters.AddWithValue("@licenseend", items(2))
            ''        .Parameters.AddWithValue("@type", items(3))
            ''        .Parameters.AddWithValue("@name", items(4))
            ''        .Parameters.AddWithValue("@address", items(5) & " " & items(6))
            ''        .Parameters.AddWithValue("@city", items(7))
            ''        .Parameters.AddWithValue("@state", items(8))
            ''        .Parameters.AddWithValue("@zip", items(9))
            ''        .Parameters.AddWithValue("@county", items(12))
            ''        .Parameters.AddWithValue("@phone", items(10))
            ''        .ExecuteNonQuery()
            ''    End With
            ''End If

            ocomm.CommandText = "Select count(*) from Providers where ProviderID = @providerid"

            If InStr(items(0), "DC") = 0 Then
                ocomm.Parameters.Clear()
                ocomm.Parameters.AddWithValue("@providerid", items(0))

                If ocomm.ExecuteScalar = 0 Then

                    Dim result As Integer _
                        = InsertNewRecord(Now, Now, False, items(0), items(1), items(2), _
                                                                items(3), items(4), items(5) & " " & items(6), items(7), _
                                                                items(8), items(9), items(12), items(10), oconn)

                    If result <> 1 Then
                        MMLogger.WriteToErrorLog("", "", "Error inserting record:")
                    End If

                    ocomm.CommandText = "Select * from CountyList where CountyID = @countyid"

                    ocomm.Parameters.Clear()
                    ocomm.Parameters.AddWithValue("@countyid", items(12))

                    items(12) = ocomm.ExecuteScalar()

                    newProviders.Add("License: " & items(0) & vbCrLf & "Name: " & items(4) & _
                                    vbCrLf & "County: " & items(12) & vbCrLf & "Address: " & items(5) & " " & items(6) & _
                                    vbCrLf & "Zip: " & items(9) & vbCrLf & "Phone: " & items(10) & vbCrLf & "Begin Date: " & items(1) & vbCrLf & vbCrLf)
                    newrecordcounter += 1
                End If

            End If

            recordcount += 1

        Next


        MMLogger.AppendLog(Now & " " & "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<")
        MMLogger.AppendLog("Total Records Downloaded: " & recordcount)
        MMLogger.AppendLog("New Records Added: " & newrecordcounter)


        Dim sFileName As String = CreateReport()

        If EmailRecordsThroughGmail(newProviders, sfilename) = 1 Then
            MMLogger.AppendLog("Mail Sent Successfully")
        End If

        ocomm.CommandText = "Drop table #tmpProviders"
        ocomm.ExecuteNonQuery()
        oconn.Close()

        'MMDownloader.DownloadFromCNAP()

    End Sub

    Private Shared Function EmailRecords(ByVal newProviders As List(Of String)) As Integer
        Try
            Dim SmtpServer As New SmtpClient()
            Dim mail As New MailMessage()
            '          SmtpServer.Credentials = New  _
            'Net.NetworkCredential("username@gmail.com", "password")
            SmtpServer.Port = 25
            SmtpServer.Host = "mail.midmichigancc.com"
            mail = New MailMessage()
            mail.From = New MailAddress("kevin@midmichigancc.com")
            mail.To.Add("kimhow@midmichigancc.com")
            'mail.To.Add("kevin@merchsystems.com")
            mail.CC.Add("kevin@merchsystems.com")
            mail.Bcc.Add("cinpie@midmichigancc.com")
            mail.Subject = "New Providers for " & Now.Date
            Dim sProviders As String = ""
            If newProviders.Count > 0 Then
                For Each provider In newProviders
                    sProviders = sProviders & provider
                Next
                mail.Body = "Sorry for the formatting I couldn't get the pdf set up properly I'll fix the county number and consulant name in the next day or so" & vbCrLf & vbCrLf
                mail.Body = mail.Body & sProviders
            Else
                mail.Body = "No new providers today"
            End If

            SmtpServer.Send(mail)

        Catch ex As Exception
            Return 0
        End Try
        Return 1
    End Function

    Public Shared Function EmailRecordsThroughGmail(ByVal sProviders As List(Of String), _
                                        ByVal sFileNameAttach As String) As Boolean

        Dim MyMailMessage As New MailMessage()
        MyMailMessage.From = New MailAddress("NewProviders@midmichigancc.com")
        MyMailMessage.To.Add("kevin@merchsystems.com")


        MyMailMessage.Subject = "New Providers for " & Now.ToString("M/d/yyyy")
        Dim newProviders As String = ""

        If sProviders.Count > 0 Then
            For Each provider In sProviders
                newProviders = newProviders & provider
            Next
            MyMailMessage.Body = MyMailMessage.Body & newProviders
            If sFileNameAttach.Length > 0 Then
                MyMailMessage.Attachments.Add(New System.Net.Mail.Attachment(sFileNameAttach))
            End If
        Else
            MyMailMessage.Body = "No new providers today"
        End If

        Dim SMTPServer As New SmtpClient("smtp.gmail.com")
        SMTPServer.Port = 587
        SMTPServer.Credentials = New System.Net.NetworkCredential("caresharelogs@gmail.com", _
                                                                  "#careshare142")
        SMTPServer.EnableSsl = True

        Try
            SMTPServer.Send(MyMailMessage)

            Return (True)
        Catch ex As SmtpException
            MMLogger.WriteToErrorLog(ex.Message, ex.StatusCode.ToString, "Error Sending Email")
            Return False
        End Try


    End Function

    Private Shared Function InsertNewRecord( _
    ByVal dtAdded As Date, _
    ByVal dtmodified As Date, _
    ByVal reviewed As Boolean, _
    ByVal providerid As String, _
    ByVal licensebegin As Date, _
    ByVal licenseend As Date, _
    ByVal type As String, _
    ByVal name As String, _
    ByVal address As String, _
    ByVal city As String, _
    ByVal state As String, _
    ByVal zip As String, _
    ByVal county As String, _
    ByVal phone As String, _
    ByRef oconn As SqlConnection) As Integer

        Dim ocomm As New SqlCommand("", oconn)
        With ocomm
            .CommandText = "Set NOCOUNT OFF;Insert into Providers " & _
                                "(dtAdded,dtmodified,reviewed,ProviderID,LicenseBegin,LicenseEnd,Type,Name,Address," & _
                                "City,State,Zip,county,Phone) Values (@dtadded,@dtmodified,@reviewed,@providerid,@licensebegin,@licenseend," & _
                                "@type,@name,@address,@city,@state,@zip,@county,@phone)"
            .Parameters.Clear()
            .Parameters.AddWithValue("@dtadded", dtAdded)
            .Parameters.AddWithValue("@dtmodified", dtmodified)
            .Parameters.AddWithValue("@reviewed", reviewed)
            .Parameters.AddWithValue("@providerid", providerid)
            .Parameters.AddWithValue("@licensebegin", licensebegin)
            .Parameters.AddWithValue("@licenseend", licenseend)
            .Parameters.AddWithValue("@type", type)
            .Parameters.AddWithValue("@name", name)
            .Parameters.AddWithValue("@address", address)
            .Parameters.AddWithValue("@city", city)
            .Parameters.AddWithValue("@state", state)
            .Parameters.AddWithValue("@zip", zip)
            .Parameters.AddWithValue("@county", county)
            .Parameters.AddWithValue("@phone", phone)
            Return .ExecuteNonQuery()
        End With


        'DLID, dtdownloaded, DCID, LicenseBegin, LicenseEnd, Type, Name, Address, " & _"
        '"City,State,Zip,Phone) Values (@dlid,@dtdownloaded,@dcid,@licensebegin,@licenseend," & _
        '"@type,@name,@address,@city,@state,@zip,@phone
    End Function

    Private Shared Function CreateReport() As String

        Dim cfg As New MMConfig
        ' MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0, _
        '        System.Reflection.Assembly.GetExecutingAssembly.Location.LastIndexOf("\") + 1) & "downloadercfg.txt")
        MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0, _
               System.Reflection.Assembly.GetExecutingAssembly.Location.IndexOf("DownloaderTestClient")) & "DownloaderSVC\bin\Debug\downloadercfg.txt")

        Dim sDirectory As String = MMConfig.GetOption("ReportDirectory")
        Dim sFileName As String = Now.ToString("yyyy-MM-dd") & ".pdf"
        Dim oconn As New SqlConnection(connstr)
        oconn.Open()
        Dim ocomm As New SqlCommand("", oconn)
        ocomm.CommandText = "Select * from Providers left join Countylist on Providers.county = countylist.countyid where (DATEDIFF(d, Providers.dtModified, '" & Now.ToString("M/d/yyyy") & "') = 0) order by dtAdded Desc"
        Dim daRecruit As New SqlDataAdapter(ocomm)

        Dim myDS As New dsRecruitmentReport
        daRecruit.Fill(myDS, "Providers")
        Dim rpt As New rptDailyRecruitment
        rpt.Database.Tables("Providers").SetDataSource(myDS.Tables("providers"))
        'rpt.SetDataSource(myDS)

        'CrystalReportViewer1.ReportSource = rpt
        'CrystalReportViewer1.Show()
        Dim myDiskFileDestinationOptions As New DiskFileDestinationOptions()

        '  rpt.
        rpt.ExportToDisk(ExportFormatType.PortableDocFormat, sDirectory & sFileName)

        Return sDirectory & sFileName
    End Function
End Class
