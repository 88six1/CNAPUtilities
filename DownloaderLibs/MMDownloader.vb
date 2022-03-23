Imports System.IO
Imports System.Net
Imports System.Threading
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Text
Imports DownloaderLibs
Imports System.Net.Mail
Imports CrystalDecisions.Shared
Imports Newtonsoft.Json
Imports ClosedXML.Excel
Imports DevExpress.XtraReports.UI
Imports DevExpress.Data.Utils

Public Class Provider
  Public provider_number As String
  Public provider_name As String
End Class
Public Class MMDownloader

  Public Shared connstr As String

  ''''Public Shared Function read_from_json_text_file(filename As String)

  ''''    Dim oconn As New SqlConnection("Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;")
  ''''    Dim ocomm As New SqlCommand("", oconn)
  ''''    oconn.Open()

  ''''    Dim json As String = File.ReadAllText(filename)

  ''''    Dim result = JsonConvert.DeserializeObject(json)

  ''''    Dim total_records As Integer = result("TotalCount")

  ''''    Dim myfile As System.IO.StreamWriter
  ''''    myfile = My.Computer.FileSystem.OpenTextFileWriter("C:\temp\providertempDF.txt", True)


  ''''    For i = 0 To total_records - 1
  ''''        Dim license As String = result("Data")(i)("CdcLicNbr")

  ''''        ocomm.CommandText = "Select count(*) from Providers where ProviderID = '" & license & "'"  ''''''''''''And 

  ''''        If ocomm.ExecuteScalar = 0 Then
  ''''            myfile.WriteLine(result("Data")(i)("CdcLicNbr") & " - " & result("Data")(i)("CdcLicName"))
  ''''        End If

  ''''    Next

  ''''    myfile.Close()
  ''''End Function   

  Public Shared Function DownloadFromCNAP(ByVal argconnstr As String) As Integer

    connstr = argconnstr
    ServicePointManager.Expect100Continue = True
    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    Dim req As WebRequest = WebRequest.Create("https://documents.apps.lara.state.mi.us/bchs/cdc.txt")
    ''http://www.cis.state.mi.us/fhs/brs/txt/cdc.txt")
    req.Timeout = 15000
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
      If str.StartsWith("D") Then
        providers.Add(str)
        linecount += 1
        counter += str.Length
      End If


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

    Dim newCenterRecordCounter = 0
    ''''''''''''''''''''''''''''''''''''''''''''
    oconn.Open()

    ''ocomm.CommandText = "Set NoCount ON; Insert into DownloadList (dtDownloaded,Merged) Values " &
    ''"('" & Now & "'," & 0 & "); Select IDENT_CURRENT('DownloadList') as expr1; CREATE TABLE #tmpProviders(" &
    ''  "[DCKEY] [int] IDENTITY(1,1) PRIMARY KEY NOT NULL ," &
    '' "[dlid] [int] NULL," &
    '' "[dtDownloaded] [datetime] NULL," &
    '' "[DCID] [varchar](14) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[LicenseBegin] [datetime] NULL," &
    '' "[LicenseEnd] [datetime] NULL," &
    '' "[Type] [varchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[Name] [varchar](150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[Address] [varchar](150) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[City] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[State] [varchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[Zip] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[County] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[Phone] [varchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[Type2] [varchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL," &
    '' "[statusflags] [int] NULL)"

    ''Dim id As Integer = ocomm.ExecuteScalar


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
    Dim newCenterProviders As New List(Of String)
    '''LOOOPPPPP
    For Each line In providers

      If (line.Contains("LITTLE HOUSE")) Then

        Dim s As String
        s = "SDFDF"
      End If
      line = Replace(line, "\n", String.Empty)
      line = Replace(line, "\r", String.Empty)
      line = Replace(line, "\t", String.Empty)
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


      If InStr(items(0), "DC") = 0 Then   '' Do the non centers first
        ocomm.CommandText = "Select count(*) from Providers where ProviderID = @providerid" ''''''''''''And 

        ocomm.Parameters.Clear()
        ocomm.Parameters.AddWithValue("@providerid", items(0))
        ocomm.Parameters.AddWithValue("@phone", items(10))

        If ocomm.ExecuteScalar = 0 Then

          Dim result As Integer _
              = InsertNewRecord(Now, Now, False, items(0), items(1), items(2),
                                                      items(3), items(4), items(5) & " " & items(6), items(7),
                                                      items(8), items(9), items(12), items(10), oconn)

          If result <> 1 Then
            MMLogger.WriteToErrorLog("", "", "Error inserting record:")
          End If

          ocomm.CommandText = "Select * from CountyList where CountyID = @countyid"

          ocomm.Parameters.Clear()
          ocomm.Parameters.AddWithValue("@countyid", items(12))

          items(12) = ocomm.ExecuteScalar()

          newProviders.Add("License: " & items(0) & vbCrLf & "Name: " & items(4) &
                          vbCrLf & "County: " & items(12) & vbCrLf & "Address: " & items(5) & " " & items(6) &
                          vbCrLf & "Zip: " & items(9) & vbCrLf & "Phone: " & items(10) & vbCrLf & "Begin Date: " & items(1) & vbCrLf & vbCrLf)
          newrecordcounter += 1
        End If

      ElseIf InStr(items(0), "DC") = 1 Then            '' Do the centers
        '   ocomm.Parameters.Clear()
        ocomm.CommandText = "Select count(*) from CenterProviders where ProviderID = @providerid" ''''''''''''And 

        ocomm.Parameters.Clear()
        ocomm.Parameters.AddWithValue("@providerid", items(0))
        ocomm.Parameters.AddWithValue("@phone", items(10))

        If ocomm.ExecuteScalar = 0 Then

          Dim result As Integer _
              = InsertNewCenterRecord(Now, Now, False, items(0), items(1), items(2),
                                                      items(3), items(4), items(5) & " " & items(6), items(7),
                                                      items(8), items(9), items(12), items(10), oconn)

          If result <> 1 Then
            MMLogger.WriteToErrorLog("", "", "Error inserting record:")
          End If

          ocomm.CommandText = "Select * from CountyList where CountyID = @countyid"

          ocomm.Parameters.Clear()
          ocomm.Parameters.AddWithValue("@countyid", items(12))

          items(12) = ocomm.ExecuteScalar()

          newCenterProviders.Add("License: " & items(0) & vbCrLf & "Name: " & items(4) &
                          vbCrLf & "County: " & items(12) & vbCrLf & "Address: " & items(5) & " " & items(6) &
                          vbCrLf & "Zip: " & items(9) & vbCrLf & "Phone: " & items(10) & vbCrLf & "Begin Date: " & items(1) & vbCrLf & vbCrLf)
          newCenterRecordCounter += 1
        End If
      End If

      recordcount += 1

    Next


    MMLogger.AppendLog(Now & " " & "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<")
    MMLogger.AppendLog("Total Records Downloaded: " & recordcount)
    MMLogger.AppendLog("New Records Added: " & newrecordcounter)
    MMLogger.AppendLog("New Center Records Added: " & newCenterRecordCounter)


    Dim sFileName As String = CreateReport()

    If EmailRecordsThroughGmail(newProviders, sFileName) = True Then
      MMLogger.AppendLog("Mail Sent Successfully")
    End If

    Dim sCenterFileName As String = CreateCenterReport()

    If EmailCenterRecordsThroughGmail(newCenterProviders, sCenterFileName) = True Then
      MMLogger.AppendLog("Mail Sent Successfully")
    End If
    '    ocomm.CommandText = "Drop table #tmpProviders"
    ' ocomm.ExecuteNonQuery()
    oconn.Close()

    Return newrecordcounter
    'MMDownloader.DownloadFromCNAP()

  End Function


  Public Shared Function QueryCenterExcelSheet(sqlconnstr As String)

    Dim providerlist As List(Of String) = DownloadProviderList()

    '// Create connection string variable. Modify the "Data Source"
    '// parameter as appropriate for your environment.
    Dim sConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;" &
                                  "Data Source=c:\temp\site info FY19 Rachel 8.21.192.xls;" &
                                  "Extended Properties=Excel 8.0;"

    Dim connstr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""c:\temp\cacfp.xls"";Extended Properties=""Excel 12.0 Xml;HDR=YES"""

    Dim conn As New System.Data.OleDb.OleDbConnection(connstr)

    conn.Open()
    ' Dim mytablename As String = conn.GetSchema("Tables").Rows(0)("TABLE_NAME")
    conn.Close()
    Dim strsql As String = "SELECT * FROM [CACFP$]"

    Dim cmd As New System.Data.OleDb.OleDbCommand(strsql, conn)
    Dim ds As New DataSet
    Dim adapter As New System.Data.OleDb.OleDbDataAdapter(cmd)
    adapter.Fill(ds)

    Dim dt As DataTable = ds.Tables(0)


    Dim oconn = New SqlConnection(sqlconnstr)
    oconn.Open()

    Dim dupcount As Int16 = 0

    For Each provider In providerlist
      Dim items() As String = ReturnRecordDetail(provider)

      Dim query = From prov In dt.AsEnumerable()
                  Where prov.Item("Name").ToString.ToLower.Contains(items(4).ToLower) And
              prov.Item("Address1").ToString.ToLower.Contains(items(5).ToLower)
                  Select prov


      If query.Count > 1 Then
        dupcount += 1
      End If
      For Each pr In query
        Dim result As Integer _
                    = InsertNewCenterRecord(Now, Now, False, items(0), items(1), items(2),
                                      items(3), items(4), items(5) & " " & items(6), items(7),
                                      items(8), items(9), items(12), items(10), oconn)
      Next


    Next

    MsgBox(dupcount)

  End Function

  Private Shared Function DownloadProviderList() As List(Of String)

    ServicePointManager.Expect100Continue = True
    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    Dim req As WebRequest = WebRequest.Create("https://documents.apps.lara.state.mi.us/bchs/cdc.txt")
    ''http://www.cis.state.mi.us/fhs/brs/txt/cdc.txt")
    req.Timeout = 15000
    Dim resp As WebResponse = req.GetResponse()
    Dim reader As New StreamReader(resp.GetResponseStream())
    Dim str As String = "" 'reader.ReadLine()
    Dim counter As Integer = 0
    Dim linecount As Integer = 0
    Dim providers As New List(Of String)

    While True
      str = reader.ReadLine()

      If str Is Nothing Then
        Exit While
      End If

      If str.Substring(0, 2) = "DC" Then
        providers.Add(str)

        linecount += 1
        counter += str.Length
      End If

    End While

    Return providers
  End Function

  Private Shared Function ReturnRecordDetail(line As String) As String()

    Const DATA_QUOTE_CHAR = "_dataquote_"
    Const DATA_DELIM = "_datadelim_"

    Dim bAbortOnError As Boolean
    Dim iFirstQuote As Integer
    Dim iSecondQuote As Integer

    Dim sField As String
    Dim sNewField As String

    '''LOOOPPPPP
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

    Return items
  End Function

  Private Shared Function InsertNewCenterRecord(
  ByVal dtAdded As Date,
  ByVal dtmodified As Date,
  ByVal reviewed As Boolean,
  ByVal providerid As String,
  ByVal licensebegin As String,
  ByVal licenseend As String,
  ByVal type As String,
  ByVal name As String,
  ByVal address As String,
  ByVal city As String,
  ByVal state As String,
  ByVal zip As String,
  ByVal county As String,
  ByVal phone As String,
  ByRef oconn As SqlConnection) As Integer

    If licensebegin <> "" Then
      licensebegin = CDate(licensebegin)
    End If

    If licenseend <> "" Then
      licenseend = CDate(licenseend)
    Else
      licenseend = CDate("01/01/1900")
    End If


    Dim ocomm As New SqlCommand("", oconn)
    With ocomm
      .CommandText = "Set NOCOUNT OFF;Insert into CenterProviders " &
                          "(dtAdded,dtmodified,reviewed,ProviderID,LicenseBegin,LicenseEnd,Type,Name,Address," &
                          "City,State,Zip,county,Phone) Values (@dtadded,@dtmodified,@reviewed,@providerid,@licensebegin,@licenseend," &
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
      mail.To.Add("jamwet@midmichigancc.com")
      mail.To.Add("helmcf@midmichigancc.com")
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

  Public Shared Function EmailRecordsThroughGmail(ByVal sProviders As List(Of String),
                                      ByVal sFileNameAttach As String) As Boolean

    Dim MyMailMessage As New MailMessage()
    MyMailMessage.From = New MailAddress("New Providers <midmichcc@gmail.com>")
    MyMailMessage.Bcc.Add("kevin@midmichigancc.com")
    MyMailMessage.Bcc.Add("rachel@midmichigancc.com")
    MyMailMessage.Bcc.Add("dongow@midmichigancc.com")


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
    SMTPServer.Credentials = New System.Net.NetworkCredential("midmichcc@gmail.com",
                                                              "lucy0000")
    SMTPServer.EnableSsl = True

    Try
      SMTPServer.Send(MyMailMessage)

      Return True
    Catch ex As SmtpException
      MMLogger.WriteToErrorLog(ex.Message, ex.StatusCode.ToString, "Error Sending Email")
      Return False
    End Try

  End Function

  Public Shared Function EmailCenterRecordsThroughGmail(ByVal sProviders As List(Of String),
                                      ByVal sFileNameAttach As String) As Boolean

    Dim MyMailMessage As New MailMessage()
    MyMailMessage.From = New MailAddress("New Centers <midmichcc@gmail.com>")
    MyMailMessage.Bcc.Add("kevin@midmichigancc.com")
    MyMailMessage.Bcc.Add("rachel@midmichigancc.com")
    MyMailMessage.Bcc.Add("dongow@midmichigancc.com")


    MyMailMessage.Subject = "New Centers for " & Now.ToString("M/d/yyyy")
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
      MyMailMessage.Body = "No new Centers today"
    End If

    Dim SMTPServer As New SmtpClient("smtp.gmail.com")
    SMTPServer.Port = 587
    SMTPServer.Credentials = New System.Net.NetworkCredential("midmichcc@gmail.com",
                                                              "lucy0000")
    SMTPServer.EnableSsl = True

    Try
      SMTPServer.Send(MyMailMessage)

      Return True
    Catch ex As SmtpException
      MMLogger.WriteToErrorLog(ex.Message, ex.StatusCode.ToString, "Error Sending Email")
      Return False
    End Try

  End Function

  Private Shared Function InsertNewRecord(
  ByVal dtAdded As Date,
  ByVal dtmodified As Date,
  ByVal reviewed As Boolean,
  ByVal providerid As String,
  ByVal licensebegin As String,
  ByVal licenseend As String,
  ByVal type As String,
  ByVal name As String,
  ByVal address As String,
  ByVal city As String,
  ByVal state As String,
  ByVal zip As String,
  ByVal county As String,
  ByVal phone As String,
  ByRef oconn As SqlConnection) As Integer

    If licensebegin <> "" Then
      licensebegin = CDate(licensebegin)
    End If

    If licenseend <> "" Then
      licenseend = CDate(licenseend)
    Else
      licenseend = CDate("01/01/1900")
    End If


    Dim ocomm As New SqlCommand("", oconn)
    With ocomm
      .CommandText = "Set NOCOUNT OFF;Insert into Providers " &
                          "(dtAdded,dtmodified,reviewed,ProviderID,LicenseBegin,LicenseEnd,Type,Name,Address," &
                          "City,State,Zip,county,Phone) Values (@dtadded,@dtmodified,@reviewed,@providerid,@licensebegin,@licenseend," &
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

  Public Shared Function CreateReport() As String

    '  Dim cfg As New MMConfig
    '    MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0,
    '   System.Reflection.Assembly.GetExecutingAssembly.Location.LastIndexOf("\") + 1) & "downloadercfg.txt")
    ' MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0, _
    '      System.Reflection.Assembly.GetExecutingAssembly.Location.IndexOf("DownloaderTestClient")) & "DownloaderSVC\bin\Debug\downloadercfg.txt")
    Dim sDirectory As String = MMConfig.GetOption("ReportDirectory")
    Dim sFileName As String = Now.ToString("yyyy-MM-dd") & ".pdf"
    Dim oconn As New SqlConnection(connstr)
    oconn.Open()
    Dim ocomm As New SqlCommand("", oconn)
    ocomm.CommandText = "Select * from Providers left join Countylist on Providers.county = countylist.countyid where (DATEDIFF(d, Providers.dtModified, '" & Now.ToString("M/d/yyyy") & "') = 0) order by dtAdded Desc"
    Dim daRecruit As New SqlDataAdapter(ocomm)

    Dim myDS As New dsRecruitmentReport
    daRecruit.Fill(myDS, "Providers")
    '' Dim rpt As New rptDailyRecruitment
    '' rpt.Database.Tables("Providers").SetDataSource(myDS.Tables("providers"))
    'rpt.SetDataSource(myDS)


    Dim rpt2 As XtraReport = New ProviderListReport()
    rpt2.DataSource = myDS
    ''   Dim printTool As New ReportPrintTool(rpt2)

    ''   printTool.ShowPreview()
    ''  rpt2.DataSource = myDS
    rpt2.ExportToPdf(sDirectory & sFileName)




    'CrystalReportViewer1.ReportSource = rpt
    'CrystalReportViewer1.Show()
    ''   Dim myDiskFileDestinationOptions As New DiskFileDestinationOptions()

    '  rpt.
    ''   rpt.ExportToDisk(ExportFormatType.PortableDocFormat, sDirectory & sFileName)

    Return sDirectory & sFileName
  End Function


  Public Shared Function CreateCenterReport() As String

    '  Dim connstr = "Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;"

    ' Dim cfg As New MMConfig
    ' MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0,
    ' System.Reflection.Assembly.GetExecutingAssembly.Location.LastIndexOf("\") + 1) & "downloadercfg.txt")
    ' MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0, _
    '      System.Reflection.Assembly.GetExecutingAssembly.Location.IndexOf("DownloaderTestClient")) & "DownloaderSVC\bin\Debug\downloadercfg.txt")
    Dim sDirectory As String = MMConfig.GetOption("ReportDirectory")
    Dim sFileName As String = "NewCenters_" & Now.ToString("yyyy-MM-dd") & ".xlsx"
    Dim oconn As New SqlConnection(connstr)
    oconn.Open()
    Dim ocomm As New SqlCommand("", oconn)
    ocomm.CommandText = "Select * from CenterProviders left join Countylist on CenterProviders.county = countylist.countyid where (DATEDIFF(d, CenterProviders.dtModified, '" & Now.ToString("M/d/yyyy") & "') = 0) order by dtAdded Desc"

    Dim dr As SqlDataReader = ocomm.ExecuteReader

    Dim wb = New XLWorkbook

    Dim ws = wb.AddWorksheet("centerProviders")
    Dim currentRow As Integer = 1

    ws.Cell(currentRow, 1).Value = "LicenseBegin"
    ws.Cell(currentRow, 2).Value = "LicenseEnd"
    ws.Cell(currentRow, 3).Value = "ProviderId"
    ws.Cell(currentRow, 4).Value = "Name"
    ws.Cell(currentRow, 5).Value = "Address"
    ws.Cell(currentRow, 6).Value = "City"
    ws.Cell(currentRow, 7).Value = "State"
    ws.Cell(currentRow, 8).Value = "Zip"
    ws.Cell(currentRow, 9).Value = "County"
    ws.Cell(currentRow, 10).Value = "Phone"

    For i = 1 To 10
      ws.Cell(currentRow, i).Style.Font.Bold = True
      ws.Cell(currentRow, i).Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent1, 0.5)
      ws.Cell(currentRow, i).Style.Font.FontColor = XLColor.White
    Next

    currentRow += 1

    While dr.Read()
      ws.Cell(currentRow, 1).Value = dr.Item("LicenseBegin")
      ws.Cell(currentRow, 2).Value = dr.Item("LicenseEnd")
      ws.Cell(currentRow, 3).Value = dr.Item("ProviderId")
      ws.Cell(currentRow, 4).Value = dr.Item("Name")
      ws.Cell(currentRow, 5).Value = dr.Item("Address")
      ws.Cell(currentRow, 6).Value = dr.Item("City")
      ws.Cell(currentRow, 7).Value = dr.Item("State")
      ws.Cell(currentRow, 8).Value = dr.Item("Zip")
      ws.Cell(currentRow, 9).Value = dr.Item("CountyName")
      ws.Cell(currentRow, 10).Value = dr.Item("Phone")
      currentRow += 1
    End While

    '    sFileName = Now.ToString("NewCenters_yyyy-MM-dd") & ".xlsx"
    wb.SaveAs(sDirectory & sFileName)

    Return sDirectory & sFileName
  End Function
End Class
