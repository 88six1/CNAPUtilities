Imports DownloaderLibs
Imports System.Drawing.Imaging
Imports CrystalDecisions.Shared
Imports CrystalDecisions
Imports System.Data.SqlClient
Imports System.Net.Mail

Public Class Form1

    Private connstr As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connstr = "Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim MyMailMessage As New MailMessage()
        MyMailMessage.From = New MailAddress("Kevin Fisk <NewProviders@midmichigancc.com>")
        MyMailMessage.To.Add("kevin@merchsystems.com")

        MyMailMessage.Subject = "New Providers for " & Now.ToString("M/d/yyyy")
        Dim newProviders As String = ""

        MyMailMessage.Body = "testing testing"
         
        Dim SMTPServer As New SmtpClient("smtp.gmail.com")
        SMTPServer.Port = 587
        SMTPServer.Credentials = New System.Net.NetworkCredential("caresharelogs@gmail.com", _
                                                                  "#careshare142")
        SMTPServer.EnableSsl = True

        SMTPServer.Send(MyMailMessage)


        'MessageBox.Show(TimeString)


        ' ''Dim cfg As New MMConfig
        'MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0, _
        '       System.Reflection.Assembly.GetExecutingAssembly.Location.IndexOf("DownloaderTestClient")) & "DownloaderSVC\bin\Debug\downloadercfg.txt")

        ' ''Dim stime As String = MMConfig.GetOption("DownloadTime")


        ' ''Dim sDirectory As String = "C:\Temp\"
        ' ''Dim sFileName As String = Now.ToString("yyyy-MM-dd") & ".pdf"
        ' ''Dim oconn As New SqlConnection("Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;")
        ' ''oconn.Open()
        ' ''Dim ocomm As New SqlCommand("", oconn)
        ' ''ocomm.CommandText = "Select * from Providers left join Countylist on Providers.county = countylist.countyid where (DATEDIFF(d, Providers.dtModified, '9/8/2011') = 0) order by dtAdded Desc"
        ' ''Dim daRecruit As New SqlDataAdapter(ocomm)

        ' ''Dim myDS As New dsRecruitmentReport
        ' ''daRecruit.Fill(myDS, "Providers")
        ' ''Dim rpt As New rptDailyRecruitment
        ' ''rpt.Database.Tables("Providers").SetDataSource(myDS.Tables("providers"))
        '' ''rpt.SetDataSource(myDS)

        ' ''CrystalReportViewer1.ReportSource = rpt
        '' ''CrystalReportViewer1.Show()
        ' ''Dim myDiskFileDestinationOptions As New DiskFileDestinationOptions()

        '' ''  rpt.
        ' ''rpt.ExportToDisk(ExportFormatType.PortableDocFormat, sDirectory & sFileName)

        MMDownloader.DownloadFromCNAP(connstr)

        '        Try


        '        Catch ex As Exception

        '    Dim sb As New StringBuilder
        '    Dim st As New System.Diagnostics.StackTrace(ex, True)

        '    Dim frame As System.Diagnostics.StackFrame
        '    For Each frame In st.GetFrames()

        '        sb.AppendLine(frame.GetFileName & ": " & frame.GetMethod().ToString & ": " & frame.GetFileLineNumber())

        '        result = sb.ToString()
        '    Next

        ' End Try

    End Sub
End Class
