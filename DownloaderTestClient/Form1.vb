Imports DownloaderLibs
Imports System.Drawing.Imaging
Imports CrystalDecisions.Shared
Imports CrystalDecisions
Imports System.Data.SqlClient
Imports System.Net.Mail

Imports System.Text

Public Class Form1

    Private connstr As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connstr = "Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim count As Integer = MMDownloader.DownloadFromCNAP(connstr) ' MMDownloader.QueryCenterExcelSheet(connstr)
        'Dim count As Integer = MMDownloader.read_from_json_text_file("c:\temp\cdcDF.txt")
        '    MessageBox.Show(count & " new providers added")

        'MessageBox.Show(TimeString)


        ' ''Dim cfg As New MMConfig
        'MMConfig.Initialize(System.Reflection.Assembly.GetExecutingAssembly.Location.Substring(0, _
        '       System.Reflection.Assembly.GetExecutingAssembly.Location.IndexOf("DownloaderTestClient")) & "DownloaderSVC\bin\Debug\downloadercfg.txt")

        ' ''Dim stime As String = MMConfig.GetOption("DownloadTime")


        'Dim sdate As Date = #9/16/2011#
        'Dim sDirectory As String = "C:\Temp\"
        'Dim sFileName As String = sdate.ToString("yyyy-MM-dd") & " .pdf"
        'Dim oconn As New SqlConnection("Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;")
        'oconn.Open()
        'Dim ocomm As New SqlCommand("", oconn)
        'ocomm.CommandText = "Select * from Providers left join Countylist on Providers.county = countylist.countyid where (DATEDIFF(d, Providers.dtModified, '" & sdate & "') = 0) order by dtAdded Desc"
        'Dim daRecruit As New SqlDataAdapter(ocomm)

        'Dim myDS As New dsRecruitmentReport
        'daRecruit.Fill(myDS, "Providers")
        'Dim rpt As New rptDailyRecruitment
        'rpt.Database.Tables("Providers").SetDataSource(myDS.Tables("providers"))


        'CrystalReportViewer1.ReportSource = rpt
        'CrystalReportViewer1.Show()
        'Dim myDiskFileDestinationOptions As New DiskFileDestinationOptions()


        'rpt.ExportToDisk(ExportFormatType.PortableDocFormat, sDirectory & sFileName)

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

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

        Dim code As String = TextBox1.Text


        Dim bytearray() As Byte = Encoding.UTF8.GetBytes(code)

        'TextBox2.Text = New System.Text.ASCIIEncoding().GetString(Convert.FromBase64String(TextBox1.Text))
        '   Dim str As String = Convert.FromBase64String(bytearray.ToString).ToString
        'Dim decoded() As Byte = Convert.FromBase64Strin
        Dim src As String = TextBox1.Text
        'Dim bt64 As Byte() =

        '        Dim streamWriter As New System.IO.StreamWriter(memoryStream)

        'streamWriter.Write(bt64.ToString)
        'memoryStream.Position = 0
        
        'Dim jpg As System.Drawing.Image
        Dim oconn As New SqlConnection(connstr)
        Dim ocomm As New SqlCommand("Select ProviderAssistantConsultantSignature from midmichreviewimport where reviewid = 263", oconn)
        oconn.Open()
        Dim sigstring As String = ocomm.ExecuteScalar()
        oconn.Close()

        Dim fs As New System.IO.FileStream("C:\Users\Kevin\Downloads\testing.jpg", IO.FileMode.Open)

        '  Dim mystr As String = Encoding.UTF8.GetString(Convert.FromBase64String(src))

        Dim filestring As String
        Dim strrd As New System.IO.StreamReader("C:\Users\Kevin\Downloads\encoded.txt")
        filestring = TextBox1.Text 'strrd.ReadToEnd

        Dim bytes() As Byte = New Byte((fs.Length) - 1) {}
        Dim i As Integer = fs.Read(bytes, 0, fs.Length)
        Dim ms As New System.IO.MemoryStream(Convert.FromBase64String(Mid(sigstring, 26)))

        Dim imagestream As New System.IO.MemoryStream
        'Dim streamwriter As New System.IO.StreamWriter(imagestream)
        ' streamwriter.Write(fs.ReadToEnd)
 
        PictureBox1.Image = Image.FromStream(ms)

        ' Dim bytearray() As Byte = Convert.FromBase64String(code)
        '       Dim decoded As String = Encoding.UTF8.GetString(bytearray)
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MMDownloader.connstr = "Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;"
        MMDownloader.CreateReport()
    End Sub
End Class
