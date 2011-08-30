Imports DownloaderLibs

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Globals.connstr = "Data Source=200.200.100.20;Initial Catalog=MidMich;USER ID=sa;Password=noblank2day;"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MMDownloader.DownloadFromCNAP()
    End Sub
End Class
