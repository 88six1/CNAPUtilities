<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class ProviderListReport
    Inherits DevExpress.XtraReports.UI.XtraReport

    'XtraReport overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Designer
    'It can be modified using the Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TopMargin = New DevExpress.XtraReports.UI.TopMarginBand()
        Me.BottomMargin = New DevExpress.XtraReports.UI.BottomMarginBand()
        Me.Detail = New DevExpress.XtraReports.UI.DetailBand()
        Me.XrLabel7 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrLabel6 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrLabel5 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrLabel4 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrLabel3 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrLabel2 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrLabel1 = New DevExpress.XtraReports.UI.XRLabel()
        Me.XrCrossBandBox1 = New DevExpress.XtraReports.UI.XRCrossBandBox()
        Me.XrCrossBandLine2 = New DevExpress.XtraReports.UI.XRCrossBandLine()
        Me.XrCrossBandLine3 = New DevExpress.XtraReports.UI.XRCrossBandLine()
        Me.XrCrossBandLine4 = New DevExpress.XtraReports.UI.XRCrossBandLine()
        Me.XrCrossBandLine5 = New DevExpress.XtraReports.UI.XRCrossBandLine()
        Me.XrCrossBandLine6 = New DevExpress.XtraReports.UI.XRCrossBandLine()
        Me.XrCrossBandLine1 = New DevExpress.XtraReports.UI.XRCrossBandLine()
        Me.XrLine1 = New DevExpress.XtraReports.UI.XRLine()
        CType(Me, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'TopMargin
        '
        Me.TopMargin.HeightF = 101.0!
        Me.TopMargin.Name = "TopMargin"
        '
        'BottomMargin
        '
        Me.BottomMargin.Name = "BottomMargin"
        '
        'Detail
        '
        Me.Detail.Controls.AddRange(New DevExpress.XtraReports.UI.XRControl() {Me.XrLine1, Me.XrLabel7, Me.XrLabel6, Me.XrLabel5, Me.XrLabel4, Me.XrLabel3, Me.XrLabel2, Me.XrLabel1})
        Me.Detail.HeightF = 20.79168!
        Me.Detail.Name = "Detail"
        '
        'XrLabel7
        '
        Me.XrLabel7.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[LicenseEnd]")})
        Me.XrLabel7.LocationFloat = New DevExpress.Utils.PointFloat(820.0!, 0!)
        Me.XrLabel7.Multiline = True
        Me.XrLabel7.Name = "XrLabel7"
        Me.XrLabel7.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel7.SizeF = New System.Drawing.SizeF(75.0!, 20.20833!)
        Me.XrLabel7.Text = "XrLabel7"
        Me.XrLabel7.TextFormatString = "{0:d}"
        '
        'XrLabel6
        '
        Me.XrLabel6.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[LicenseBegin]")})
        Me.XrLabel6.LocationFloat = New DevExpress.Utils.PointFloat(735.7917!, 0!)
        Me.XrLabel6.Multiline = True
        Me.XrLabel6.Name = "XrLabel6"
        Me.XrLabel6.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel6.SizeF = New System.Drawing.SizeF(80.20831!, 20.20833!)
        Me.XrLabel6.Text = "XrLabel6"
        Me.XrLabel6.TextFormatString = "{0:d}"
        '
        'XrLabel5
        '
        Me.XrLabel5.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[Phone]")})
        Me.XrLabel5.LocationFloat = New DevExpress.Utils.PointFloat(642.9167!, 0!)
        Me.XrLabel5.Multiline = True
        Me.XrLabel5.Name = "XrLabel5"
        Me.XrLabel5.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel5.SizeF = New System.Drawing.SizeF(87.5!, 20.20833!)
        Me.XrLabel5.Text = "XrLabel5"
        '
        'XrLabel4
        '
        Me.XrLabel4.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[Address]")})
        Me.XrLabel4.LocationFloat = New DevExpress.Utils.PointFloat(433.9167!, 0!)
        Me.XrLabel4.Multiline = True
        Me.XrLabel4.Name = "XrLabel4"
        Me.XrLabel4.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel4.SizeF = New System.Drawing.SizeF(204.1667!, 20.20833!)
        Me.XrLabel4.Text = "XrLabel4"
        '
        'XrLabel3
        '
        Me.XrLabel3.CanGrow = False
        Me.XrLabel3.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[CountyName]")})
        Me.XrLabel3.LocationFloat = New DevExpress.Utils.PointFloat(336.8749!, 0!)
        Me.XrLabel3.Multiline = True
        Me.XrLabel3.Name = "XrLabel3"
        Me.XrLabel3.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel3.SizeF = New System.Drawing.SizeF(89.58331!, 20.20833!)
        Me.XrLabel3.Text = "XrLabel3"
        '
        'XrLabel2
        '
        Me.XrLabel2.CanGrow = False
        Me.XrLabel2.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[Name]")})
        Me.XrLabel2.LocationFloat = New DevExpress.Utils.PointFloat(115.0002!, 0!)
        Me.XrLabel2.Multiline = True
        Me.XrLabel2.Name = "XrLabel2"
        Me.XrLabel2.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel2.SizeF = New System.Drawing.SizeF(216.6667!, 20.20833!)
        Me.XrLabel2.Text = "XrLabel2"
        '
        'XrLabel1
        '
        Me.XrLabel1.ExpressionBindings.AddRange(New DevExpress.XtraReports.UI.ExpressionBinding() {New DevExpress.XtraReports.UI.ExpressionBinding("BeforePrint", "Text", "[providerid]")})
        Me.XrLabel1.LocationFloat = New DevExpress.Utils.PointFloat(6.00001!, 0!)
        Me.XrLabel1.Multiline = True
        Me.XrLabel1.Name = "XrLabel1"
        Me.XrLabel1.Padding = New DevExpress.XtraPrinting.PaddingInfo(2, 2, 0, 0, 100.0!)
        Me.XrLabel1.SizeF = New System.Drawing.SizeF(102.0833!, 20.20833!)
        Me.XrLabel1.Text = "XrLabel1"
        '
        'XrCrossBandBox1
        '
        Me.XrCrossBandBox1.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandBox1.EndBand = Me.BottomMargin
        Me.XrCrossBandBox1.EndPointFloat = New DevExpress.Utils.PointFloat(0!, 3.125!)
        Me.XrCrossBandBox1.Name = "XrCrossBandBox1"
        Me.XrCrossBandBox1.StartBand = Me.TopMargin
        Me.XrCrossBandBox1.StartPointFloat = New DevExpress.Utils.PointFloat(0!, 98.95834!)
        Me.XrCrossBandBox1.WidthF = 900.0!
        '
        'XrCrossBandLine2
        '
        Me.XrCrossBandLine2.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandLine2.EndBand = Me.Detail
        Me.XrCrossBandLine2.EndPointFloat = New DevExpress.Utils.PointFloat(333.875!, 20.79168!)
        Me.XrCrossBandLine2.Name = "XrCrossBandLine2"
        Me.XrCrossBandLine2.StartBand = Me.Detail
        Me.XrCrossBandLine2.StartPointFloat = New DevExpress.Utils.PointFloat(333.875!, 0!)
        Me.XrCrossBandLine2.WidthF = 1.999998!
        '
        'XrCrossBandLine3
        '
        Me.XrCrossBandLine3.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandLine3.EndBand = Me.Detail
        Me.XrCrossBandLine3.EndPointFloat = New DevExpress.Utils.PointFloat(430.9167!, 20.79168!)
        Me.XrCrossBandLine3.Name = "XrCrossBandLine3"
        Me.XrCrossBandLine3.StartBand = Me.Detail
        Me.XrCrossBandLine3.StartPointFloat = New DevExpress.Utils.PointFloat(430.9167!, 0!)
        Me.XrCrossBandLine3.WidthF = 2.000014!
        '
        'XrCrossBandLine4
        '
        Me.XrCrossBandLine4.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandLine4.EndBand = Me.Detail
        Me.XrCrossBandLine4.EndPointFloat = New DevExpress.Utils.PointFloat(639.1249!, 20.79168!)
        Me.XrCrossBandLine4.Name = "XrCrossBandLine4"
        Me.XrCrossBandLine4.StartBand = Me.Detail
        Me.XrCrossBandLine4.StartPointFloat = New DevExpress.Utils.PointFloat(639.1249!, 0!)
        Me.XrCrossBandLine4.WidthF = 2.0!
        '
        'XrCrossBandLine5
        '
        Me.XrCrossBandLine5.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandLine5.EndBand = Me.Detail
        Me.XrCrossBandLine5.EndPointFloat = New DevExpress.Utils.PointFloat(732.8333!, 20.79168!)
        Me.XrCrossBandLine5.Name = "XrCrossBandLine5"
        Me.XrCrossBandLine5.StartBand = Me.Detail
        Me.XrCrossBandLine5.StartPointFloat = New DevExpress.Utils.PointFloat(732.8333!, 0!)
        Me.XrCrossBandLine5.WidthF = 2.000014!
        '
        'XrCrossBandLine6
        '
        Me.XrCrossBandLine6.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandLine6.EndBand = Me.Detail
        Me.XrCrossBandLine6.EndPointFloat = New DevExpress.Utils.PointFloat(818.0!, 20.79168!)
        Me.XrCrossBandLine6.Name = "XrCrossBandLine6"
        Me.XrCrossBandLine6.StartBand = Me.Detail
        Me.XrCrossBandLine6.StartPointFloat = New DevExpress.Utils.PointFloat(818.0!, 0!)
        Me.XrCrossBandLine6.WidthF = 2.000014!
        '
        'XrCrossBandLine1
        '
        Me.XrCrossBandLine1.AnchorVertical = CType((DevExpress.XtraReports.UI.VerticalAnchorStyles.Top Or DevExpress.XtraReports.UI.VerticalAnchorStyles.Bottom), DevExpress.XtraReports.UI.VerticalAnchorStyles)
        Me.XrCrossBandLine1.EndBand = Me.BottomMargin
        Me.XrCrossBandLine1.EndPointFloat = New DevExpress.Utils.PointFloat(109.3751!, 1.041666!)
        Me.XrCrossBandLine1.Name = "XrCrossBandLine1"
        Me.XrCrossBandLine1.StartBand = Me.Detail
        Me.XrCrossBandLine1.StartPointFloat = New DevExpress.Utils.PointFloat(109.3751!, 1.041667!)
        Me.XrCrossBandLine1.WidthF = 2.000046!
        '
        'XrLine1
        '
        Me.XrLine1.BorderWidth = 0!
        Me.XrLine1.LineWidth = 2.0!
        Me.XrLine1.LocationFloat = New DevExpress.Utils.PointFloat(0!, 18.79168!)
        Me.XrLine1.Name = "XrLine1"
        Me.XrLine1.SizeF = New System.Drawing.SizeF(900.0!, 2.0!)
        Me.XrLine1.StylePriority.UseBorderWidth = False
        '
        'ProviderListReport
        '
        Me.Bands.AddRange(New DevExpress.XtraReports.UI.Band() {Me.TopMargin, Me.BottomMargin, Me.Detail})
        Me.CrossBandControls.AddRange(New DevExpress.XtraReports.UI.XRCrossBandControl() {Me.XrCrossBandLine1, Me.XrCrossBandLine6, Me.XrCrossBandLine5, Me.XrCrossBandLine4, Me.XrCrossBandLine3, Me.XrCrossBandLine2, Me.XrCrossBandBox1})
        Me.DataMember = "Providers"
        Me.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.Landscape = True
        Me.Margins = New System.Drawing.Printing.Margins(50, 40, 101, 100)
        Me.PageHeight = 850
        Me.PageWidth = 1100
        Me.SnappingMode = DevExpress.XtraReports.UI.SnappingMode.None
        Me.Version = "19.2"
        CType(Me, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Friend WithEvents DsRecruitmentReport1 As dsRecruitmentReport
    Friend WithEvents TopMargin As DevExpress.XtraReports.UI.TopMarginBand
    Friend WithEvents BottomMargin As DevExpress.XtraReports.UI.BottomMarginBand
    Friend WithEvents Detail As DevExpress.XtraReports.UI.DetailBand
    Friend WithEvents XrLabel7 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrLabel6 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrLabel5 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrLabel4 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrLabel3 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrLabel2 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrLabel1 As DevExpress.XtraReports.UI.XRLabel
    Friend WithEvents XrCrossBandBox1 As DevExpress.XtraReports.UI.XRCrossBandBox
    Friend WithEvents XrCrossBandLine2 As DevExpress.XtraReports.UI.XRCrossBandLine
    Friend WithEvents XrCrossBandLine3 As DevExpress.XtraReports.UI.XRCrossBandLine
    Friend WithEvents XrCrossBandLine4 As DevExpress.XtraReports.UI.XRCrossBandLine
    Friend WithEvents XrCrossBandLine5 As DevExpress.XtraReports.UI.XRCrossBandLine
    Friend WithEvents XrCrossBandLine6 As DevExpress.XtraReports.UI.XRCrossBandLine
    Friend WithEvents XrCrossBandLine1 As DevExpress.XtraReports.UI.XRCrossBandLine
    Friend WithEvents XrLine1 As DevExpress.XtraReports.UI.XRLine
End Class
