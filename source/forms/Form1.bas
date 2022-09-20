Version =21
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7370
    DatasheetFontHeight =11
    ItemSuffix =13
    Right =25575
    Bottom =13680
    DatasheetGridlinesColor =15132391
    RecSrcDt = Begin
        0x2db1e3f5d5afe540
    End
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    OrderByOnLoad =0
    OrderByOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Tab
            Width =5103
            Height =3402
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =3
            BackThemeColorIndex =1
            BackShade =85.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =2
            BorderTint =60.0
            HoverThemeColorIndex =1
            PressedThemeColorIndex =1
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
            ForeThemeColorIndex =0
            ForeTint =75.0
        End
        Begin FormHeader
            Height =1700
            BackColor =15064278
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Left =1133
                    Top =566
                    Width =5101
                    Height =573
                    ForeColor =4210752
                    Name ="Command0"
                    Caption ="Call API"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =1133
                    LayoutCachedTop =566
                    LayoutCachedWidth =6234
                    LayoutCachedHeight =1139
                    BackColor =14461583
                    BorderColor =14461583
                    HoverColor =15189940
                    PressedColor =9917743
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
        Begin Section
            Height =976
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =576
                    Top =144
                    Width =1994
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text9"
                    GridlineColor =10921638

                    LayoutCachedLeft =576
                    LayoutCachedTop =144
                    LayoutCachedWidth =2570
                    LayoutCachedHeight =459
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =144
                    Width =1994
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text11"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =144
                    LayoutCachedWidth =4874
                    LayoutCachedHeight =459
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5184
                    Top =144
                    Width =1994
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Text12"
                    GridlineColor =10921638

                    LayoutCachedLeft =5184
                    LayoutCachedTop =144
                    LayoutCachedWidth =7178
                    LayoutCachedHeight =459
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command0_Click()

Dim objRs As New ADODB.Recordset

With objRs
.Fields.Append "OrderID", adVarChar, 128, adFldMayBeNull
.Fields.Append "totalList", adVarChar, 128, adFldMayBeNull
.Fields.Append "staffCode", adVarChar, 128, adFldMayBeNull
.Fields.Append "ClearedByMRP", adVarChar, 128, adFldMayBeNull
.Fields.Append "tag", adVarChar, 128, adFldMayBeNull
.Fields.Append "onHoldReason", adVarChar, 128, adFldMayBeNull
.Fields.Append "csr", adVarChar, 128, adFldMayBeNull
.Fields.Append "complexity", adVarChar, 128, adFldMayBeNull
.Fields.Append "pdr", adVarChar, 128, adFldMayBeNull
.Fields.Append "project", adVarChar, 128, adFldMayBeNull
.Fields.Append "productionLotNumber", adVarChar, 128, adFldMayBeNull
.Fields.Append "orderDate", adVarChar, 128, adFldMayBeNull
.Fields.Append "totalNet", adVarChar, 128, adFldMayBeNull
.Fields.Append "faxOrderNotes", adVarChar, 128, adFldMayBeNull
.Fields.Append "lastOfTransactionNote", adVarChar, 128, adFldMayBeNull
.Fields.Append "runOnMRP", adVarChar, 128, adFldMayBeNull

.CursorType = adOpenKeyset
.CursorLocation = adUseClient
.LockType = adLockPessimistic

.Open
End With

Dim objXML           As Object
Dim strURL           As String
Dim rows() As String

Set objXML = CreateObject("MSXML2.XMLHTTP")

strURL = "http://localhost:10000/GetOrderSearches"

objXML.Open "GET", strURL, False
objXML.Send

Dim Json As Object
Set Json = JsonConverter.ParseJson(objXML.ResponseText)
Dim Values As Variant
ReDim Values(Json("Results").Count, 16)

Dim Value As Dictionary
Dim i As Long

i = 0
For Each Value In Json("Results")
  Values(i, 0) = Value("orderID")
  Values(i, 1) = Value("totalList")
  Values(i, 2) = Value("staffCode")
  Values(i, 3) = Value("ClearedByMRP")
  Values(i, 4) = Value("tag")
  Values(i, 5) = Value("onHoldReason")
  Values(i, 6) = Value("csr")
  Values(i, 7) = Value("complexity")
  Values(i, 8) = Value("pdr")
  Values(i, 9) = Value("project")
  Values(i, 10) = Value("productionLotNumber")
  Values(i, 11) = Value("orderDate")
  Values(i, 12) = Value("totalNet")
  Values(i, 13) = Value("faxOrderNotes")
  Values(i, 14) = Value("lastOfTransactionNote")
  Values(i, 15) = Value("runOnMRP")
  i = i + 1
Next Value

Dim w As Long
For w = 0 To UBound(Values)
    With objRs
    
    .AddNew
    .Fields(0) = Values(w, 0)
    .Fields(1) = Values(w, 1)
    .Fields(2) = Values(w, 2)
    .Fields(3) = Values(w, 3)
    .Fields(4) = Values(w, 4)
    .Fields(5) = Values(w, 5)
    .Fields(6) = Values(w, 6)
    .Fields(7) = Values(w, 7)
    .Fields(8) = Values(w, 8)
    .Fields(9) = Values(w, 9)
    .Fields(10) = Values(w, 10)
    .Fields(11) = Values(w, 11)
    .Fields(12) = Values(w, 12)
    .Fields(13) = Values(w, 13)
    .Fields(14) = Values(w, 14)
    .Fields(15) = Values(w, 15)
    .Update
    
    End With
Next

Set Forms("Form1").Recordset = objRs

Me!Text9.ControlSource = objRs.Fields(0).Name
Me!Text11.ControlSource = objRs.Fields(1).Name
Me!Text12.ControlSource = objRs.Fields(2).Name

objRs.Close

End Sub
