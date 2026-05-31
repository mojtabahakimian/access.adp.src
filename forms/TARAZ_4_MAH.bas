Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    AllowUpdating =4
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11340
    RowHeight =315
    ItemSuffix =27
    Left =2955
    Right =14310
    Bottom =7845
    HelpContextId =5009
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    PaintPalette = Begin
        0x000359000000000000ccff0005cdff000aceff000fcfff0014d0ff001ad1ff00 ,
        0x1fd2ff0024d3ff0029d4ff002ed5ff0034d6ff0039d7ff003ed8ff0043d9ff00 ,
        0x48daff004edbff0053dcff0058ddff005ddeff0062dfff0068e0ff006de1ff00 ,
        0x72e2ff0077e3ff007ce4ff0082e6ff0087e7ff008ce8ff0091e9ff0096eaff00 ,
        0x9cebff00a1ecff00a6edff00abeeff00b0efff00b6f0ff00bbf1ff00c0f2ff00 ,
        0xc5f3ff00caf4ff00cccccc00cfcfcf00d0f5ff00d1d1d100d4d4d400d5f6ff00 ,
        0xd6d6d600d9d9d900daf7ff00dbdbdb00dedede00dff8ff00e0e0e000e3e3e300 ,
        0xe4f9ff00e5e5e500e8e8e800eafaff00ebebeb00ededed00effbff00f0f0f000 ,
        0xf2f2f200f4fcff00f5f5f500f7f7f700f9fdff00fafafa00fcfcfc00ffffff00 ,
        0x800000000080000080800000000080008000800000808000c0c0c000c0dcc000 ,
        0xa6caf000fffbf000a0a0a40080808000ff00000000ff0000ffff00000000ff00 ,
        0xff00ff0000ffff0000000000
    End
    Filter ="((NUMBER=411))"
    RecSrcDt = Begin
        0x3394c5807084e340
    End
    RecordSource ="SELECT     TOP 100 PERCENT dbo.TOTA_HES.NUMBER, dbo.Umonth(dbo.DEED_HED.DATE_S) "
        "AS MON, dbo.TOTA_HES.NAME,                        dbo.UIIF(SUM(dbo.DEED_DTL.BED)"
        " - SUM(dbo.DEED_DTL.BES), N'>', 0, SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEED_DTL.BES)"
        ", 0) AS BED,                        dbo.UIIF(SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEE"
        "D_DTL.BES), N'<', 0, (SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEED_DTL.BES)) * - 1, 0)  "
        "                      AS BES FROM         dbo.DEED_HED INNER JOIN               "
        "        dbo.DEED_DTL INNER JOIN                       dbo.TOTA_HES ON dbo.DEED_D"
        "TL.HES_K = dbo.TOTA_HES.NUMBER ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S GROUP BY d"
        "bo.TOTA_HES.NUMBER, dbo.TOTA_HES.NAME, dbo.Umonth(dbo.DEED_HED.DATE_S) ORDER BY "
        "dbo.TOTA_HES.NUMBER, dbo.Umonth(dbo.DEED_HED.DATE_S)"
    Caption ="تراز حسابها به تفكيك ماه"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            BorderWidth =1
            Width =1701
        End
        Begin Image
            SpecialEffect =1
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =1
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =1
            BorderWidth =3
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin BoundObjectFrame
            SpecialEffect =1
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            SpecialEffect =1
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            SpecialEffect =1
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =1
            BackStyle =0
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial (Arabic)"
        End
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =300
            BackColor =10079487
            Name ="FormHeader"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =10358
                    Width =975
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="NUMBER Label"
                    Caption ="شماره "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NUMBER_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =7310
                    Width =1020
                    Height =300
                    ForeColor =8388608
                    Name ="NAME Label"
                    Caption ="نام حساب"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NAME_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =1864
                    Width =1765
                    Height =300
                    BackColor =13421619
                    ForeColor =8388608
                    Name ="bed Label"
                    Caption ="مانده بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bed_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =75
                    Width =1776
                    Height =300
                    BackColor =16776960
                    ForeColor =8388608
                    Name ="bes Label"
                    Caption ="مانده بستانكار"
                    FontName ="WeblogmaYekan"
                    OnDblClick ="[Event Procedure]"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bes_Label"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =9135
                    Width =1215
                    Height =285
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="Label24"
                    Caption ="ماه "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =315
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =10389
                    Top =-3
                    Width =939
                    Height =315
                    ColumnWidth =1065
                    BackColor =13434828
                    Name ="شماره حساب"
                    ControlSource ="NUMBER"
                    OnDblClick ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="شماره_حساب"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3685
                    Top =-3
                    Width =5439
                    Height =315
                    ColumnWidth =3330
                    TabIndex =1
                    Name ="نام حساب"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="نام_حساب"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1863
                    Top =-3
                    Width =1764
                    Height =315
                    ColumnWidth =1545
                    TabIndex =2
                    BackColor =13421619
                    Name ="مانده بدهكار"
                    ControlSource ="BED"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="مانده_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =75
                    Top =-3
                    Width =1776
                    Height =315
                    ColumnWidth =1530
                    TabIndex =3
                    BackColor =16776960
                    Name ="مانده بستانكار"
                    ControlSource ="BES"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="مانده_بستانكار"
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =2
                    Left =11340
                    Width =0
                    Height =283
                    Name ="Line22"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9151
                    Top =-3
                    Width =1164
                    Height =315
                    TabIndex =4
                    BackColor =13434828
                    Name ="ماه"
                    ControlSource ="MON"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =5079
                    Width =939
                    Height =315
                    TabIndex =5
                    BackColor =13434828
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            Height =345
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Width =1749
                    Height =330
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=Sum([bed])"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                    InputMask ="#,###"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Width =1761
                    Height =330
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text15"
                    ControlSource ="=Sum([bes])"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "TARAZ_4_MAH.cls"
