Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8957
    RowHeight =300
    ItemSuffix =61
    Left =2175
    Top =1080
    Right =13170
    Bottom =6045
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Tag ="@FORMS___F_MENU_ANBAR_TARAZ___DT2=[FORMS]![F_MENU_ANBAR_TARAZ]![DT2]"
    RecSrcDt = Begin
        0x35523f9eecc1e340
    End
    InputParameters ="@FORMS___F_MENU_ANBAR_TARAZ___DT2=[FORMS]![F_MENU_ANBAR_TARAZ]![DT2],@ANB=[FORMS"
        "]![F_MENU_ANBAR_TARAZ]![ANBAR]"
    RecordSource ="dbo.C_TARAZ_ANBAR_KHAS"
    Caption ="تراز موجودي انبار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,220 } Begin
        0x400046004f0052004d0053005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f0054004100520041005a005f005f005f0044005400 ,
        0x32003d005b0046004f0052004d0053005d0021005b0046005f004d0045004e00 ,
        0x55005f0041004e004200410052005f0054004100520041005a005d0021005b00 ,
        0x4400540032005d002c00400041004e0042003d005b0046004f0052004d005300 ,
        0x5d0021005b0046005f004d0045004e0055005f0041004e004200410052005f00 ,
        0x54004100520041005a005d0021005b0041004e004200410052005d00
    End
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =12960
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5194
                    Top =113
                    Width =1440
                    ColumnWidth =900
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8010
                            Top =113
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =453
                    ColumnWidth =915
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MEG"
                    ControlSource ="MEGHAVM"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000000000000000000000000000c00000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900390000003900390039003900 ,
                        0x3900390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8115
                            Top =453
                            Width =390
                            Height =240
                            Name ="Label3"
                            Caption ="مقدار اوليه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =793
                    ColumnWidth =1305
                    ColumnOrder =4
                    TabIndex =2
                    BackColor =16777164
                    Name ="MABAVM"
                    ControlSource ="MABAVM"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x0100000092000000010000000000000000000000000000000c00000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310034003100310031003100310031003100310000003900390039003900 ,
                        0x390039003900390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7365
                            Top =793
                            Width =1140
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ اوليه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =1133
                    ColumnWidth =1095
                    ColumnOrder =15
                    TabIndex =3
                    Name ="MEGHVARED"
                    ControlSource ="MEGHVARED"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000000000000000000000000000c00000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900390000003900390039003900 ,
                        0x3900390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7815
                            Top =1133
                            Width =990
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار افزايش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =1474
                    ColumnWidth =1455
                    ColumnOrder =17
                    TabIndex =4
                    Name ="MABVARED"
                    ControlSource ="MABVARED"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000000000000000000000000000b00000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900000039003900390039003900 ,
                        0x390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =1474
                            Width =915
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ افزايش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =1814
                    ColumnWidth =1050
                    ColumnOrder =30
                    TabIndex =5
                    Name ="MEGHSADER"
                    ControlSource ="MEGHSADER"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000008c000000010000000000000000000000000000000b00000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900000039003900390039003900 ,
                        0x390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =1814
                            Width =945
                            Height =240
                            Name ="Label11"
                            Caption ="مقدار كاهش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =2154
                    ColumnWidth =1515
                    ColumnOrder =32
                    TabIndex =6
                    Name ="MABSADER"
                    ControlSource ="MABSADER"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000008e000000010000000000000000000000000000000c00000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900390000003900390039003900 ,
                        0x3900390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7920
                            Top =2154
                            Width =870
                            Height =240
                            Name ="Label13"
                            Caption ="مبلغ كاهش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =2494
                    ColumnWidth =1170
                    ColumnOrder =33
                    TabIndex =7
                    Name ="MOG"
                    ControlSource ="MOG"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x0100000092000000010000000000000000000000000000000d00000001000000 ,
                        0x0000000033996600000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900390039000000390039003900 ,
                        0x390039003900390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7890
                            Top =2494
                            Width =1065
                            Height =240
                            Name ="Label15"
                            Caption ="مقدار موجودي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4933
                    Top =2834
                    ColumnWidth =1980
                    ColumnOrder =35
                    TabIndex =8
                    Name ="MABLM"
                    ControlSource ="MABLM"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x0100000092000000010000000000000000000000000000000c00000001000000 ,
                        0x0000000033996600000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00390039003900390039003900390039003900390000003900390039003900 ,
                        0x390039003900390039003900390039000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7890
                            Top =2834
                            Width =990
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4924
                    Top =3174
                    Width =1710
                    ColumnWidth =5910
                    ColumnOrder =1
                    TabIndex =9
                    Name ="KALA"
                    ControlSource ="KALA"
                    Format ="Fixed"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8010
                            Top =3174
                            Width =540
                            Height =240
                            Name ="Label19"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4975
                    Top =3540
                    Width =1590
                    ColumnWidth =2175
                    ColumnOrder =39
                    TabIndex =10
                    Name ="ANBNAM"
                    ControlSource ="ANBNAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7936
                            Top =3546
                            Width =570
                            Height =240
                            Name ="Label21"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =283
                    Width =1440
                    ColumnWidth =615
                    ColumnOrder =36
                    TabIndex =11
                    Name ="Text22"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3382
                            Top =283
                            Width =495
                            Height =240
                            Name ="Label23"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5364
                    Top =3855
                    ColumnWidth =645
                    ColumnOrder =37
                    TabIndex =12
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8166
                            Top =3855
                            Width =600
                            Height =240
                            Name ="Label25"
                            Caption ="كد انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5330
                    Top =4422
                    Width =261
                    ColumnOrder =38
                    TabIndex =13
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6662
                            Top =4422
                            Width =630
                            Height =240
                            Name ="Label26"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =737
                    ColumnWidth =1125
                    ColumnOrder =16
                    TabIndex =14
                    Name ="Text27"
                    ControlSource ="=[MABVARED]/IIf([MEGHVARED]=0,1,[MEGHVARED])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3100
                            Top =737
                            Width =915
                            Height =240
                            Name ="Label28"
                            Caption ="في افزايش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1097
                    ColumnWidth =1110
                    ColumnOrder =31
                    TabIndex =15
                    Name ="Text29"
                    ControlSource ="=[MABSADER]/IIf([MEGHSADER]=0,1,[MEGHSADER])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3100
                            Top =1097
                            Width =915
                            Height =240
                            Name ="Label30"
                            Caption ="في كاهش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1457
                    ColumnWidth =1035
                    ColumnOrder =34
                    TabIndex =16
                    Name ="Text31"
                    ControlSource ="=Fix([MABLM])/IIf([MOG]=0,1,[MOG])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3105
                            Top =1455
                            Width =930
                            Height =240
                            Name ="Label32"
                            Caption ="في موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1817
                    ColumnWidth =1065
                    ColumnOrder =3
                    TabIndex =17
                    Name ="Text33"
                    ControlSource ="=[MABAVM]/IIf([MEGHAVM]=0,1,[MEGHAVM])"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3100
                            Top =1817
                            Width =915
                            Height =240
                            Name ="Label34"
                            Caption ="في اوليه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =3174
                    ColumnOrder =40
                    TabIndex =18
                    Name ="grname"
                    ControlSource ="grname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2740
                            Top =3174
                            Width =945
                            Height =240
                            Name ="Label35"
                            Caption ="نام گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =4195
                    ColumnWidth =690
                    ColumnOrder =41
                    TabIndex =19
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3025
                            Top =4195
                            Width =660
                            Height =240
                            Name ="Label36"
                            Caption ="كد گروه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =5555
                    ColumnWidth =645
                    ColumnOrder =5
                    TabIndex =20
                    Name ="MEGHKHM"
                    ControlSource ="MEGHKHM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =5555
                            Width =870
                            Height =240
                            Name ="Label37"
                            Caption ="خريد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =5896
                    ColumnOrder =6
                    TabIndex =21
                    Name ="MABKHM"
                    ControlSource ="MABKHM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =5896
                            Width =765
                            Height =240
                            Name ="Label38"
                            Caption ="مبلغ خريد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =6236
                    ColumnWidth =600
                    ColumnOrder =18
                    TabIndex =22
                    Name ="MEGHFRM"
                    ControlSource ="MEGHFRM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =6236
                            Width =870
                            Height =240
                            Name ="Label39"
                            Caption ="فروش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =6576
                    ColumnOrder =19
                    TabIndex =23
                    Name ="MABFRM"
                    ControlSource ="MABFRM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =6570
                            Width =840
                            Height =240
                            Name ="Label40"
                            Caption ="مبلغ فروش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =6916
                    ColumnWidth =1080
                    ColumnOrder =20
                    TabIndex =24
                    Name ="MEGHENKRM"
                    ControlSource ="MEGHENKRM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =6916
                            Width =1065
                            Height =240
                            Name ="Label41"
                            Caption ="انتقالي خروج"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =7256
                    ColumnOrder =21
                    TabIndex =25
                    Name ="MABENKRM"
                    ControlSource ="MABENKRM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =7260
                            Width =1320
                            Height =240
                            Name ="Label42"
                            Caption ="مبلغ انتقالي خروج"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =7596
                    ColumnWidth =1005
                    ColumnOrder =7
                    TabIndex =26
                    Name ="MEGHENVOM"
                    ControlSource ="MEGHENVOM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =7596
                            Width =1080
                            Height =240
                            Name ="Label43"
                            Caption ="انتقالي ورود"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =7937
                    ColumnOrder =8
                    TabIndex =27
                    Name ="MABENVOM"
                    ControlSource ="MABENVOM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =7935
                            Width =1245
                            Height =240
                            Name ="Label44"
                            Caption ="مبلغ انتقالي ورود"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =8277
                    ColumnWidth =765
                    ColumnOrder =9
                    TabIndex =28
                    Name ="MEGHTOM"
                    ControlSource ="MEGHTOM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =8277
                            Width =885
                            Height =240
                            Name ="Label45"
                            Caption ="توليد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =8617
                    ColumnOrder =10
                    TabIndex =29
                    Name ="MABTOM"
                    ControlSource ="MABTOM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =8617
                            Width =780
                            Height =240
                            Name ="Label46"
                            Caption ="مبلغ توليد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =8957
                    ColumnWidth =870
                    ColumnOrder =22
                    TabIndex =30
                    Name ="MEGHEXM"
                    ControlSource ="MEGHEXM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =8957
                            Width =855
                            Height =240
                            Name ="Label47"
                            Caption ="خروج مواد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =9297
                    ColumnOrder =23
                    TabIndex =31
                    Name ="MABEXM"
                    ControlSource ="MABEXM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =9300
                            Width =1110
                            Height =240
                            Name ="Label48"
                            Caption ="مبلغ خروج مواد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =9637
                    ColumnWidth =960
                    ColumnOrder =24
                    TabIndex =32
                    Name ="MEGHEXSM"
                    ControlSource ="MEGHEXSM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =9630
                            Width =1140
                            Height =240
                            Name ="Label49"
                            Caption ="خروج ساير "
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =9977
                    ColumnOrder =25
                    TabIndex =33
                    Name ="MABEXSM"
                    ControlSource ="MABEXSM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =9975
                            Width =1155
                            Height =240
                            Name ="Label50"
                            Caption ="مبلغ خروج ساير"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =10318
                    ColumnWidth =960
                    ColumnOrder =11
                    TabIndex =34
                    Name ="MEGHKASM"
                    ControlSource ="MEGHKASM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =10318
                            Width =960
                            Height =240
                            Name ="Label51"
                            Caption ="كسري انبار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =10658
                    ColumnOrder =12
                    TabIndex =35
                    Name ="MABKASM"
                    ControlSource ="MABKASM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =10665
                            Width =1200
                            Height =240
                            Name ="Label52"
                            Caption ="مبلغ كسري انبار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =10998
                    ColumnWidth =915
                    ColumnOrder =26
                    TabIndex =36
                    Name ="MEGHEZM"
                    ControlSource ="MEGHEZM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =10998
                            Width =855
                            Height =240
                            Name ="Label53"
                            Caption ="اضافه  انبار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =11338
                    ColumnOrder =27
                    TabIndex =37
                    Name ="MABEZM"
                    ControlSource ="MABEZM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =11340
                            Width =1110
                            Height =240
                            Name ="Label54"
                            Caption ="مبلغ اضافه انبار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =11678
                    ColumnWidth =945
                    ColumnOrder =13
                    TabIndex =38
                    Name ="MEGHBFM"
                    ControlSource ="MEGHBFM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =11678
                            Width =855
                            Height =240
                            Name ="Label55"
                            Caption ="وارده- ساير"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =12018
                    ColumnOrder =14
                    TabIndex =39
                    Name ="MABBFM"
                    ControlSource ="MABBFM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =12015
                            Width =1230
                            Height =240
                            Name ="Label56"
                            Caption ="مبلغ وارده - ساير"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =12378
                    ColumnWidth =1065
                    ColumnOrder =28
                    TabIndex =40
                    Name ="MEGHBKHAM"
                    ControlSource ="MEGHBKHAM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =12375
                            Width =960
                            Height =240
                            Name ="Label58"
                            Caption ="برگشت خريد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1757
                    Top =12718
                    ColumnOrder =29
                    TabIndex =41
                    Name ="MABKHAM"
                    ControlSource ="MABKHAM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4536
                            Top =12720
                            Width =1305
                            Height =240
                            Name ="Label60"
                            Caption ="مبلغ برگشت خريد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "C_TARAZ_ANBAR_KHAS.cls"
