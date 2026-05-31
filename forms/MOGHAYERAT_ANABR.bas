Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6973
    RowHeight =360
    ItemSuffix =20
    Left =1620
    Top =255
    Right =12210
    Bottom =5070
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    Tag ="@DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@ANBAR=FORMS![AK_MOGUDI_ANBAR_LIST]![A"
        "NBAR],@KOL = FORMS![BASEKNOW]![MOGODIA]"
    OrderBy ="MOGHA_ANBAR.TAFBES DESC"
    RecSrcDt = Begin
        0x78cd48bb2fdce540
    End
    InputParameters ="@DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@ANBAR=FORMS![AK_MOGUDI_ANBAR_LIST]![A"
        "NBAR],@KOL = FORMS![BASEKNOW]![MOGODIA]"
    RecordSource ="dbo.MOGHA_ANBAR"
    Caption ="مغايرت كارت انبار و حسابداري"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,238 } Begin
        0x4000440054003200200042004900470049004e00540020003d005b0046006f00 ,
        0x72006d0073005d0021005b0046005f004d0045004e0055005f0041004e004200 ,
        0x410052005d0021005b004400540032005d002c00400041004e00420041005200 ,
        0x3d0046004f0052004d00530021005b0041004b005f004d004f00470055004400 ,
        0x49005f0041004e004200410052005f004c004900530054005d0021005b004100 ,
        0x4e004200410052005d002c0040004b004f004c0020003d00200046004f005200 ,
        0x4d00530021005b0042004100530045004b004e004f0057005d0021005b004d00 ,
        0x4f0047004f004400490041005d00
    End
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =5385
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3231
                    Top =907
                    Width =2490
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1424
                            Top =914
                            Width =510
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3231
                    Top =1247
                    ColumnWidth =2040
                    ColumnOrder =1
                    TabIndex =1
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Format ="#,###"
                    OnDblClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000003000000000000000500000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d0030002e00320000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1424
                            Top =1244
                            Width =1125
                            Height =240
                            Name ="Label3"
                            Caption ="مانده  كارت انبار"
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
                    Left =3231
                    Top =1587
                    ColumnWidth =1170
                    ColumnOrder =6
                    TabIndex =2
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1424
                            Top =1589
                            Width =1125
                            Height =240
                            Name ="Label5"
                            Caption ="مقدار موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3231
                    Top =1927
                    ColumnWidth =1545
                    ColumnOrder =2
                    TabIndex =3
                    Name ="mab"
                    ControlSource ="mab"
                    Format ="#,###"
                    OnDblClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000003000000000000000500000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d0030002e00320000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1424
                            Top =1934
                            Width =1545
                            Height =240
                            Name ="Label7"
                            Caption ="مانده كارت حسابداري"
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
                    Left =3231
                    Top =2268
                    ColumnOrder =4
                    TabIndex =4
                    Name ="tafBED"
                    ControlSource ="tafBED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1424
                            Top =2264
                            Width =1050
                            Height =240
                            Name ="Label9"
                            Caption ="مغايرت بدهكار"
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
                    Left =3231
                    Top =2608
                    ColumnOrder =5
                    TabIndex =5
                    Name ="TAFBES"
                    ControlSource ="TAFBES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1424
                            Top =2609
                            Width =1170
                            Height =240
                            Name ="Label11"
                            Caption ="مغايرت بستانكار"
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
                    Left =3231
                    Top =2948
                    ColumnWidth =1200
                    ColumnOrder =7
                    TabIndex =6
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1417
                            Top =2948
                            Width =1035
                            Height =240
                            Name ="Label13"
                            Caption =" تفضيلي"
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
                    Left =3231
                    Top =3288
                    ColumnOrder =8
                    TabIndex =7
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1417
                            Top =3288
                            Width =780
                            Height =240
                            Name ="Label15"
                            Caption ="حساب كل"
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
                    Left =3231
                    Top =3628
                    ColumnOrder =9
                    TabIndex =8
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1417
                            Top =3628
                            Width =945
                            Height =240
                            Name ="Label17"
                            Caption =" معين"
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
                    Left =3231
                    Top =3968
                    Width =2490
                    ColumnWidth =1980
                    ColumnOrder =3
                    TabIndex =9
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1417
                            Top =3968
                            Width =540
                            Height =240
                            Name ="Label19"
                            Caption ="حساب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MOGHAYERAT_ANABR.cls"
