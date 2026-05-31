Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5621
    RowHeight =330
    ItemSuffix =24
    Left =510
    Top =600
    Right =21450
    Bottom =10020
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Tag ="@FROMDATE=FORMS![F_MENU_DATE]![DT1],@TODATE=FORMS![F_MENU_DATE]![DT2],@KONTROL=F"
        "ORMS![BASEKNOW]![CONKAL]"
    RecSrcDt = Begin
        0x3b7f9e8f1b60e640
    End
    InputParameters ="@FROMDATE=FORMS![F_MENU_DATE]![DT1],@TODATE=FORMS![F_MENU_DATE]![DT2],@KONTROL=F"
        "ORMS![BASEKNOW]![CONKAL]"
    RecordSource ="sp_AMAR_MASRAF"
    Caption ="گزارش مصرف مواد"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,208 } Begin
        0x4000460052004f004d0044004100540045003d0046004f0052004d0053002100 ,
        0x5b0046005f004d0045004e0055005f0044004100540045005d0021005b004400 ,
        0x540031005d002c00400054004f0044004100540045003d0046004f0052004d00 ,
        0x530021005b0046005f004d0045004e0055005f0044004100540045005d002100 ,
        0x5b004400540032005d002c0040004b004f004e00540052004f004c003d004600 ,
        0x4f0052004d00530021005b0042004100530045004b004e004f0057005d002100 ,
        0x5b0043004f004e004b0041004c005d00
    End
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =5166
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
                    Left =2608
                    Top =113
                    Width =2490
                    ColumnWidth =615
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =113
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2608
                    Top =453
                    ColumnWidth =1125
                    ColumnOrder =3
                    TabIndex =1
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =795
                            Top =450
                            Width =1020
                            Height =240
                            Name ="Label3"
                            Caption ="مصرف واقعي"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2608
                    Top =793
                    ColumnWidth =1410
                    ColumnOrder =6
                    TabIndex =2
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =795
                            Top =795
                            Width =1125
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ كل واقعي"
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
                    Left =2608
                    Top =1133
                    Width =2490
                    ColumnWidth =3210
                    ColumnOrder =1
                    TabIndex =3
                    Name ="KALA"
                    ControlSource ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =1133
                            Width =450
                            Height =240
                            Name ="Label7"
                            Caption ="كالا"
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
                    Left =2608
                    Top =1474
                    Width =2490
                    ColumnWidth =750
                    ColumnOrder =2
                    TabIndex =4
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =1474
                            Width =630
                            Height =240
                            Name ="Label9"
                            Caption ="واحد"
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
                    Left =2608
                    Top =1814
                    Width =2490
                    ColumnWidth =1230
                    ColumnOrder =12
                    TabIndex =5
                    Name ="GRP"
                    ControlSource ="GRP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =794
                            Top =1814
                            Width =885
                            Height =240
                            Name ="Label11"
                            Caption ="گروه كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =227
                    Top =3004
                    ColumnWidth =1410
                    ColumnOrder =7
                    TabIndex =6
                    Name ="mabst"
                    ControlSource ="mabst"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3060
                            Top =3000
                            Width =1290
                            Height =240
                            Name ="Label12"
                            Caption ="مبلغ كل استاندارد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =226
                    Top =2725
                    ColumnWidth =1095
                    ColumnOrder =8
                    TabIndex =7
                    Name ="price"
                    ControlSource ="price"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3151
                            Top =2721
                            Width =990
                            Height =240
                            Name ="Label13"
                            Caption ="في استاندارد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =226
                    Top =3296
                    ColumnWidth =1410
                    ColumnOrder =10
                    TabIndex =8
                    Name ="mognerkh"
                    ControlSource ="mognerkh"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3345
                            Top =3285
                            Width =825
                            Height =240
                            Name ="Label15"
                            Caption ="مغايرت نرخ"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =340
                    Top =3802
                    ColumnWidth =1410
                    ColumnOrder =9
                    TabIndex =9
                    Name ="MABLST"
                    ControlSource ="MABLST"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3053
                            Top =3798
                            Width =1170
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ استاندارد  "
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =226
                    Top =4199
                    ColumnWidth =1155
                    ColumnOrder =5
                    TabIndex =10
                    Name ="MOGMAS"
                    ControlSource ="MOGMAS"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3345
                            Top =4200
                            Width =1050
                            Height =240
                            Name ="Label19"
                            Caption ="مغايرت مصرف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =226
                    Top =4563
                    ColumnWidth =1305
                    ColumnOrder =4
                    TabIndex =11
                    Name ="MEGHST"
                    ControlSource ="MEGHST"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3345
                            Top =4560
                            Width =1200
                            Height =240
                            Name ="Label21"
                            Caption ="مصرف استاندارد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =226
                    Top =4926
                    ColumnWidth =1560
                    ColumnOrder =11
                    TabIndex =12
                    Name ="MOGRMAS"
                    ControlSource ="MOGRMAS"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3345
                            Top =4920
                            Width =1455
                            Height =240
                            Name ="Label23"
                            Caption ="مغايرت ريالي مصرف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GOZARESH_MASRAF_MAVAD.cls"
