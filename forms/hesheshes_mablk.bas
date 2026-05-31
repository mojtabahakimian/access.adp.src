Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =345
    ItemSuffix =33
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x89a89c7153b0e540
    End
    RecordSource ="SELECT ROUND(MABL_K, 0) AS mablkr, ROUND(MEGHk * MABL, 0) AS mghmablk, MABL, TAG"
        ", NUMBER, CODE, MEGH, MEGHk, ANBAR, ROUND(MABL_K, 0) - ROUND(MEGHk * MABL, 0) AS"
        " ekh FROM INVO_LST WHERE (ROUND(MABL_K, 0) <> ROUND(MEGHk * MABL, 0)) AND (ROUND"
        "(MABL_K, 0) - ROUND(MEGHk * MABL, 0) > 1 AND ROUND(MABL_K, 0) - ROUND(MEGHk * MA"
        "BL, 0) < - 1) ORDER BY ROUND(MABL_K, 0) - ROUND(MEGHk * MABL, 0) DESC"
    Caption ="محاسبات سطر فاکتور داراي اختلاف"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
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
            Height =4435
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =1474
                    ColumnWidth =1680
                    ColumnOrder =4
                    Name ="mablkr"
                    ControlSource ="mablkr"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000003000000000000000300000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2970
                            Top =1470
                            Width =1050
                            Height =240
                            Name ="Label23"
                            Caption ="مبلغ ثبت شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =1814
                    ColumnWidth =1740
                    ColumnOrder =5
                    TabIndex =1
                    Name ="mghmablk"
                    ControlSource ="mghmablk"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x0100000068000000010000000000000003000000000000000300000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x2d00310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2730
                            Top =1815
                            Width =1410
                            Height =240
                            Name ="Label24"
                            Caption ="مبلغ محاسبه شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =2154
                    ColumnWidth =2100
                    ColumnOrder =7
                    TabIndex =2
                    Name ="MABL"
                    ControlSource ="MABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3136
                            Top =2154
                            Width =435
                            Height =240
                            Name ="Label25"
                            Caption ="في"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =2494
                    ColumnWidth =735
                    ColumnOrder =1
                    TabIndex =3
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2881
                            Top =2494
                            Width =690
                            Height =240
                            Name ="Label26"
                            Caption ="نوع برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =2834
                    ColumnWidth =960
                    ColumnOrder =0
                    TabIndex =4
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2551
                            Top =2834
                            Width =1020
                            Height =240
                            Name ="Label27"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =3174
                    ColumnOrder =3
                    TabIndex =5
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3001
                            Top =3174
                            Width =570
                            Height =240
                            Name ="Label28"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =3514
                    ColumnOrder =8
                    TabIndex =6
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2836
                            Top =3514
                            Width =735
                            Height =240
                            Name ="Label29"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =3855
                    ColumnOrder =9
                    TabIndex =7
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2551
                            Top =3855
                            Width =1020
                            Height =240
                            Name ="Label30"
                            Caption ="مقدار كل كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =4195
                    ColumnWidth =450
                    ColumnOrder =2
                    TabIndex =8
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3181
                            Top =4195
                            Width =390
                            Height =240
                            Name ="Label31"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1360
                    Top =396
                    ColumnOrder =6
                    TabIndex =9
                    Name ="ekh"
                    ControlSource ="ekh"
                    Format ="#,###"
                    ConditionalFormat = Begin
                        0x010000006e000000020000000000000004000000000000000200000001000000 ,
                        0x00000000ff000000000000000500000003000000060000000100000000000000 ,
                        0xff00000000000000000000000000000000000000000000000000000000000000 ,
                        0x3100000000002d00310000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4365
                            Top =390
                            Width =555
                            Height =240
                            Name ="Label32"
                            Caption ="اختلاف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "hesheshes_mablk.cls"
