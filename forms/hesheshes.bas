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
    ItemSuffix =23
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x483692bef3a4e540
    End
    RecordSource ="SELECT     TOP(100) PERCENT HES_K, HES_M, HES_T, HES, HES_T2, HES_T3, HES_T4, RT"
        "RIM(CAST(HES_K AS nvarchar)) + '-' + RTRIM(CAST(HES_M AS nvarchar))             "
        "           + '-' + RTRIM(CAST(HES_T AS nvarchar)) + CASE WHEN HES_T2 IS NULL THE"
        "N '' ELSE '-' + RTRIM(CAST(HES_T2 AS nvarchar)) END + CASE WHEN HES_T3 IS NULL  "
        "                      THEN '' ELSE '-' + RTRIM(CAST(HES_T3 AS nvarchar)) END AS "
        "hc, N_S, BED, BES, sharh FROM         dbo.DEED_DTL WHERE     (RTRIM(CAST(HES_K A"
        "S nvarchar)) + '-' + RTRIM(CAST(HES_M AS nvarchar)) + '-' + RTRIM(CAST(HES_T AS "
        "nvarchar)) + CASE WHEN HES_T2 IS NULL                        THEN '' ELSE '-' + "
        "RTRIM(CAST(HES_T2 AS nvarchar)) END + CASE WHEN HES_T3 IS NULL THEN '' ELSE '-' "
        "+ RTRIM(CAST(HES_T3 AS nvarchar)) END <> HES) ORDER BY HES_K"
    Caption ="حسابهايي که با کد حساب مغايرت دارند"
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
            Height =4208
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =495
                    ColumnOrder =0
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="کل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =600
                    ColumnOrder =1
                    TabIndex =1
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =570
                            Height =240
                            Name ="Label3"
                            Caption ="معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =735
                    ColumnOrder =2
                    TabIndex =2
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =540
                            Height =240
                            Name ="Label5"
                            Caption ="تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Width =2475
                    ColumnWidth =1800
                    ColumnOrder =6
                    TabIndex =3
                    Name ="HES"
                    ControlSource ="HES"
                    ConditionalFormat = Begin
                        0x0100000070000000010000000000000003000000000000000700000001000000 ,
                        0x00000000ffff9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e0075006c006c00220000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =360
                            Height =240
                            Name ="Label7"
                            Caption ="کد حساب وارد شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =870
                    ColumnOrder =3
                    TabIndex =4
                    Name ="HES_T2"
                    ControlSource ="HES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =630
                            Height =240
                            Name ="Label9"
                            Caption ="تفصيلي 2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =870
                    ColumnOrder =4
                    TabIndex =5
                    Name ="HES_T3"
                    ControlSource ="HES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =630
                            Height =240
                            Name ="Label11"
                            Caption ="تفصيلي 3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    ColumnWidth =870
                    ColumnOrder =5
                    TabIndex =6
                    Name ="HES_T4"
                    ControlSource ="HES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2154
                            Width =630
                            Height =240
                            Name ="Label13"
                            Caption ="تفصيلي 4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    Width =2475
                    ColumnWidth =1665
                    ColumnOrder =7
                    TabIndex =7
                    Name ="hc"
                    ControlSource ="hc"
                    ConditionalFormat = Begin
                        0x0100000070000000010000000000000003000000000000000700000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x22004e0075006c006c00220000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =240
                            Height =240
                            Name ="Label15"
                            Caption ="کد حساب ترکيبي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnWidth =990
                    ColumnOrder =8
                    TabIndex =8
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =360
                            Height =240
                            Name ="Label17"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3174
                    ColumnOrder =10
                    TabIndex =9
                    Name ="BED"
                    ControlSource ="BED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3174
                            Width =360
                            Height =240
                            Name ="Label19"
                            Caption ="بدهکار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =3514
                    ColumnOrder =11
                    TabIndex =10
                    Name ="BES"
                    ControlSource ="BES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3514
                            Width =345
                            Height =240
                            Name ="Label21"
                            Caption ="بستانکاري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =850
                    Top =3968
                    ColumnWidth =6705
                    ColumnOrder =9
                    TabIndex =11
                    Name ="sharh"
                    ControlSource ="sharh"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3772
                            Top =3968
                            Width =480
                            Height =240
                            Name ="Label22"
                            Caption ="شرح"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "hesheshes.cls"
