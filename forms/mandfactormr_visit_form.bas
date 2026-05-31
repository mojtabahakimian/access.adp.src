Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =345
    DatasheetFontHeight =10
    ItemSuffix =107
    Left =795
    Top =4035
    Right =2220
    Bottom =5445
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa722150653e9e540
    End
    RecordSource ="mandfactormr_visit"
    Caption ="ليست ويزيتورها"
    DatasheetFontName ="Arial"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            Height =10034
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =975
                    ColumnOrder =13
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =930
                            Height =240
                            Name ="Label1"
                            Caption =" حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2475
                    ColumnWidth =1140
                    ColumnOrder =1
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =660
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    ColumnWidth =480
                    ColumnOrder =14
                    TabIndex =2
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =885
                            Height =240
                            Name ="Label5"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Width =2475
                    ColumnOrder =15
                    TabIndex =3
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =975
                            Height =240
                            Name ="Label7"
                            Caption ="کاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =510
                    ColumnOrder =17
                    TabIndex =4
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =540
                            Height =240
                            Name ="Label9"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =420
                    ColumnOrder =18
                    TabIndex =5
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =375
                            Height =240
                            Name ="Label11"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2154
                    ColumnOrder =11
                    TabIndex =6
                    Name ="SMBAA"
                    ControlSource ="SMBAA"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2160
                            Width =900
                            Height =240
                            Name ="Label13"
                            Caption ="ارزش افزوده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    Width =2475
                    ColumnOrder =19
                    TabIndex =7
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2490
                            Width =780
                            Height =240
                            Name ="Label15"
                            Caption ="ملاحضلات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =4195
                    ColumnWidth =630
                    ColumnOrder =0
                    TabIndex =8
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =4200
                            Width =960
                            Height =240
                            Name ="Label25"
                            Caption =" فاکتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =5215
                    ColumnWidth =450
                    ColumnOrder =20
                    TabIndex =9
                    Name ="MAS"
                    ControlSource ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5215
                            Width =390
                            Height =240
                            Name ="Label31"
                            Caption ="مدت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =5896
                    ColumnWidth =390
                    ColumnOrder =21
                    TabIndex =10
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =5895
                            Width =390
                            Height =240
                            Name ="Label35"
                            Caption ="سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =6236
                    Width =2475
                    ColumnOrder =12
                    TabIndex =11
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =6240
                            Width =840
                            Height =240
                            Name ="Label37"
                            Caption ="کد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =6576
                    ColumnOrder =6
                    TabIndex =12
                    Name ="M_NAGHD"
                    ControlSource ="M_NAGHD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =6576
                            Width =810
                            Height =240
                            Name ="Label39"
                            Caption ="نقد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =6916
                    ColumnOrder =7
                    TabIndex =13
                    Name ="MABL_VAR"
                    ControlSource ="MABL_VAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =6916
                            Width =855
                            Height =240
                            Name ="Label41"
                            Caption ="واريزي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =7256
                    ColumnOrder =8
                    TabIndex =14
                    Name ="MABL_HAV"
                    ControlSource ="MABL_HAV"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =7256
                            Width =855
                            Height =240
                            Name ="Label43"
                            Caption ="حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =7596
                    ColumnOrder =9
                    TabIndex =15
                    Name ="MABL_HAZ"
                    ControlSource ="MABL_HAZ"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =7596
                            Width =855
                            Height =240
                            Name ="Label45"
                            Caption ="هزينه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =7993
                    ColumnOrder =10
                    TabIndex =16
                    Name ="TAKHFIF"
                    ControlSource ="TAKHFIF"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =7993
                            Width =705
                            Height =240
                            Name ="Label47"
                            Caption ="تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1977
                    Top =2948
                    ColumnWidth =1245
                    ColumnOrder =4
                    TabIndex =17
                    Name ="mab"
                    ControlSource ="mab"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =2948
                            Width =420
                            Height =240
                            Name ="Label97"
                            Caption ="مبلغ "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1977
                    Top =3289
                    ColumnOrder =5
                    TabIndex =18
                    Name ="mandah"
                    ControlSource ="mandah"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =3293
                            Width =855
                            Height =240
                            Name ="Label99"
                            Caption ="مانده فاکتور"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1497
                    Top =3798
                    ColumnWidth =900
                    ColumnOrder =16
                    TabIndex =19
                    Name ="TICMBAA1"
                    ControlSource ="TICMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =747
                            Top =3768
                            Width =780
                            Height =240
                            Name ="Label103"
                            Caption ="تيک ارزش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =226
                    Top =8560
                    ColumnWidth =2055
                    ColumnOrder =3
                    TabIndex =20
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3075
                            Top =8565
                            Width =750
                            Height =240
                            Name ="Label104"
                            Caption ="نام ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1247
                    Top =9070
                    ColumnWidth =1035
                    ColumnOrder =2
                    TabIndex =21
                    Name ="Expr1"
                    ControlSource ="Expr1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4095
                            Top =9075
                            Width =720
                            Height =240
                            Name ="Label105"
                            Caption ="کد ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =169
                    Top =4762
                    ColumnWidth =420
                    TabIndex =22
                    Name ="TAG2"
                    ControlSource ="TAG2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3046
                            Top =4762
                            Width =525
                            Height =240
                            Name ="Label106"
                            Caption ="نوع"
                        End
                    End
                End
            End
        End
    End
End
