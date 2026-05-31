Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15030
    RowHeight =420
    ItemSuffix =32
    Left =7875
    Top =1170
    Right =22950
    Bottom =10290
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x1775e7a6f291e440
    End
    RecordSource ="VISITORS_PORSANT"
    Caption ="تعيين پورسانت ويزيتور براي هر كالا"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8957
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7029
                    Top =963
                    Width =1266
                    Height =317
                    ColumnWidth =1980
                    ColumnOrder =7
                    TabIndex =2
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8391
                            Top =963
                            Width =750
                            Height =240
                            Name ="Label11"
                            Caption ="تا تاريخ :"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7377
                    Top =570
                    Width =1515
                    Height =317
                    ColumnOrder =5
                    TabIndex =3
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8952
                            Top =570
                            Width =345
                            Height =330
                            Name ="Label13"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2034
                    Top =113
                    Width =9531
                    Height =317
                    ColumnWidth =5325
                    ColumnOrder =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.HES, CUST_HESAB.NAME, CUST_HESAB.MOBILE FROM Visit_route INNE"
                        "R JOIN CUST_HESAB ON Visit_route.HES = CUST_HESAB.hes GROUP BY Visit_route.HES, "
                        "CUST_HESAB.NAME, CUST_HESAB.MOBILE"
                    ColumnWidths ="0;3969"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11649
                            Top =113
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label19"
                            Caption ="ويزيتور"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =120
                    Width =1926
                    Height =317
                    ColumnOrder =8
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES2"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT HES FROM Visit_route GROUP BY HES"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =600
                            Top =120
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =134
                    TextFontCharSet =0
                    Left =56
                    Top =967
                    Width =1020
                    Height =405
                    FontWeight =700
                    TabIndex =4
                    Name ="Command100"
                    Caption ="&چاپ "
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1670
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12980
                    Top =566
                    Width =1041
                    Height =317
                    TabIndex =5
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14115
                            Top =566
                            Width =870
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ :"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1247
                    Top =967
                    Width =4185
                    Height =405
                    FontWeight =700
                    TabIndex =6
                    Name ="Command23"
                    Caption ="دريافت ليست كالاها  از طريق گروه كالا (منو)"
                    OnClick ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6349
                    Top =963
                    Width =231
                    Height =255
                    TabIndex =7
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="True"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12472
                    Top =113
                    Width =1140
                    Height =317
                    TabIndex =8
                    Name ="PORID"
                    ControlSource ="PORID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13676
                            Top =170
                            Width =1290
                            Height =240
                            Name ="Label26"
                            Caption ="شناسه پورسانت:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =10020
                    Top =930
                    Width =5010
                    Height =240
                    ForeColor =128
                    Name ="Label27"
                    Caption =" تاريخ صرفا براي اطلاع است و تاثيري در محاسبه پورسانت ندارد"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =9418
                    Top =570
                    Width =2436
                    Height =318
                    TabIndex =9
                    BackColor =128
                    ForeColor =16777215
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND UNION SELECT 0 AS cust_cod, 'بدون گروه("
                        "همه)' AS custkname"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11879
                            Top =570
                            Width =1035
                            Height =318
                            Name ="Label28"
                            Caption ="گروه مشتري:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Top =570
                    Width =6336
                    Height =318
                    TabIndex =10
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6537
                            Top =570
                            Width =735
                            Height =318
                            Name ="Label29"
                            Caption ="توضيحات:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =113
                    Top =1587
                    Width =14686
                    Height =7145
                    TabIndex =11
                    Name ="VISITORS_PORSANT_KALAF"
                    SourceObject ="Form.VISITORS_PORSANT_KALAF"
                    LinkChildFields ="PORID"
                    LinkMasterFields ="PORID"
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_PORSANT_HEAD.cls"
