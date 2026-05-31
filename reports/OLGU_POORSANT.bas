Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9921
    DatasheetFontHeight =10
    ItemSuffix =28
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x624cbf62d2b0e440
    End
    RecordSource ="SELECT VISITORS_PORSANT.PORID, VISITORS_PORSANT.HES, VISITORS_PORSANT.VDATE, CUS"
        "T_HESAB.NAME, VISITORS_PORSANT.CUST_COD, STUF_DEF.NAME AS KALA, STUF_DEF.CODE, V"
        "ISITORS_PORSANT_KALA.PORSANT, VISITORS_PORSANT.CDATE, VISITORS_PORSANT.USERNAME,"
        " VISITORS_PORSANT.COMMENT FROM VISITORS_PORSANT_KALA INNER JOIN VISITORS_PORSANT"
        " ON VISITORS_PORSANT_KALA.PORID = VISITORS_PORSANT.PORID INNER JOIN CUST_HESAB O"
        "N VISITORS_PORSANT.HES = CUST_HESAB.hes INNER JOIN STUF_DEF ON VISITORS_PORSANT_"
        "KALA.CODE = STUF_DEF.CODE"
    Caption ="چاپ الگوي پ رداخت پورسان"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1671
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7994
                    Top =57
                    Width =906
                    Name ="PORID"
                    ControlSource ="PORID"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =8946
                            Top =56
                            Width =885
                            Height =240
                            Name ="Label0"
                            Caption ="شماره الگو:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7
                    Top =435
                    Width =1371
                    TabIndex =1
                    Name ="HES"
                    ControlSource ="HES"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =1414
                            Top =442
                            Width =765
                            Height =240
                            Name ="Label1"
                            Caption ="كد ويزيتور:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =502
                    Top =60
                    Width =861
                    TabIndex =2
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =34
                            Left =1414
                            Top =60
                            Width =510
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2384
                    Top =453
                    Width =6516
                    TabIndex =3
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =8995
                            Top =458
                            Width =810
                            Height =240
                            Name ="Label3"
                            Caption ="نام ويزيتور:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5617
                    Top =765
                    Width =1140
                    Height =317
                    TabIndex =4
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =34
                            Left =6817
                            Top =765
                            Width =345
                            Height =330
                            Name ="Label13"
                            Caption ="كاربر:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7201
                    Top =767
                    Width =1551
                    Height =318
                    TabIndex =5
                    BackColor =128
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND UNION SELECT 0 AS cust_cod, 'بدون گروه("
                        "همه)' AS custkname"
                    ColumnWidths ="0"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =34
                            Left =8777
                            Top =767
                            Width =1035
                            Height =318
                            Name ="Label28"
                            Caption ="گروه مشتري:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7
                    Top =735
                    Width =4806
                    Height =318
                    TabIndex =6
                    Name ="COMMENT"
                    ControlSource ="COMMENT"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =34
                            Left =4819
                            Top =735
                            Width =735
                            Height =318
                            Name ="Label29"
                            Caption ="توضيحات:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =4417
                    Width =1365
                    Height =465
                    FontSize =12
                    FontWeight =700
                    Name ="Label9"
                    Caption ="الگوي پورسانت"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =34
                    Left =9405
                    Top =1365
                    Width =405
                    Height =240
                    Name ="Label15"
                    Caption ="رديف"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =34
                    Left =8550
                    Top =1365
                    Width =495
                    Height =240
                    Name ="Label16"
                    Caption ="كد كالا"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =34
                    Left =4395
                    Top =1365
                    Width =525
                    Height =240
                    Name ="Label17"
                    Caption ="نام كالا"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =34
                    Left =60
                    Top =1365
                    Width =555
                    Height =240
                    Name ="Label18"
                    Caption ="درصد"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =9270
                    Top =1275
                    Width =0
                    Height =396
                    Name ="Line19"
                End
                Begin Line
                    Left =8325
                    Top =1275
                    Width =0
                    Height =396
                    Name ="Line20"
                End
                Begin Line
                    Left =907
                    Top =1275
                    Width =0
                    Height =396
                    Name ="Line21"
                End
                Begin Line
                    Top =1275
                    Width =0
                    Height =396
                    Name ="Line22"
                End
                Begin Line
                    Left =9870
                    Top =1275
                    Width =0
                    Height =396
                    Name ="Line23"
                End
                Begin Line
                    BorderWidth =1
                    Top =1260
                    Width =9879
                    Name ="Line24"
                End
                Begin Line
                    BorderWidth =1
                    Top =1650
                    Width =9879
                    Name ="Line25"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =396
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8395
                    Width =846
                    Height =345
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =801
                    Height =345
                    TabIndex =1
                    Name ="PORSANT"
                    ControlSource ="PORSANT"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1033
                    Width =7236
                    Height =345
                    TabIndex =2
                    Name ="Text6"
                    ControlSource ="KALA"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9298
                    Width =561
                    Height =345
                    TabIndex =3
                    Name ="Text8"
                    ControlSource ="=1"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =9270
                    Width =0
                    Height =396
                    Name ="Line10"
                End
                Begin Line
                    Left =8325
                    Width =0
                    Height =396
                    Name ="Line11"
                End
                Begin Line
                    Left =907
                    Width =0
                    Height =396
                    Name ="Line12"
                End
                Begin Line
                    Width =0
                    Height =396
                    Name ="Line13"
                End
                Begin Line
                    Left =9870
                    Width =0
                    Height =396
                    Name ="Line14"
                End
            End
        End
        Begin PageFooter
            Height =340
            Name ="PageFooterSection"
            Begin
                Begin Line
                    BorderWidth =1
                    Width =9879
                    Name ="Line26"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =113
            Name ="ReportFooter"
            Begin
                Begin Line
                    BorderWidth =1
                    Width =9879
                    Name ="Line27"
                End
            End
        End
    End
End
