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
    Width =9240
    DatasheetFontHeight =10
    ItemSuffix =62
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x0d653130d4b0e440
    End
    RecordSource ="SELECT visitgol_head.HES, visitgol_head.MAH, visitgol_head.CDATE, visitgol_head."
        "OKF, visitgol_head.FDATE, visitgol_head.TODATE, visitgol_head.USERNAME AS visitg"
        "ol_head_USERNAME, visitgol_dtl.HES AS visitgol_dtl_HES, visitgol_dtl.MAH AS visi"
        "tgol_dtl_MAH, visitgol_dtl.CODE, visitgol_dtl.CDATE AS visitgol_dtl_CDATE, visit"
        "gol_dtl.RADIF, visitgol_dtl.MEGH, visitgol_dtl.MEGHk, visitgol_dtl.VAHED_K, visi"
        "tgol_dtl.USERNAME AS visitgol_dtl_USERNAME, STUF_DEF.NAME AS KALA, TCOD_VAHEDS.N"
        "AMES AS VAHED FROM visitgol_head INNER JOIN visitgol_dtl ON visitgol_head.HES = "
        "visitgol_dtl.HES AND visitgol_head.MAH = visitgol_dtl.MAH INNER JOIN STUF_DEF ON"
        " visitgol_dtl.CODE = STUF_DEF.CODE INNER JOIN TCOD_VAHEDS ON visitgol_dtl.VAHED_"
        "K = TCOD_VAHEDS.CODE"
    Caption ="اهداف تعريف شده براي ويزيتور"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontSize =10
            FontWeight =700
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =2
            Width =850
            Height =850
            BorderColor =12632256
        End
        Begin Line
            Width =1701
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Times New Roman"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            ControlSource ="CODE"
        End
        Begin PageHeader
            Height =1240
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    Left =7747
                    Top =916
                    Width =798
                    Height =285
                    FontSize =8
                    FontWeight =900
                    Name ="CODE_Label"
                    Caption ="كد كالا"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =8637
                    Top =916
                    Width =558
                    Height =285
                    FontSize =8
                    FontWeight =900
                    Name ="RADIF_Label"
                    Caption ="رديف"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2218
                    Top =916
                    Width =798
                    Height =285
                    FontSize =8
                    FontWeight =900
                    Name ="MEGH_Label"
                    Caption ="مقدار"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =120
                    Top =916
                    Width =798
                    Height =285
                    FontSize =8
                    FontWeight =900
                    Name ="MEGHk_Label"
                    Caption ="مقدار كل"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2835
                    Top =510
                    Width =5736
                    Height =317
                    FontSize =8
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.HES, CUST_HESAB.NAME, CUST_HESAB.hes AS Expr1 FROM Visit_rout"
                        "e INNER JOIN CUST_HESAB ON Visit_route.HES = CUST_HESAB.hes GROUP BY Visit_route"
                        ".HES, CUST_HESAB.NAME, CUST_HESAB.hes ORDER BY CUST_HESAB.NAME"
                    ColumnWidths ="0;3969"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =8655
                            Top =510
                            Width =540
                            Height =285
                            FontSize =8
                            FontWeight =400
                            ForeColor =8388608
                            Name ="Label19"
                            Caption ="ويزيتور:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =510
                    Width =2046
                    Height =317
                    FontSize =8
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES2"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT HES FROM Visit_route GROUP BY HES ORDER BY HES"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =2070
                            Top =510
                            Width =765
                            Height =405
                            FontSize =8
                            FontWeight =400
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد ويزيتور:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7095
                    Top =60
                    Height =330
                    FontSize =8
                    TabIndex =2
                    Name ="MAH"
                    ControlSource ="MAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="MON"
                    ColumnWidths ="0"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =0
                            Left =8848
                            Top =60
                            Width =345
                            Height =240
                            FontSize =8
                            FontWeight =400
                            Name ="Label21"
                            Caption ="ماه:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3798
                    Width =1710
                    Height =435
                    FontSize =12
                    FontWeight =400
                    Name ="Label41"
                    Caption ="اهداف تعريف شده"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    Left =4792
                    Top =916
                    Width =825
                    Height =285
                    FontSize =8
                    FontWeight =900
                    Name ="Label43"
                    Caption ="كالا"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1177
                    Top =915
                    Width =825
                    Height =285
                    FontSize =8
                    FontWeight =900
                    Name ="Label44"
                    Caption ="واحد"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8560
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line53"
                End
                Begin Line
                    Left =7740
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line54"
                End
                Begin Line
                    Left =2970
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line55"
                End
                Begin Line
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line56"
                End
                Begin Line
                    Left =960
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line57"
                End
                Begin Line
                    Left =2085
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line58"
                End
                Begin Line
                    Top =1230
                    Width =9240
                    Name ="Line59"
                End
                Begin Line
                    Left =9240
                    Top =855
                    Width =0
                    Height =385
                    Name ="Line60"
                End
                Begin Line
                    Top =855
                    Width =9240
                    Name ="Line61"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =385
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7773
                    Top =31
                    Width =798
                    Height =330
                    FontSize =8
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8630
                    Top =31
                    Width =558
                    Height =330
                    FontSize =8
                    TabIndex =1
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2154
                    Top =31
                    Width =798
                    Height =330
                    FontSize =8
                    TabIndex =2
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =120
                    Top =31
                    Width =798
                    Height =330
                    FontSize =8
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1020
                    Top =31
                    Width =1017
                    Height =330
                    FontSize =8
                    TabIndex =4
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =3000
                    Top =30
                    Width =4713
                    Height =330
                    FontSize =8
                    TabIndex =5
                    Name ="KALA"
                    ControlSource ="KALA"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =8560
                    Width =0
                    Height =385
                    Name ="Line45"
                End
                Begin Line
                    Left =7740
                    Width =0
                    Height =385
                    Name ="Line46"
                End
                Begin Line
                    Left =2970
                    Width =0
                    Height =385
                    Name ="Line47"
                End
                Begin Line
                    Width =0
                    Height =385
                    Name ="Line48"
                End
                Begin Line
                    Left =960
                    Width =0
                    Height =385
                    Name ="Line49"
                End
                Begin Line
                    Left =2085
                    Width =0
                    Height =385
                    Name ="Line50"
                End
                Begin Line
                    Top =375
                    Width =9240
                    Name ="Line51"
                End
                Begin Line
                    Left =9240
                    Width =0
                    Height =385
                    Name ="Line52"
                End
            End
        End
        Begin PageFooter
            Height =513
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Left =4665
                    Top =60
                    Width =4552
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text33"
                    ControlSource ="=FARSIDATE(Date())"
                    FontName ="Arial"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    BackStyle =1
                    IMESentenceMode =3
                    Top =56
                    Width =3472
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Text34"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Arial"
                End
            End
        End
    End
End
