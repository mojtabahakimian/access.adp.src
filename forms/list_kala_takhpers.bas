Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5961
    DatasheetFontHeight =10
    ItemSuffix =35
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x120b193eb329e540
    End
    RecordSource ="SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, dbo.STUF_D"
        "EF.TOZIH, dbo.STUF_DEF.VAHED, dbo.STUF_DEF.B_SEF, \015\012                      "
        "dbo.STUF_DEF.N_SEF, dbo.STUF_DEF.MIN_M, dbo.STUF_DEF.MAX_M, dbo.STUF_DEF.RADAH, "
        "dbo.STUF_DEF.KINDK, dbo.STUF_DEF.MABL_F, \015\012                      dbo.STUF_"
        "DEF.DEPART, dbo.STUF_DEF.IDD, dbo.STUF_DEF.CMBAA, dbo.STUF_DEF.VAZN, dbo.STUF_DE"
        "F.OKF, dbo.STUF_DEF.MENUIT, \015\012                      dbo.STUF_DEF.MEGHTA, d"
        "bo.STUF_DEF.MEGHJAY, dbo.TAKHPERS.TAKH_COD, dbo.TAKHPERS.CUST_CO, dbo.TAKHPERS.T"
        "AFPER, \015\012                      dbo.TAKHPERS.PRICE_M, dbo.TAKHPERS.PERS, db"
        "o.TAKHPERS.BLNS, dbo.TAKHPERS.PUT\015\012FROM         dbo.STUF_DEF INNER JOIN\015"
        "\012                      dbo.TAKHPERS ON dbo.STUF_DEF.CODE = dbo.TAKHPERS.TAKH_"
        "COD"
    Caption ="ليست کالا و تخفيفات"
    DatasheetFontName ="WeblogmaYekan"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Section
            Height =10611
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =793
                    ColumnWidth =1065
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =793
                            Width =555
                            Height =240
                            Name ="Label0"
                            Caption ="کد کالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =1133
                    ColumnWidth =2835
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =1133
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="نام کالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =1474
                    TabIndex =2
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =1470
                            Width =915
                            Height =240
                            Name ="Label2"
                            Caption ="شماره فني "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =1814
                    TabIndex =3
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =1814
                            Width =600
                            Height =240
                            Name ="Label3"
                            Caption ="توضيح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =2494
                    TabIndex =4
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2550
                            Top =2494
                            Width =585
                            Height =240
                            Name ="Label5"
                            Caption ="في خرده فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =2834
                    ColumnWidth =1065
                    TabIndex =5
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2550
                            Top =2834
                            Width =600
                            Height =240
                            Name ="Label6"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =3174
                    ColumnWidth =1155
                    TabIndex =6
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =3174
                            Width =630
                            Height =240
                            Name ="Label7"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =3514
                    TabIndex =7
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =3514
                            Width =660
                            Height =240
                            Name ="Label8"
                            Caption ="قيمت مصرف كننده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =4195
                    ColumnWidth =405
                    TabIndex =8
                    Name ="KINDK"
                    ControlSource ="KINDK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =4195
                            Width =360
                            Height =240
                            Name ="Label10"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =4535
                    ColumnWidth =1185
                    TabIndex =9
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =4535
                            Width =705
                            Height =240
                            Name ="Label11"
                            Caption ="في عمده فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =4875
                    ColumnWidth =1245
                    TabIndex =10
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =4875
                            Width =1185
                            Height =240
                            Name ="Label12"
                            Caption ="واحد ارائه كننده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =5896
                    TabIndex =11
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =5896
                            Width =525
                            Height =240
                            Name ="Label15"
                            Caption ="وزن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =6576
                    TabIndex =12
                    Name ="MENUIT"
                    ControlSource ="MENUIT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =6576
                            Width =705
                            Height =240
                            Name ="Label17"
                            Caption ="زيرمجموعه منوي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =6916
                    ColumnWidth =885
                    TabIndex =13
                    Name ="MEGHTA"
                    ControlSource ="MEGHTA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =6916
                            Width =750
                            Height =240
                            Name ="Label18"
                            Caption ="به ازاي هر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4251
                    Top =7256
                    ColumnWidth =525
                    TabIndex =14
                    Name ="MEGHJAY"
                    ControlSource ="MEGHJAY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2550
                            Top =7256
                            Width =825
                            Height =240
                            Name ="Label19"
                            Caption ="جايزه"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3411
                    Top =7650
                    Width =1005
                    ColumnWidth =1215
                    TabIndex =15
                    Name ="TAKH_COD"
                    ControlSource ="TAKH_COD"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4502
                            Top =7650
                            Width =1140
                            Height =240
                            Name ="Label27"
                            Caption ="كد كالا و خدمات"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2070
                    Top =8330
                    Width =2346
                    ColumnWidth =795
                    TabIndex =16
                    Name ="TAFPER"
                    ControlSource ="TAFPER"
                    InputMask ="\\%###"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4712
                            Top =8330
                            Width =930
                            Height =240
                            Name ="Label28"
                            Caption ="% تخفيف"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2070
                    Top =7990
                    Width =2346
                    ColumnWidth =1320
                    TabIndex =17
                    Name ="CUST_CO"
                    ControlSource ="CUST_CO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4532
                            Top =7990
                            Width =1110
                            Height =240
                            Name ="Label29"
                            Caption ="كد نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2606
                    Top =8784
                    TabIndex =18
                    Name ="PRICE_M"
                    ControlSource ="PRICE_M"
                    Format ="#,##0.00;-#,##0.00"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5104
                            Top =8784
                            Width =795
                            Height =240
                            Name ="Label30"
                            Caption ="قيمت"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2492
                    Top =9577
                    ColumnWidth =780
                    TabIndex =19
                    Name ="PERS"
                    ControlSource ="PERS"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5391
                            Top =9577
                            Width =555
                            Height =405
                            Name ="Label31"
                            Caption ="درصد +-"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2492
                    Top =9918
                    ColumnWidth =795
                    TabIndex =20
                    Name ="BLNS"
                    ControlSource ="BLNS"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5406
                            Top =9922
                            Width =555
                            Height =405
                            Name ="Label32"
                            Caption ="بالانس+-"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3083
                    Top =10371
                    ColumnWidth =1455
                    TabIndex =21
                    Name ="PUT"
                    ControlSource ="PUT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3576
                            Top =10342
                            Width =1485
                            Height =240
                            Name ="Label33"
                            Caption ="قيمت گذاري گروهي"
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1912
                    Top =2085
                    Width =2805
                    Height =315
                    TabIndex =22
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS ORDER BY TCOD_VAHEDS"
                        ".NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4869
                            Top =2085
                            Width =465
                            Height =300
                            Name ="Label9"
                            Caption ="واحد:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Top =2652
                    Width =4701
                    Height =315
                    ColumnWidth =1545
                    TabIndex =23
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP WHERE (((TC"
                        "OD_STUFGROUP.CODE)<>0)) ORDER BY TCOD_STUFGROUP.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =4869
                            Top =2652
                            Width =795
                            Height =300
                            Name ="Label34"
                            Caption ="گروه كالا:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Top =45
                    Width =283
                    Height =245
                    TabIndex =24
                    Name ="CMBAA"
                    ControlSource ="CMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            Left =273
                            Top =22
                            Width =1485
                            Height =240
                            ForeColor =255
                            Name ="Label48"
                            Caption ="مشمول ماليات ب.ا.ا"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1800
                    Top =54
                    Width =340
                    Height =218
                    ColumnWidth =540
                    TabIndex =25
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2190
                            Width =405
                            Height =270
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد "
                        End
                    End
                End
            End
        End
    End
End
