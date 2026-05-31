Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11310
    RowHeight =285
    ItemSuffix =24
    Left =1515
    Top =5595
    Right =15810
    Bottom =10020
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    AfterInsert ="[Event Procedure]"
    BeforeDelConfirm ="[Event Procedure]"
    OrderBy ="TDETA_HES.TNUMBER"
    RecSrcDt = Begin
        0x48bc4fa87975e240
    End
    RecordSource ="TDETA_HES"
    Caption ="DETA_HES"
    OnCurrent ="[Event Procedure]"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
            Height =6710
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =113
                    ColumnOrder =0
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4640
                            Top =113
                            Width =825
                            Height =240
                            Name ="Label1"
                            Caption ="حساب كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =453
                    ColumnWidth =600
                    ColumnOrder =1
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4775
                            Top =453
                            Width =690
                            Height =240
                            Name ="Label3"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1404
                    Top =793
                    Width =2190
                    ColumnWidth =3900
                    ColumnOrder =3
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4970
                            Top =793
                            Width =495
                            Height =240
                            Name ="Label5"
                            Caption ="نام حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    Width =3450
                    ColumnWidth =3195
                    ColumnOrder =5
                    TabIndex =3
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4925
                            Top =1133
                            Width =540
                            Height =240
                            Name ="Label7"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1014
                    Top =1814
                    Width =2580
                    ColumnWidth =3285
                    ColumnOrder =6
                    TabIndex =4
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4700
                            Top =1814
                            Width =765
                            Height =240
                            Name ="Label11"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1884
                    Top =2154
                    Width =1710
                    ColumnWidth =1417
                    ColumnOrder =4
                    TabIndex =5
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5135
                            Top =2154
                            Width =330
                            Height =240
                            Name ="Label13"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1814
                    Top =1530
                    ColumnWidth =1740
                    ColumnOrder =2
                    TabIndex =6
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4696
                            Top =1530
                            Width =690
                            Height =240
                            Name ="Label17"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1870
                    Top =2664
                    ColumnOrder =8
                    TabIndex =7
                    Name ="BED_BES"
                    ControlSource ="BED_BES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4072
                            Top =2664
                            Width =1200
                            Height =240
                            Name ="Label18"
                            Caption ="اتوماتيك"
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
                    Left =1190
                    Top =3061
                    Width =2691
                    Height =315
                    ColumnWidth =2190
                    ColumnOrder =9
                    TabIndex =8
                    ForeColor =8388608
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUSTKIND.CUST_COD, CUSTKIND.CUSTKNAME FROM CUSTKIND ORDER BY CUSTKIND.CUS"
                        "TKNAME;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3920
                            Top =3061
                            Width =1020
                            Height =300
                            ForeColor =8388608
                            Name ="Label113"
                            Caption ="نوع مشتري "
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
                    Left =2040
                    Top =4252
                    Width =2661
                    Height =300
                    ColumnOrder =10
                    TabIndex =9
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4801
                            Top =4251
                            Width =870
                            Height =240
                            Name ="Label16"
                            Caption ="كد اقتصادي"
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
                    Left =2040
                    Top =4592
                    Width =2661
                    Height =300
                    ColumnOrder =11
                    TabIndex =10
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4801
                            Top =4596
                            Width =735
                            Height =240
                            Name ="Label19"
                            Caption ="كد پستي"
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
                    Left =2040
                    Top =4932
                    Width =2661
                    Height =300
                    ColumnOrder =12
                    TabIndex =11
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4801
                            Top =4932
                            Width =660
                            Height =240
                            Name ="Label20"
                            Caption ="استان"
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
                    Left =2040
                    Top =5273
                    Width =2661
                    Height =300
                    ColumnOrder =13
                    TabIndex =12
                    Name ="CITY"
                    ControlSource ="CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4801
                            Top =5271
                            Width =930
                            Height =240
                            Name ="Label21"
                            Caption ="شهرستان"
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
                    Left =2040
                    Top =5613
                    Width =2661
                    Height =300
                    ColumnWidth =1920
                    ColumnOrder =7
                    TabIndex =13
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4801
                            Top =5613
                            Width =795
                            Height =240
                            Name ="Label22"
                            Caption ="كد ملي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =14
                    Width =8811
                    Height =300
                    ColumnWidth =1875
                    TabIndex =14
                    Name ="Text23"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =8908
                            Width =2175
                            Height =240
                            Name ="Label24"
                            Caption ="موبايل جهت ارسال پيامك"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2403
                    Top =6009
                    Width =2661
                    Height =300
                    TabIndex =15
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5171
                            Top =6014
                            Width =705
                            Height =240
                            Name ="Label38"
                            Caption ="کد استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2381
                    Top =6410
                    Width =2661
                    Height =300
                    TabIndex =16
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5149
                            Top =6404
                            Width =990
                            Height =240
                            Name ="Label40"
                            Caption ="کد شهرستان"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TDETA_HES_SUB.cls"
