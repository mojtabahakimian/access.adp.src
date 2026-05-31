Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4592
    RowHeight =300
    ItemSuffix =57
    Left =60
    Top =1425
    Right =28350
    Bottom =8820
    DatasheetGridlinesColor =12632256
    BeforeDelConfirm ="[Event Procedure]"
    RecSrcDt = Begin
        0x26a19156bfa6e240
    End
    RecordSource ="INVO_LST"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =7326
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
                    Left =238
                    Top =771
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2850
                            Top =771
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =238
                    Top =1111
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =1111
                            Width =630
                            Height =240
                            Name ="Label3"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =225
                    Top =2105
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3122
                            Top =2105
                            Width =675
                            Height =240
                            Name ="Label11"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =251
                    Top =2728
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =6
                    TabIndex =5
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2713
                            Top =2728
                            Width =1110
                            Height =240
                            Name ="Label13"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =251
                    Top =3068
                    Height =255
                    ColumnWidth =1515
                    ColumnOrder =18
                    TabIndex =6
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3405
                            Top =3075
                            Width =690
                            Height =240
                            Name ="Label15"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =251
                    Top =3408
                    Height =255
                    ColumnWidth =1110
                    ColumnOrder =8
                    TabIndex =7
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00;-#,##0.00"
                    OnEnter ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2818
                            Top =3411
                            Width =1245
                            Height =285
                            Name ="Label17"
                            Caption ="مبلغ ميانگين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =251
                    Top =3748
                    Height =255
                    ColumnWidth =1170
                    ColumnOrder =9
                    TabIndex =8
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,##0.00;-#,##0.00"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3208
                            Top =3748
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =251
                    Top =4088
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =10
                    TabIndex =9
                    Name ="FROM_A"
                    ControlSource ="FROM_A"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3118
                            Top =4088
                            Width =705
                            Height =240
                            Name ="Label21"
                            Caption ="FROM_A"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =253
                    Top =4478
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =11
                    TabIndex =10
                    Name ="MEGH_R"
                    ControlSource ="MEGH_R"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2955
                            Top =4478
                            Width =870
                            Height =240
                            Name ="Label25"
                            Caption ="مقدار رسيد"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =253
                    Top =4818
                    Height =255
                    ColumnWidth =210
                    ColumnOrder =12
                    TabIndex =11
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3540
                            Top =4818
                            Width =285
                            Height =240
                            Name ="Label27"
                            Caption ="رده"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =253
                    Top =5498
                    Height =255
                    ColumnWidth =1155
                    ColumnOrder =13
                    TabIndex =13
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2700
                            Top =5498
                            Width =1125
                            Height =240
                            Name ="Label31"
                            Caption ="شماره مشتري"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =253
                    Top =5838
                    Height =255
                    ColumnWidth =945
                    ColumnOrder =14
                    TabIndex =14
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3030
                            Top =5838
                            Width =795
                            Height =240
                            Name ="Label33"
                            Caption ="انبار فرعي"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =246
                    Top =1765
                    Width =1680
                    Height =255
                    ColumnWidth =12301
                    ColumnOrder =3
                    TabIndex =3
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;4536;567"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnKeyUp ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3287
                            Top =1765
                            Width =510
                            Height =240
                            Name ="Label9"
                            Caption =" كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =3
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =249
                    Top =2388
                    Height =255
                    ColumnWidth =840
                    ColumnOrder =7
                    TabIndex =16
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2631
                            Top =2388
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="مقدار كل "
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =253
                    Top =6179
                    Height =255
                    ColumnWidth =1425
                    ColumnOrder =4
                    TabIndex =15
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnEnter ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3210
                            Top =6179
                            Width =615
                            Height =240
                            Name ="Label35"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =253
                    Top =1451
                    Width =1686
                    Height =255
                    ColumnWidth =2175
                    ColumnOrder =2
                    TabIndex =2
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES FROM TCOD_ANBAR ORDER BY TCOD_ANBAR.COD"
                        "E; "
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3480
                            Top =1451
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =255
                    Top =120
                    Width =1680
                    Height =255
                    ColumnWidth =1440
                    ColumnOrder =15
                    TabIndex =17
                    Name ="CODEO"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =251
                    Top =6519
                    Height =255
                    ColumnWidth =4245
                    ColumnOrder =16
                    TabIndex =18
                    Name ="N_RASID"
                    ControlSource ="N_RASID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CAST(HEAD_MANF.FNUMB AS nvarchar) AS fnumb, ISNULL(HEAD_MANF.NAMES, STUF_"
                        "DEF.NAME) AS NAM, CAST(HEAD_MANF.FNUMB AS nvarchar) AS Expr1 FROM STUF_DEF RIGHT"
                        " OUTER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE LEFT OUTER JOIN DTL_MANF"
                        " ON HEAD_MANF.FNUMB = DTL_MANF.FNUMB WHERE (DTL_MANF.CODE IS NULL) UNION SELECT "
                        "hes, NAME AS nam, hes AS Expr1 FROM CUST_HESAb"
                    ColumnWidths ="0;3402;567"
                    ValidationRule ="Not Is Null"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2308
                            Top =6519
                            Width =945
                            Height =240
                            Name ="Label23"
                            Caption ="محل مصرف"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =566
                    ColumnWidth =135
                    ColumnOrder =17
                    TabIndex =19
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="=[RecordsetClone].[RecordCount]+1"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =253
                    Top =5158
                    ColumnWidth =1305
                    ColumnOrder =19
                    TabIndex =12
                    Name ="SANAD_NO"
                    ControlSource ="SANAD_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MHAZ_NO, MHAZNAME FROM TCOD_MARKAZHAZ"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2940
                            Top =5158
                            Width =885
                            Height =240
                            Name ="Label29"
                            Caption ="مركز هزينه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =623
                    Top =6859
                    TabIndex =20
                    Name ="AVRAGE"
                    ControlSource ="AVRAGE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3290
                            Top =6859
                            Width =735
                            Height =240
                            Name ="Label55"
                            Caption ="ميانگين"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =566
                    Top =7086
                    TabIndex =21
                    Name ="AVRAGE2"
                    ControlSource ="AVRAGE2"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3143
                            Top =7086
                            Width =825
                            Height =240
                            Name ="Label56"
                            Caption ="ميانگين"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =270
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =360
                    Top =15
                    Height =255
                    Name ="SMABLK"
                    ControlSource ="=Sum([MABL_K])"
                End
            End
        End
    End
End
CodeBehindForm
' See "HAVALAH_EXIT_SAYER_SUB.cls"
