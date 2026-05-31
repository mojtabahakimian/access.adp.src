Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14513
    RowHeight =315
    ItemSuffix =20
    Left =2160
    Top =2130
    Right =17235
    Bottom =11220
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xc78422d1775fe640
    End
    RecordSource ="SELECT CUST_HESAB.NAME, STUF_DEF.NAME AS kala, INVO_LST.TAG, TAGCOD.BARGAH, INVO"
        "_LST.CODE, HEAD_LST.CUST_NO, SUM(INVO_LST.MEGH_MAR) AS mar, SUM(INVO_LST.MEGHk) "
        "AS mrgh FROM INVO_LST INNER JOIN HEAD_LST ON INVO_LST.NUMBER = HEAD_LST.NUMBER A"
        "ND INVO_LST.TAG = HEAD_LST.TAG INNER JOIN CUST_HESAB ON HEAD_LST.CUST_NO = CUST_"
        "HESAB.hes INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE INNER JOIN TAGCOD"
        " ON INVO_LST.TAG = TAGCOD.CODE WHERE (HEAD_LST.DATE_N > 870611) AND (HEAD_LST.DA"
        "TE_N < 99999999) AND (INVO_LST.CODE = N'2952' OR INVO_LST.CODE = N'3357 ' OR INV"
        "O_LST.CODE = N'3359 ' OR INVO_LST.CODE = N'3358 ') GROUP BY INVO_LST.CODE, HEAD_"
        "LST.CUST_NO, CUST_HESAB.NAME, INVO_LST.TAG, STUF_DEF.NAME, TAGCOD.BARGAH ORDER B"
        "Y CUST_HESAB.NAME, STUF_DEF.NAME, INVO_LST.TAG"
    Caption ="ليست"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin Tab
            TextFontCharSet =178
            BackStyle =0
            Width =5103
            Height =3402
        End
        Begin FormHeader
            Height =896
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =11781
                    Top =348
                    Width =345
                    Height =315
                    Name ="NAME_Label"
                    Caption ="نام "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =6270
                    Top =348
                    Width =405
                    Height =315
                    BackColor =65280
                    Name ="kala_Label"
                    Caption ="جنس"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =8071
                    Top =348
                    Width =720
                    Height =315
                    BackColor =13434828
                    Name ="BARGAH_Label"
                    Caption ="نوع بركه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =2643
                    Top =345
                    Width =525
                    Height =315
                    Name ="CODE_Label"
                    Caption ="كد كالا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =13241
                    Top =345
                    Width =1050
                    Height =315
                    Name ="CUST_NO_Label"
                    Caption ="شماره مشتري"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =3567
                    Top =348
                    Width =600
                    Height =315
                    Name ="mar_Label"
                    Caption ="مرجوعي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =4666
                    Top =345
                    Width =465
                    Height =315
                    BackColor =16777164
                    ForeColor =0
                    Name ="mrgh_Label"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =855
                    Top =56
                    Width =1309
                    Height =330
                    Name ="DT1"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="13721102"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            TextFontFamily =2
                            Left =2230
                            Top =56
                            Width =330
                            Height =315
                            Name ="Label17"
                            Caption ="از :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =858
                    Top =465
                    Width =1309
                    Height =330
                    TabIndex =1
                    Name ="DT2"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=FARSIDATE(Date())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            TextFontFamily =2
                            Left =2233
                            Top =465
                            Width =300
                            Height =315
                            Name ="Label19"
                            Caption ="تا :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =417
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8956
                    Top =57
                    Width =4317
                    Height =330
                    ColumnWidth =3060
                    ColumnOrder =1
                    Name ="نام"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5448
                    Top =57
                    Width =1962
                    Height =330
                    ColumnWidth =1305
                    ColumnOrder =2
                    TabIndex =1
                    BackColor =65280
                    Name ="جنس"
                    ControlSource ="kala"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7465
                    Top =57
                    Width =1437
                    Height =330
                    ColumnWidth =1350
                    ColumnOrder =3
                    TabIndex =2
                    BackColor =13434828
                    Name ="نوع بركه"
                    ControlSource ="BARGAH"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="نوع_بركه"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2568
                    Top =57
                    Width =687
                    Height =330
                    ColumnWidth =870
                    ColumnOrder =4
                    TabIndex =3
                    Name ="كد كالا"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="كد_كالا"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =13316
                    Top =57
                    Width =1197
                    Height =330
                    ColumnWidth =1230
                    ColumnOrder =0
                    TabIndex =4
                    Name ="شماره مشتري"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="شماره_مشتري"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3407
                    Top =57
                    Width =912
                    Height =330
                    ColumnWidth =795
                    TabIndex =5
                    Name ="مرجوعي"
                    ControlSource ="mar"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4376
                    Top =57
                    Width =1023
                    Height =330
                    ColumnWidth =840
                    TabIndex =6
                    BackColor =16777164
                    ForeColor =0
                    Name ="مقدار"
                    ControlSource ="mrgh"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "LIST_KALA.cls"
