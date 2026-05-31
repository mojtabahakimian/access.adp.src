Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8524
    RowHeight =315
    ItemSuffix =34
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x5c923e5f0ca3e340
    End
    RecordSource ="SELECT ORDR_HED.id AS ORDR_HED_id, ORDR_HED.DATE AS ORDR_HED_DATE, ORDR_HED.MOLA"
        "H, ORDR_HED.N_S, ORDR_HED.CUST_NO AS ORDR_HED_CUST_NO, ORDR_LST.CODE, ORDR_LST.D"
        "ATE AS ORDR_LST_DATE, ORDR_LST.RADIF, ORDR_LST.MEGH, ORDR_LST.MEGHK, ORDR_LST.VA"
        "HED_K, ORDR_LST.CUST_NO AS ORDR_LST_CUST_NO, ORDR_HED.USER_NAME FROM ORDR_HED IN"
        "NER JOIN ORDR_LST ON ORDR_HED.id = ORDR_LST.ID"
    Caption ="گزارش سفارشات"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4848
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4186
                    Top =7
                    Width =1386
                    Height =315
                    ColumnWidth =1020
                    ColumnOrder =1
                    BackColor =16777215
                    ForeColor =0
                    Name ="ORDR_HED_DATE"
                    ControlSource ="ORDR_HED_DATE"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="FARSIDATE(Date())"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5673
                            Top =7
                            Width =540
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label1"
                            Caption ="تاريخ"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5
                    Top =359
                    Width =7620
                    Height =315
                    ColumnOrder =3
                    TabIndex =1
                    BackColor =16777215
                    ForeColor =0
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7726
                            Top =359
                            Width =795
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label3"
                            Caption ="ملاحظات"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6286
                    Top =7
                    Width =1026
                    Height =285
                    ColumnWidth =675
                    ColumnOrder =0
                    TabIndex =2
                    BackColor =16777215
                    ForeColor =0
                    Name ="ORDR_HED_id"
                    ControlSource ="ORDR_HED_id"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7339
                            Top =11
                            Width =1185
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label10"
                            Caption ="شماره "
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =2706
                    Height =300
                    ColumnWidth =4470
                    ColumnOrder =2
                    TabIndex =3
                    BackColor =16777215
                    ForeColor =0
                    Name ="ORDR_HED_CUST_NO"
                    ControlSource ="ORDR_HED_CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_NO FROM ORDR_HED GROUP BY CUST_NO ORDER BY CUST_NO"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2788
                            Top =7
                            Width =1155
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label11"
                            Caption ="سفارش دهنده"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1165
                    Top =2891
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =5
                    TabIndex =4
                    BackColor =16777215
                    ForeColor =0
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Format ="Fixed"
                    DefaultValue ="0"
                    FontName ="Tahoma"
                    Tag ="0"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4062
                            Top =2891
                            Width =675
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label28"
                            Caption ="مقداركالا"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =1186
                    Top =2551
                    Width =1680
                    Height =255
                    ColumnWidth =5865
                    ColumnOrder =4
                    TabIndex =5
                    BackColor =16777215
                    ForeColor =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME;"
                    ColumnWidths ="0;2268;567"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4227
                            Top =2551
                            Width =510
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label9"
                            Caption =" كالا"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1189
                    Top =3174
                    Height =255
                    ColumnWidth =840
                    TabIndex =6
                    BackColor =16777215
                    ForeColor =0
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    DefaultValue ="0"
                    FontName ="Tahoma"
                    Tag ="0"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3571
                            Top =3174
                            Width =1020
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label38"
                            Caption ="مقدار كل "
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1186
                    Top =4251
                    Width =1821
                    Height =255
                    ColumnWidth =1140
                    TabIndex =7
                    BackColor =16777215
                    ForeColor =0
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCO"
                        "D_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY T"
                        "COD_VAHEDS.NAMES"
                    ColumnWidths ="0"
                    FontName ="Tahoma"
                    Tag ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3693
                            Top =4251
                            Width =615
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label35"
                            Caption ="واحدكالا"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1133
                    Top =1247
                    TabIndex =8
                    BackColor =16777215
                    ForeColor =0
                    Name ="ORDR_LST_DATE"
                    ControlSource ="ORDR_LST_DATE"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1133
                    Top =1700
                    TabIndex =9
                    BackColor =16777215
                    ForeColor =0
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="=[RecordsetClone].[RecordCount]+1"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1246
                    Top =3571
                    Height =255
                    ColumnWidth =2835
                    TabIndex =10
                    BackColor =16777215
                    ForeColor =0
                    Name ="ORDR_LST_CUST_NO"
                    ControlSource ="ORDR_LST_CUST_NO"
                    DefaultValue ="Null"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3693
                            Top =3571
                            Width =1125
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label31"
                            Caption ="توضيحات"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5
                    Top =794
                    Width =7620
                    Height =315
                    TabIndex =11
                    BackColor =16777215
                    ForeColor =0
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    FontName ="Tahoma"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7726
                            Top =794
                            Width =795
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label33"
                            Caption ="كاربر"
                            FontName ="Tahoma"
                        End
                    End
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
' See "ORDR_HED1.cls"
