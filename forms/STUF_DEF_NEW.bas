Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11463
    DatasheetFontHeight =10
    ItemSuffix =32
    Left =2745
    Top =1875
    Right =14205
    Bottom =7860
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xbe8f909639e3e340
    End
    Caption ="تعريف مشخصات كالا بر اساس كدينگ"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
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
            Height =6007
            BackColor =-2147483633
            Name ="Detail"
            Begin
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
                    Left =1020
                    Top =226
                    Width =5046
                    Height =397
                    Name ="Text0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6236
                            Top =283
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="نام كالا"
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
                    Left =7817
                    Top =170
                    Width =1881
                    Height =397
                    TabIndex =2
                    Name ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9810
                            Top =225
                            Width =1125
                            Height =240
                            Name ="Label3"
                            Caption ="كد فني(باركد)"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5782
                    Top =996
                    Width =3861
                    Height =330
                    TabIndex =1
                    Name ="T1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME, MPCODE AS Expr1 FROM TCOD_MAP WHERE (MPP = 1) ORDER BY MP"
                        "NAME"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =993
                            Width =1650
                            Height =300
                            Name ="l1"
                            Caption ="سايز"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5782
                    Top =1449
                    Width =3861
                    Height =330
                    TabIndex =3
                    Name ="T2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME, MPCODE AS Expr1 FROM TCOD_MAP WHERE (MPP = 2) ORDER BY MP"
                        "NAME"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =1446
                            Width =1650
                            Height =300
                            Name ="L2"
                            Caption ="درجه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5782
                    Top =1902
                    Width =3861
                    Height =330
                    TabIndex =4
                    Name ="T3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME, MPCODE AS Expr1 FROM TCOD_MAP WHERE (MPP = 3) ORDER BY MP"
                        "NAME"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =1899
                            Width =1650
                            Height =300
                            Name ="L3"
                            Caption ="طرح"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5782
                    Top =2355
                    Width =3861
                    Height =330
                    TabIndex =5
                    Name ="T4"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME, MPCODE AS Expr1 FROM TCOD_MAP WHERE (MPP = 4)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =2352
                            Width =1650
                            Height =300
                            Name ="L4"
                            Caption ="رده"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5782
                    Top =2808
                    Width =3861
                    Height =330
                    TabIndex =6
                    Name ="T5"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 5)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =2805
                            Width =1650
                            Height =300
                            Name ="L5"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =900
                    Top =2265
                    Width =1469
                    Height =397
                    TabIndex =7
                    Name ="Command14"
                    Caption ="از نو"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2379
                    Top =998
                    Width =1311
                    Height =337
                    TabIndex =8
                    Name ="digi"
                    Format ="#\"  رقم\""
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =3774
                            Top =1053
                            Width =1545
                            Height =255
                            Name ="Label16"
                            Caption ="طول بخش سريالي كد"
                        End
                    End
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =907
                    Top =963
                    Width =1424
                    Height =446
                    TabIndex =9
                    Name ="Command17"
                    Caption ="اضافه كن"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =3630
                    Top =5550
                    Width =1814
                    Height =397
                    TabIndex =10
                    Name ="Command18"
                    Caption ="انتقال به تعريف كالا"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =5450
                    Top =5550
                    Width =1814
                    Height =397
                    TabIndex =11
                    Name ="Command19"
                    Caption ="خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =793
                    Top =3004
                    Width =1881
                    Height =397
                    TabIndex =12
                    Name ="TText0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5782
                    Top =3261
                    Width =3861
                    Height =330
                    TabIndex =13
                    Name ="t6"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 6)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =3258
                            Width =1650
                            Height =300
                            Name ="L6"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5782
                    Top =3714
                    Width =3861
                    Height =330
                    TabIndex =14
                    Name ="t7"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 7)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =3711
                            Width =1650
                            Height =300
                            Name ="L7"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5782
                    Top =4167
                    Width =3861
                    Height =330
                    TabIndex =15
                    Name ="t8"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 8)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =4164
                            Width =1650
                            Height =300
                            Name ="l8"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5782
                    Top =4620
                    Width =3861
                    Height =330
                    TabIndex =16
                    Name ="t9"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP =9)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =4617
                            Width =1650
                            Height =300
                            Name ="l9"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    Enabled = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5782
                    Top =5073
                    Width =3861
                    Height =330
                    TabIndex =17
                    Name ="t10"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 8)"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =2
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =9813
                            Top =5070
                            Width =1650
                            Height =300
                            Name ="l10"
                            Caption ="كد كالا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "STUF_DEF_NEW.cls"
