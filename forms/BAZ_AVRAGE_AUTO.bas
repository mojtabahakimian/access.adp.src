Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5952
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =2970
    Top =45
    Right =8925
    Bottom =3495
    HelpContextId =131
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbece4ccb101ce240
    End
    Caption ="بازسازي مبالغ ميانگين انبارها"
    HelpFile ="AMIN98.HLP>MAIN"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3458
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =200
                    TextFontCharSet =0
                    Left =2895
                    Top =2790
                    Width =1644
                    Height =510
                    FontWeight =700
                    TabIndex =7
                    Name ="Command2"
                    Caption ="&بازسازي"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =539
                    Top =60
                    Width =4740
                    Height =855
                    FontWeight =700
                    Name ="Label9"
                    Caption ="لطفاً مطمئن شويد كه كاربر ديگري از سيستم استفاده نمي كند سپس روي گزينه بازسازي ك"
                        "ليك كنيد"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextFontFamily =10
                    BackStyle =0
                    Left =539
                    Top =967
                    Width =4743
                    Height =255
                    FontSize =6
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text0"
                    DefaultValue ="\"n\""
                    FontName ="Wingdings"
                End
                Begin Rectangle
                    Visible = NotDefault
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =539
                    Top =997
                    Width =4743
                    Height =245
                    Name ="Box7"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =1245
                    Top =2792
                    Width =1644
                    Height =510
                    FontWeight =700
                    TabIndex =8
                    Name ="Command14"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3514
                    Top =1700
                    Width =1116
                    Height =255
                    Name ="DT"
                    DefaultValue ="=[Forms]![baseknow]![yea]*10000"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =4681
                            Top =1700
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="از تاريخ:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =566
                    Top =1704
                    Width =1176
                    Height =255
                    TabIndex =4
                    Name ="Text19"
                    DefaultValue ="0"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =1866
                            Top =1700
                            Width =1395
                            Height =240
                            Name ="Label20"
                            Caption ="تعداد پردازش شده:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3514
                    Top =1303
                    Width =1116
                    Height =255
                    TabIndex =1
                    Name ="co"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =4681
                            Top =1303
                            Width =645
                            Height =240
                            Name ="Label22"
                            Caption ="كد كالا:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =566
                    Top =1307
                    Width =1176
                    Height =255
                    TabIndex =2
                    Name ="Text23"
                    DefaultValue ="0"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =1866
                            Top =1303
                            Width =750
                            Height =240
                            Name ="Label24"
                            Caption ="تعداد كالا:"
                            FontName ="Tahoma"
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
                    Left =540
                    Top =2040
                    Width =3906
                    Height =315
                    TabIndex =5
                    Name ="RADAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_STUFGROUP.CODE, TCOD_STUFGROUP.NAMES FROM TCOD_STUFGROUP WHERE (((TC"
                        "OD_STUFGROUP.CODE)<>0)) ORDER BY TCOD_STUFGROUP.NAMES;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4609
                            Top =2040
                            Width =705
                            Height =240
                            Name ="Label19"
                            Caption ="گروه كالا:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =3690
                    Top =2479
                    Width =794
                    Height =234
                    TabIndex =6
                    Name ="FORMOL"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =3
                            Left =975
                            Top =2385
                            Width =3240
                            Height =240
                            FontWeight =700
                            Name ="Label26"
                            Caption ="مبالغ استاندارد در فرمولها نيز اصلاح شود"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "BAZ_AVRAGE_AUTO.cls"
