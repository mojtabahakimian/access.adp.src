Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14116
    RowHeight =270
    ItemSuffix =20
    Left =1695
    Top =2445
    Right =13425
    Bottom =4770
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xcffedea96d9ce340
    End
    Caption ="استخراج آمار موجودي ابتداي دوره"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
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
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2339
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1125
                    Top =682
                    Width =924
                    Height =405
                    ColumnOrder =2
                    FontSize =12
                    FontWeight =700
                    TabIndex =7
                    BackColor =16764057
                    Name ="moen"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =1141
                            Top =165
                            Width =900
                            Height =510
                            Name ="Label3"
                            Caption ="موجودي انتهاي دوره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5925
                    Top =682
                    Width =924
                    Height =405
                    ColumnWidth =750
                    ColumnOrder =3
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="afkh"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =5941
                            Top =172
                            Width =900
                            Height =510
                            Name ="Label5"
                            Caption ="آف.خريد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4965
                    Top =682
                    Width =924
                    Height =405
                    ColumnWidth =855
                    ColumnOrder =4
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="affr"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =4981
                            Top =165
                            Width =900
                            Height =510
                            Name ="Label7"
                            Caption ="آف.فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4005
                    Top =682
                    Width =924
                    Height =405
                    ColumnWidth =1320
                    ColumnOrder =5
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="afbkh"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =4021
                            Top =165
                            Width =900
                            Height =510
                            Name ="Label10"
                            Caption ="آف.برگشت خريد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3045
                    Top =682
                    Width =924
                    Height =405
                    ColumnWidth =1425
                    ColumnOrder =6
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="afbfr"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =3061
                            Top =177
                            Width =900
                            Height =510
                            Name ="Label12"
                            Caption ="آف.برگشت فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2085
                    Top =682
                    Width =924
                    Height =405
                    ColumnWidth =915
                    ColumnOrder =7
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="aben"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =2101
                            Top =172
                            Width =900
                            Height =510
                            Name ="Label14"
                            Caption ="آب.انتقالي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =165
                    Top =682
                    Width =924
                    Height =405
                    ColumnWidth =1560
                    ColumnOrder =8
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    BackColor =13434828
                    Name ="moeb"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =181
                            Top =172
                            Width =900
                            Height =510
                            Name ="Label16"
                            Caption ="موجودي ابتداي دوره"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =12415
                    Top =675
                    Width =1584
                    Height =375
                    ColumnWidth =1560
                    ColumnOrder =0
                    FontSize =12
                    FontWeight =700
                    BackColor =16764057
                    Name ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR WHERE (CODE > 0)"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =12431
                            Top =165
                            Width =1560
                            Height =510
                            Name ="Label18"
                            Caption ="انبار"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =6915
                    Top =675
                    Width =5445
                    Height =390
                    ColumnWidth =2970
                    ColumnOrder =1
                    TabIndex =1
                    Name ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;4536;567"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            TextAlign =2
                            Left =6915
                            Top =165
                            Width =5445
                            Height =465
                            Name ="Label9"
                            Caption =" كالا"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =501
                    Top =1303
                    Width =1531
                    Height =397
                    TabIndex =9
                    Name ="Command19"
                    Caption ="اصلاح موجودي"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "anbargardani-baz.cls"
