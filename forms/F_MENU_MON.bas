Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
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
    Width =5725
    DatasheetFontHeight =10
    ItemSuffix =34
    Left =3510
    Top =2910
    Right =9240
    Bottom =4770
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x52ea7ec6afa6e240
    End
    Caption ="گزارش"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =1870
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2948
                    Top =1366
                    Width =870
                    Height =390
                    FontWeight =700
                    TabIndex =2
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2080
                    Top =1366
                    Width =870
                    Height =392
                    FontWeight =700
                    TabIndex =3
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =175
                    Top =741
                    Width =234
                    Height =255
                    Name ="DTT"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =175
                    Top =1116
                    Width =234
                    Height =255
                    TabIndex =4
                    Name ="MMOIN"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3401
                    Top =170
                    Width =1431
                    Height =300
                    TabIndex =1
                    Name ="MON"
                    RowSourceType ="Value List"
                    RowSource ="1;فروردين;2;ارديبهشت;3;خرداد;4;تير;5;مرداد;6;شهريور;7;مهر;8;آبان;9;آذر;10;دي;11;"
                        "بهمن;12;اسفند"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4926
                            Top =172
                            Width =645
                            Height =285
                            Name ="Label21"
                            Caption ="ماه:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3401
                    Top =590
                    Width =1431
                    Height =300
                    TabIndex =5
                    Name ="VAH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4920
                            Top =585
                            Width =720
                            Height =285
                            Name ="Label31"
                            Caption ="واحد كالا:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =223
                    Top =175
                    Width =2001
                    Height =300
                    TabIndex =6
                    Name ="GRP"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_STUFGROUP"
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2312
                            Top =170
                            Width =720
                            Height =285
                            Name ="Label33"
                            Caption ="گروه كالا:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_MON.cls"
