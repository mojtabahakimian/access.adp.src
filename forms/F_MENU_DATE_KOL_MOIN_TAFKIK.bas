Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
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
    Width =8900
    DatasheetFontHeight =10
    ItemSuffix =51
    Left =1500
    Top =2025
    Right =10395
    Bottom =4500
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="تفكيك معين"
    HelpFile ="amin98.hlp>main"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            Height =2494
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4374
                    Top =2097
                    Width =870
                    Height =330
                    FontWeight =700
                    TabIndex =8
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3506
                    Top =2097
                    Width =870
                    Height =332
                    FontWeight =700
                    TabIndex =9
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7022
                    Top =1515
                    Width =1056
                    Height =300
                    TabIndex =6
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8172
                            Top =1517
                            Width =645
                            Height =285
                            Name ="Label21"
                            Caption ="از تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5040
                    Top =1521
                    Width =1236
                    Height =300
                    TabIndex =7
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6370
                            Top =1523
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =106
                    Top =225
                    Width =3861
                    Height =300
                    TabIndex =2
                    Name ="HKOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME, TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TO"
                        "TA_HES.NAME;"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4037
                            Top =225
                            Width =1005
                            Height =285
                            Name ="Label17"
                            Caption ="حساب كل:"
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =106
                    Top =660
                    Width =3861
                    Height =300
                    TabIndex =3
                    Name ="HMOIN1"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;2835;567"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4037
                            Top =662
                            Width =840
                            Height =285
                            Name ="Label4"
                            Caption ="معين:"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =6173
                    Top =165
                    Width =2654
                    Height =1273
                    Name ="HHES"
                    DefaultValue ="3"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =8477
                            Top =408
                            OptionValue =1
                            Name ="Option34"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =6660
                                    Top =375
                                    Width =1800
                                    Height =285
                                    Name ="Label35"
                                    Caption ="حذف حسـابهاي بدون مانده"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =8477
                            Top =738
                            OptionValue =2
                            Name ="Option36"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =6662
                                    Top =710
                                    Width =1815
                                    Height =285
                                    Name ="Label37"
                                    Caption ="حذف حسابهاي داراي مانده"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =8477
                            Top =1068
                            OptionValue =3
                            Name ="Option38"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =6300
                                    Top =1035
                                    Width =2175
                                    Height =285
                                    Name ="Label39"
                                    Caption ="نمايــش همــه حســـــابهــــا"
                                End
                            End
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =5106
                    Top =165
                    Width =947
                    Height =1273
                    TabIndex =1
                    Name ="BEDBES"
                    DefaultValue ="3"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5775
                            Top =408
                            OptionValue =1
                            Name ="Option41"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =5248
                                    Top =375
                                    Width =510
                                    Height =285
                                    Name ="Label42"
                                    Caption ="بدهكار"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5775
                            Top =738
                            OptionValue =2
                            Name ="Option43"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =5175
                                    Top =710
                                    Width =600
                                    Height =285
                                    Name ="Label44"
                                    Caption ="بستانكار"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5775
                            Top =1068
                            OptionValue =3
                            Name ="Option45"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =5368
                                    Top =1035
                                    Width =405
                                    Height =285
                                    Name ="Label46"
                                    Caption ="هردو"
                                End
                            End
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =106
                    Top =1080
                    Width =3861
                    Height =300
                    TabIndex =4
                    Name ="TAFZ1"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;2835;567"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4037
                            Top =1080
                            Width =975
                            Height =285
                            Name ="Label48"
                            Caption ="ازتفصيلي:"
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =106
                    Top =1522
                    Width =3861
                    Height =300
                    TabIndex =5
                    Name ="TAFZ2"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;2835;567"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4037
                            Top =1530
                            Width =915
                            Height =285
                            Name ="Label50"
                            Caption ="تاتفصيلي:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_DATE_KOL_MOIN_TAFKIK.cls"
