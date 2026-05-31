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
    Width =5952
    DatasheetFontHeight =10
    ItemSuffix =38
    Left =5610
    Top =3195
    Right =11565
    Bottom =5385
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="گزارش حواله انبار گروهي"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =2211
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2852
                    Top =1644
                    Width =870
                    Height =315
                    FontWeight =700
                    TabIndex =6
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1984
                    Top =1644
                    Width =870
                    Height =317
                    FontWeight =700
                    TabIndex =7
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =687
                    Top =882
                    Width =234
                    Height =255
                    TabIndex =2
                    Name ="DTT"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3351
                    Top =768
                    Width =1416
                    Height =300
                    TabIndex =3
                    Name ="F1"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4867
                            Top =765
                            Width =675
                            Height =285
                            Name ="Label23"
                            Caption ="از فاكتور:"
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
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3334
                    Top =225
                    Width =1431
                    Height =300
                    Name ="CANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE FROM TCOD_ANBAR GROUP BY TCOD_ANBAR.CODE; "
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4859
                            Top =227
                            Width =495
                            Height =285
                            Name ="Label4"
                            Caption ="انبار:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =390
                    Top =239
                    Width =2841
                    Height =300
                    TabIndex =1
                    Name ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES FROM TCOD_ANBAR GROUP BY TCOD_ANBAR.COD"
                        "E, TCOD_ANBAR.NAMES ORDER BY TCOD_ANBAR.NAMES; "
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =390
                    Top =882
                    Width =234
                    Height =255
                    TabIndex =8
                    Name ="MANBAR"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =383
                    Top =1452
                    Width =234
                    Height =255
                    TabIndex =9
                    Name ="MKALA"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =1027
                    Top =768
                    Width =1416
                    Height =300
                    TabIndex =4
                    Name ="F2"
                    OnExit ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2542
                            Top =765
                            Width =660
                            Height =285
                            Name ="Label35"
                            Caption ="تا فاكتور:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =396
                    Top =1191
                    Width =4371
                    Height =300
                    TabIndex =5
                    Name ="VF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4860
                            Top =1185
                            Width =960
                            Height =285
                            Name ="Label37"
                            Caption ="و فاكتورهاي:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_ANBAR_FRKH_GRP.cls"
