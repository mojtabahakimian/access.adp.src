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
    Width =8310
    DatasheetFontHeight =10
    ItemSuffix =31
    Left =5400
    Top =2775
    Right =13710
    Bottom =5430
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="صدور سند اختتاميه"
    OnOpen ="[Event Procedure]"
    HelpFile ="amin98.hlp>main"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
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
            Height =2664
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3645
                    Top =2211
                    Width =885
                    Height =318
                    FontWeight =700
                    TabIndex =4
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2777
                    Top =2211
                    Width =870
                    Height =317
                    FontWeight =700
                    TabIndex =5
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
                    Left =4932
                    Top =2211
                    Width =729
                    Height =255
                    TabIndex =1
                    Name ="DTT"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3510
                    Top =390
                    Width =1431
                    Height =300
                    Name ="S2"
                    Format ="#"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5035
                            Top =392
                            Width =1545
                            Height =285
                            Name ="Label4"
                            Caption ="شماره سند اختتاميه:"
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
                    Left =3510
                    Top =810
                    Width =1431
                    Height =300
                    TabIndex =2
                    Name ="DT"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5035
                            Top =812
                            Width =1545
                            Height =285
                            Name ="Label21"
                            Caption ="تاريخ سند اختتاميه:"
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
                    Left =120
                    Top =1230
                    Width =4821
                    Height =300
                    TabIndex =3
                    Name ="SH"
                    DefaultValue ="\"سند اختتاميه\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5035
                            Top =1232
                            Width =1545
                            Height =285
                            Name ="Label23"
                            Caption ="شرح سند اختتاميه:"
                        End
                    End
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =1973
                    Top =1644
                    Width =4037
                    Height =448
                    TabIndex =6
                    Name ="KIND"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =3296
                            Top =1785
                            OptionValue =2
                            Name ="Option29"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =2381
                                    Top =1757
                                    Width =915
                                    Height =240
                                    Name ="Label30"
                                    Caption ="كليه حسابها"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5286
                            Top =1785
                            OptionValue =1
                            Name ="Option27"
                            Begin
                                Begin Label
                                    OverlapFlags =119
                                    Left =3741
                                    Top =1757
                                    Width =1545
                                    Height =240
                                    Name ="Label28"
                                    Caption ="فقط حسابهاي دائمي"
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SANAD_EKHTETAMIYAH.cls"
