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
    Width =5725
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =5730
    Top =3315
    Right =11460
    Bottom =5175
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
                    TabIndex =4
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
                    Left =175
                    Top =741
                    Width =234
                    Height =255
                    TabIndex =1
                    Name ="DTT"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3009
                    Top =706
                    Width =1431
                    Height =300
                    TabIndex =2
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4534
                            Top =708
                            Width =645
                            Height =285
                            Name ="Label21"
                            Caption ="از تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =781
                    Top =714
                    Width =1416
                    Height =300
                    TabIndex =3
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2291
                            Top =716
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
                        End
                    End
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
                    TabIndex =6
                    Name ="MMOIN"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =445
                    Top =240
                    Width =4101
                    Height =300
                    Name ="HHMOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4642
                            Top =240
                            Width =900
                            Height =285
                            Name ="Label29"
                            Caption ="نام شخص:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_CHEK.cls"
