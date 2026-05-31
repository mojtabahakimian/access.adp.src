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
    AllowAdditions = NotDefault
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
    ItemSuffix =24
    Left =3510
    Top =2910
    Right =12405
    Bottom =6135
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="صدور سند افتتاحيه"
    OnOpen ="[Event Procedure]"
    HelpFile ="amin98.hlp>main"
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
            Height =3242
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4995
                    Top =2835
                    Width =870
                    Height =405
                    FontWeight =700
                    TabIndex =5
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4127
                    Top =2835
                    Width =870
                    Height =407
                    FontWeight =700
                    TabIndex =6
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
                    Left =6225
                    Top =2835
                    Width =729
                    Height =255
                    TabIndex =2
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
                    Left =4650
                    Top =1200
                    Width =1431
                    Height =300
                    Name ="S1"
                    Format ="#"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6166
                            Top =1200
                            Width =2460
                            Height =285
                            Name ="Label17"
                            Caption ="شماره سند اختتاميه سال قبل:"
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
                    Left =4650
                    Top =1597
                    Width =1431
                    Height =300
                    TabIndex =1
                    Name ="S2"
                    Format ="#"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6175
                            Top =1599
                            Width =1875
                            Height =285
                            Name ="Label4"
                            Caption ="شماره سند افتتاحيه:"
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
                    Left =4650
                    Top =2017
                    Width =1431
                    Height =300
                    TabIndex =3
                    Name ="DT"
                    Format ="0000\\/00\\/00"
                    ValidationRule ="Not Is Null"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6175
                            Top =2019
                            Width =1875
                            Height =285
                            Name ="Label21"
                            Caption ="تاريخ سند افتتاحيه:"
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
                    Left =1260
                    Top =2437
                    Width =4821
                    Height =300
                    TabIndex =4
                    Name ="SH"
                    DefaultValue ="\"سند افتتاحيه\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6175
                            Top =2439
                            Width =1875
                            Height =285
                            Name ="Label23"
                            Caption ="شرح سند افتتاحيه:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1530
                    Top =17
                    Width =4499
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =7
                    BackColor =13434828
                    Name ="odbs"
                    ValidationRule ="Not Is Null"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =6109
                            Width =2400
                            Height =510
                            FontSize =12
                            FontWeight =700
                            Name ="Label5"
                            Caption ="نام ديتا بيس در سال جاري:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1530
                    Top =596
                    Width =4514
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =8
                    Name ="ndbs"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Times New Roman"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =6120
                            Top =566
                            Width =2610
                            Height =525
                            FontSize =12
                            FontWeight =700
                            Name ="Label9"
                            Caption ="نام ديتا بيس  براي سال قبل:"
                            FontName ="Titr"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SANAD_EFTETAHIYAH.cls"
