Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
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
    Width =4478
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =3705
    Top =2250
    Right =8190
    Bottom =3930
    HelpContextId =16
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x095a33e52d41e240
    End
    Caption ="انتخاب  سال"
    HelpFile ="SECRET.HLP"
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1701
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =360
                    Width =3329
                    Height =345
                    Name ="YEA"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3737
                            Top =367
                            Width =435
                            Height =330
                            Name ="Label1"
                            Caption ="سال:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =2095
                    Top =1077
                    Width =1095
                    Height =405
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="Command3"
                    Caption ="&تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =877
                    Top =1078
                    Width =1095
                    Height =405
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="Command4"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "F_SELECTYEAR.cls"
