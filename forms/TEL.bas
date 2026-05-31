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
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14116
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =570
    Top =1095
    Right =14685
    Bottom =8220
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0ae26492107fe240
    End
    Caption ="دفتر تلفن"
    DatasheetFontName ="Arial (Arabic)"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7143
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =117
                    Top =1133
                    Width =13770
                    Height =5670
                    TabIndex =2
                    Name ="TELSUB"
                    SourceObject ="Form.TELSUB"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4680
                    Top =225
                    Width =8549
                    Height =315
                    FontWeight =700
                    Name ="Text2"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13415
                            Top =225
                            Width =360
                            Height =300
                            FontWeight =700
                            Name ="Label3"
                            Caption ="نام:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4680
                    Top =660
                    Width =8549
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="Text8"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13410
                            Top =660
                            Width =480
                            Height =300
                            FontWeight =700
                            Name ="Label9"
                            Caption ="تلفن:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TEL.cls"
