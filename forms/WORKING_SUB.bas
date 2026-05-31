Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3973
    RowHeight =285
    ItemSuffix =12
    Left =525
    Top =5370
    Right =14445
    Bottom =8535
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xf7815bf0ba81e240
    End
    RecordSource ="WORKING"
    Caption ="WORKING"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
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
        Begin Section
            Height =2174
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =113
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2966
                            Top =113
                            Width =750
                            Height =240
                            Name ="Label1"
                            Caption ="كد پرسنل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =453
                    TabIndex =1
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3326
                            Top =453
                            Width =390
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =793
                    ColumnWidth =1590
                    TabIndex =2
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2681
                            Top =793
                            Width =1035
                            Height =240
                            Name ="Label5"
                            Caption ="روزهاي كاركرد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    ColumnWidth =1710
                    TabIndex =3
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3026
                            Top =1133
                            Width =690
                            Height =240
                            Name ="Label7"
                            Caption ="اضافه كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1474
                    ColumnWidth =1725
                    TabIndex =4
                    Name ="PADASH"
                    ControlSource ="PADASH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3240
                            Top =1470
                            Width =555
                            Height =240
                            Name ="Label9"
                            Caption ="راندمان"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1814
                    ColumnWidth =1590
                    TabIndex =5
                    Name ="KASR_VAM"
                    ControlSource ="KASR_VAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3041
                            Top =1814
                            Width =675
                            Height =240
                            Name ="Label11"
                            Caption ="كسر وام"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "WORKING_SUB.cls"
