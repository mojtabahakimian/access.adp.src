Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4258
    RowHeight =292
    ItemSuffix =16
    Left =3615
    Top =1320
    Right =9600
    Bottom =6255
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x040d7256b9a1e240
    End
    RecordSource ="RAAS"
    Caption ="RAAS"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =2097
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =283
                    Top =453
                    ColumnWidth =1125
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    ValidationRule ="Is Not Null"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2805
                            Top =453
                            Width =1050
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =795
                    Height =255
                    ColumnWidth =1575
                    TabIndex =1
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3392
                            Top =795
                            Width =465
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =120
                    ColumnWidth =780
                    TabIndex =2
                    Name ="ROOS"
                    ControlSource ="ROOS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3012
                            Top =120
                            Width =675
                            Height =240
                            Name ="Label8"
                            Caption ="مدت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =291
                    Top =1187
                    ColumnWidth =2040
                    TabIndex =3
                    Name ="MABK"
                    ControlSource ="MABK"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3018
                            Top =1187
                            Width =675
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ روز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =300
                    Top =1530
                    TabIndex =4
                    Name ="ID"
                    ControlSource ="ID"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =396
                    Top =1814
                    TabIndex =5
                    Name ="USERID"
                    ControlSource ="USERID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3108
                            Top =1814
                            Width =690
                            Height =240
                            Name ="Label15"
                            Caption ="کاربر"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1320
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Name ="Text6"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3392
                            Width =465
                            Height =240
                            Name ="Label7"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =360
                    TabIndex =1
                    Name ="SMABK"
                    ControlSource ="=Sum([MABK])"
                    Format ="#,###"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =720
                    TabIndex =2
                    Name ="SROOS"
                    ControlSource ="=Sum([ROOS])"
                    Format ="#,###"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =285
                    Top =1080
                    TabIndex =3
                    Name ="CID"
                    ControlSource ="=Count([ID])"
                    Format ="#,###"
                End
            End
        End
    End
End
CodeBehindForm
' See "RAAS_SUB.cls"
