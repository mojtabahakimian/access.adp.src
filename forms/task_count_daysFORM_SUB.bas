Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12642
    RowHeight =315
    ItemSuffix =5
    Left =8895
    Top =1140
    Right =14160
    Bottom =4650
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7c9a1db5f2f3e540
    End
    RecordSource ="task_count_days"
    DatasheetFontName ="Tahoma"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
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
            Height =3344
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6689
                    Top =2097
                    ColumnOrder =2
                    Name ="CountOfUSERNAME"
                    ControlSource ="CountOfUSERNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4988
                            Top =2097
                            Width =1560
                            Height =240
                            Name ="Label0"
                            Caption ="تعداد كار  ورودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6689
                    Top =2437
                    ColumnOrder =0
                    TabIndex =1
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4988
                            Top =2437
                            Width =705
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6689
                    Top =2777
                    ColumnWidth =2040
                    ColumnOrder =1
                    TabIndex =2
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4988
                            Top =2777
                            Width =945
                            Height =240
                            Name ="Label2"
                            Caption ="كاربر"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =623
                    Name ="sc"
                    ControlSource ="=Sum([CountOfUSERNAME])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =623
                            Width =1560
                            Height =240
                            Name ="Label4"
                            Caption ="تعداد كار  ورودي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "task_count_daysFORM_SUB.cls"
