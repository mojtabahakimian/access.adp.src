Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =1
    ViewsAllowed =2
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6810
    RowHeight =315
    ItemSuffix =28
    Left =2145
    Top =480
    Right =16650
    Bottom =6735
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x78f9fe8d4edee540
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME + N' ' + ISNULL(STUF_DEF.TOZIH, N' ') + N' '"
        " + ISNULL(STUF_DEF.N_FANI, N' ') AS name, STUF_DEF.MABL_F, STUF_STK.MOGODI_A + S"
        "TUF_STK.MOGODI AS MOG, STUF_STK.ANBAR, STUF_DEF.VAHED, STUF_DEF.B_SEF, STUF_DEF."
        "TOZIH, STUF_DEF.N_FANI FROM STUF_DEF INNER JOIN STUF_STK ON STUF_DEF.CODE = STUF"
        "_STK.CODE INNER JOIN VAHEDS ON STUF_DEF.CODE = VAHEDS.CODE GROUP BY STUF_DEF.COD"
        "E, STUF_DEF.NAME + N' ' + ISNULL(STUF_DEF.TOZIH, N' ') + N' ' + ISNULL(STUF_DEF."
        "N_FANI, N' '), STUF_DEF.MABL_F, STUF_STK.MOGODI_A + STUF_STK.MOGODI, STUF_STK.AN"
        "BAR, STUF_DEF.VAHED, STUF_DEF.B_SEF, STUF_DEF.TOZIH, STUF_DEF.N_FANI"
    Caption ="."
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    AllowFormView =0
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
            Height =3358
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3047
                    Top =165
                    ColumnWidth =840
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5894
                            Top =165
                            Width =555
                            Height =240
                            Name ="Label16"
                            Caption ="كد كالا"
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
                    Left =3047
                    Top =505
                    ColumnWidth =6900
                    ColumnOrder =1
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5894
                            Top =505
                            Width =555
                            Height =240
                            Name ="Label17"
                            Caption ="نام كالا"
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
                    Left =3047
                    Top =845
                    ColumnWidth =1440
                    ColumnOrder =2
                    TabIndex =2
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5265
                            Top =855
                            Width =1260
                            Height =405
                            Name ="Label18"
                            Caption ="في عمده فروشي"
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
                    Left =3047
                    Top =1185
                    ColumnWidth =1080
                    ColumnOrder =4
                    TabIndex =3
                    Name ="MOG"
                    ControlSource ="mog"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =3
                            Left =4920
                            Top =1180
                            Width =1320
                            Height =240
                            Name ="Label19"
                            Caption ="موجودي"
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
                    Left =3047
                    Top =1525
                    ColumnWidth =480
                    ColumnOrder =7
                    TabIndex =4
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5804
                            Top =1525
                            Width =645
                            Height =240
                            Name ="Label20"
                            Caption ="انبار"
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
                    Left =3030
                    Top =1920
                    ColumnWidth =1395
                    ColumnOrder =3
                    TabIndex =5
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5248
                            Top =1930
                            Width =1260
                            Height =405
                            Name ="Label22"
                            Caption ="في خرده فروشي"
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
                    Left =3004
                    Top =2551
                    ColumnOrder =6
                    TabIndex =6
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5851
                            Top =2551
                            Width =555
                            Height =240
                            Name ="Label24"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2721
                    Top =3118
                    ColumnOrder =8
                    TabIndex =7
                    Name ="VAHED"
                    ControlSource ="VAHED"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5493
                            Top =3118
                            Width =630
                            Height =240
                            Name ="Label25"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3004
                    Top =2911
                    ColumnWidth =1830
                    ColumnOrder =5
                    TabIndex =8
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5850
                            Top =2910
                            Width =870
                            Height =240
                            Name ="Label27"
                            Caption ="شماره فني"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SERCHKAL.cls"
