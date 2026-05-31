Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
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
    Width =7143
    RowHeight =315
    ItemSuffix =22
    Left =2550
    Top =1125
    Right =9585
    Bottom =5910
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x89dbddaad22ce340
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.MABL_F, STUF_STK.MOGODI_A + STUF_S"
        "TK.MOGODI AS MOG, STUF_STK.ANBAR, STUF_DEF.VAHED, STUF_DEF.N_FANI FROM STUF_DEF "
        "INNER JOIN STUF_STK ON STUF_DEF.CODE = STUF_STK.CODE INNER JOIN VAHEDS ON STUF_D"
        "EF.CODE = VAHEDS.CODE GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.MABL_F, ST"
        "UF_STK.MOGODI_A + STUF_STK.MOGODI, STUF_STK.ANBAR, STUF_DEF.VAHED, STUF_DEF.N_FA"
        "NI"
    Caption ="."
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    AllowFormView =0
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
            Height =2154
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3677
                    ColumnWidth =615
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6524
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3677
                    Top =340
                    ColumnWidth =3150
                    TabIndex =1
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6524
                            Top =340
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3677
                    Top =680
                    ColumnWidth =990
                    TabIndex =2
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5895
                            Top =685
                            Width =840
                            Height =240
                            Name ="Label18"
                            Caption ="مبلغ فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3677
                    Top =1020
                    ColumnWidth =765
                    TabIndex =3
                    Name ="MOG"
                    ControlSource ="MOG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5550
                            Top =1015
                            Width =645
                            Height =240
                            Name ="Label19"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3677
                    Top =1360
                    ColumnWidth =480
                    TabIndex =4
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6434
                            Top =1360
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2600
                    Top =1870
                    TabIndex =5
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5310
                            Top =1875
                            Width =870
                            Height =240
                            Name ="Label21"
                            Caption ="شماره فني"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SERCHKALF12.cls"
