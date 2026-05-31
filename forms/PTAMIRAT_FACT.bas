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
    Width =4440
    RowHeight =315
    ItemSuffix =50
    Left =570
    Top =1470
    Right =15030
    Bottom =9435
    HelpContextId =6001
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe0f994ed51d5e340
    End
    RecordSource ="SELECT INVO_LST.NUMBER, INVO_LST.TAG, INVO_LST.ANBAR, INVO_LST.CODE, INVO_LST.ME"
        "GH, INVO_LST.MEGHk, INVO_LST.MANDAH, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.VA"
        "HED_K, HEAD_LST.FNUMCO, HEAD_LST.MAS FROM INVO_LST INNER JOIN HEAD_LST ON INVO_L"
        "ST.NUMBER = HEAD_LST.NUMBER AND INVO_LST.TAG = HEAD_LST.TAG WHERE (HEAD_LST.MAS "
        "= 1010)"
    Caption ="INVO_LST"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =6859
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =612
                    Top =480
                    Height =255
                    ColumnWidth =1305
                    ColumnOrder =0
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3224
                            Top =480
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =612
                    Top =820
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3554
                            Top =820
                            Width =630
                            Height =240
                            Name ="Label3"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =599
                    Top =1814
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3496
                            Top =1814
                            Width =675
                            Height =240
                            Name ="Label11"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =2777
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =20
                    TabIndex =5
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3780
                            Top =2775
                            Width =660
                            Height =240
                            Name ="Label15"
                            Caption ="ملاحظات"
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
                    Left =625
                    Top =3117
                    Height =255
                    ColumnWidth =1050
                    ColumnOrder =8
                    TabIndex =6
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3822
                            Top =3117
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =625
                    Top =3457
                    Height =255
                    ColumnWidth =1245
                    ColumnOrder =9
                    TabIndex =7
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3582
                            Top =3457
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =620
                    Top =1474
                    Width =1680
                    Height =255
                    ColumnWidth =4695
                    ColumnOrder =3
                    TabIndex =3
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME FROM STUF_DEF"
                    ColumnWidths ="0;2268;567"
                    ValidationRule ="Is Not Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3661
                            Top =1474
                            Width =510
                            Height =240
                            Name ="Label9"
                            Caption =" كالا"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =623
                    Top =2097
                    Height =255
                    ColumnWidth =840
                    ColumnOrder =7
                    TabIndex =9
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3005
                            Top =2097
                            Width =1020
                            Height =240
                            Name ="Label38"
                            Caption ="مقدار كل "
                        End
                    End
                End
                Begin ComboBox
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =625
                    Top =6179
                    Width =696
                    Height =255
                    ColumnWidth =945
                    ColumnOrder =4
                    TabIndex =8
                    Name ="VAHED_K"
                    ControlSource ="VAHED_K"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3582
                            Top =6179
                            Width =615
                            Height =240
                            Name ="Label35"
                            Caption ="واحدكالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =627
                    Top =1160
                    Width =1686
                    Height =255
                    ColumnWidth =1725
                    ColumnOrder =2
                    TabIndex =2
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES FROM TCOD_ANBAR ORDER BY TCOD_ANBAR.COD"
                        "E;"
                    ColumnWidths ="0"
                    DefaultValue ="3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3854
                            Top =1160
                            Width =330
                            Height =240
                            Name ="Label5"
                            Caption ="انبار"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =311
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "PTAMIRAT_FACT.cls"
