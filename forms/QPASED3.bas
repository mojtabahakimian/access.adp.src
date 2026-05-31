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
    Width =4257
    RowHeight =315
    ItemSuffix =25
    Left =4320
    Top =1380
    Right =15120
    Bottom =8640
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb8875c495d2ae340
    End
    RecordSource ="QPASED3"
    Caption ="QPASED3"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3855
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =428
                    Top =113
                    ColumnOrder =1
                    Name ="PRGID"
                    ControlSource ="PRGID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3460
                            Top =113
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="PRGID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =428
                    Top =793
                    ColumnWidth =915
                    ColumnOrder =2
                    TabIndex =1
                    Name ="CountOfCODBB"
                    ControlSource ="=[CountOfCODB]*[Forms]![PRGHEAD]![NUM]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2890
                            Top =793
                            Width =1110
                            Height =240
                            Name ="Label5"
                            Caption ="مقدار"
                        End
                    End
                End
                Begin CheckBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1869
                    Top =1133
                    ColumnOrder =3
                    TabIndex =2
                    Name ="PASED"
                    ControlSource ="PASED"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3445
                            Top =1133
                            Width =555
                            Height =240
                            Name ="Label7"
                            Caption ="PASED"
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
                    Left =428
                    Top =1474
                    ColumnWidth =765
                    ColumnOrder =4
                    TabIndex =3
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Format ="General Number"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3715
                            Top =1474
                            Width =285
                            Height =240
                            Name ="Label9"
                            Caption ="رده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =428
                    Top =1814
                    ColumnWidth =1365
                    ColumnOrder =5
                    TabIndex =4
                    Name ="MANF1"
                    ControlSource ="=[MANF]*[Forms]![prghead]![num]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2755
                            Top =1814
                            Width =1245
                            Height =240
                            Name ="Label11"
                            Caption ="دستمزد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =428
                    Top =2154
                    ColumnWidth =1380
                    ColumnOrder =6
                    TabIndex =5
                    Name ="SAR1"
                    ControlSource ="=[SAR]*[Forms]![prghead]![num]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2725
                            Top =2154
                            Width =1275
                            Height =240
                            Name ="Label13"
                            Caption =" سربار"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =397
                    Top =453
                    Width =1680
                    ColumnWidth =3660
                    ColumnOrder =0
                    TabIndex =6
                    Name ="CODB"
                    ControlSource ="CODB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME FROM STUF_DEF;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2463
                            Top =453
                            Width =1485
                            Height =240
                            Name ="Label14"
                            Caption ="كالاي نيمه ساخته"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =435
                    Top =2490
                    TabIndex =7
                    Name ="CountOfCODB"
                    ControlSource ="CountOfCODB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3192
                            Top =2490
                            Width =645
                            Height =240
                            Name ="Label16"
                            Caption ="Text15:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =453
                    Top =3174
                    ColumnWidth =1380
                    TabIndex =8
                    Name ="saniyeh"
                    ControlSource ="=[sanieh]*[Forms]![prghead]![num]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2750
                            Top =3174
                            Width =1275
                            Height =240
                            Name ="Label22"
                            Caption ="زمان"
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
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Name ="SIMBIBE_MANF1"
                    ControlSource ="=Sum([MANF])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2327
                            Width =1245
                            Height =240
                            Name ="Label18"
                            Caption ="جذب هزينه تبديل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =340
                    TabIndex =1
                    Name ="SIMBIBE_SAR1"
                    ControlSource ="=Sum([sAR])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2297
                            Top =340
                            Width =1275
                            Height =240
                            Name ="Label20"
                            Caption ="جذب هزينه سربار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =700
                    TabIndex =2
                    Name ="sanieh1"
                    ControlSource ="=Sum([sanieh])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2297
                            Top =700
                            Width =1275
                            Height =240
                            Name ="Label24"
                            Caption ="جذب هزينه سربار"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "QPASED3.cls"
