Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =300
    ItemSuffix =14
    Left =6900
    Top =3375
    Right =-30391
    Bottom =14775
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbf32e6cff1f7e340
    End
    RecordSource ="ANBARGRD_SUB2"
    Caption ="ANBGRD_LST"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =4265
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnOrder =0
                    Name ="GRD_NUM"
                    ControlSource ="GRD_NUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =525
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =1365
                    ColumnWidth =1185
                    ColumnOrder =5
                    TabIndex =2
                    Name ="MOG"
                    ControlSource ="MOG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1365
                            Width =1080
                            Height =240
                            Name ="Label5"
                            Caption ="موجودي فعلي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =1705
                    ColumnWidth =1020
                    ColumnOrder =6
                    TabIndex =3
                    Name ="NUM1"
                    ControlSource ="NUM1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1705
                            Width =915
                            Height =240
                            Name ="Label7"
                            Caption ="شمارش اول"
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
                    Left =1934
                    Top =2046
                    ColumnWidth =1050
                    ColumnOrder =7
                    TabIndex =4
                    Name ="NUM2"
                    ControlSource ="NUM2"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2046
                            Width =945
                            Height =240
                            Name ="Label9"
                            Caption ="شمارش دوم"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =2386
                    ColumnWidth =1140
                    ColumnOrder =8
                    TabIndex =5
                    Name ="NUM3"
                    ControlSource ="NUM3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2386
                            Width =1035
                            Height =240
                            Name ="Label11"
                            Caption ="شمارش سوم"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnWidth =7200
                    ColumnOrder =3
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAME FROM STUF_DEF"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =510
                            Height =240
                            Name ="Label3"
                            Caption ="كالا"
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
                    IMESentenceMode =3
                    Left =2954
                    Top =840
                    Width =1410
                    ColumnWidth =1170
                    ColumnOrder =1
                    TabIndex =6
                    Name ="Combo12"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =840
                            Width =510
                            Height =240
                            Name ="Label13"
                            Caption ="كد كالا"
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
                    Left =1984
                    Top =2664
                    ColumnWidth =660
                    ColumnOrder =9
                    TabIndex =7
                    Name ="EKHT2"
                    ControlSource ="EKH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =2664
                            Width =915
                            Height =240
                            Name ="Label15"
                            Caption ="اختلاف"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2040
                    Top =3004
                    ColumnWidth =1275
                    ColumnOrder =4
                    TabIndex =8
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =226
                            Top =3004
                            Width =915
                            Height =240
                            Name ="Label16"
                            Caption ="واحد"
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
                    Left =2154
                    Top =3401
                    ColumnWidth =1560
                    ColumnOrder =2
                    TabIndex =9
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =340
                            Top =3401
                            Width =915
                            Height =240
                            Name ="Label18"
                            Caption ="شماره فني"
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
                    Left =2210
                    Top =3855
                    TabIndex =10
                    Name ="Text19"
                    ControlSource ="grp"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =3855
                            Width =915
                            Height =240
                            Name ="Label20"
                            Caption ="گروه"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ANBGRD_LST_SUB2.cls"
