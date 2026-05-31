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
    ItemSuffix =18
    Left =6900
    Top =3375
    Right =-30391
    Bottom =15705
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xaf70f4957bdfe340
    End
    RecordSource ="ANBARGRD_SUB3"
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
            Height =4605
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
                    ColumnOrder =4
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
                    ColumnWidth =1710
                    ColumnOrder =5
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
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =2046
                    ColumnWidth =1050
                    ColumnOrder =6
                    TabIndex =4
                    Name ="NUM2"
                    ControlSource ="NUM2"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1934
                    Top =2386
                    ColumnWidth =1140
                    ColumnOrder =7
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
                    ColumnWidth =6945
                    ColumnOrder =2
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
                    ColumnWidth =1515
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
                    Left =1934
                    Top =2715
                    ColumnWidth =1110
                    ColumnOrder =8
                    TabIndex =7
                    Name ="EKHT3"
                    ControlSource ="EKH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =2715
                            Width =915
                            Height =240
                            Name ="Label15"
                            Caption ="اختلاف"
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
                    Left =1303
                    Top =3174
                    ColumnWidth =1417
                    ColumnOrder =9
                    TabIndex =8
                    Name ="MABL"
                    ControlSource ="MABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4180
                            Top =3174
                            Width =525
                            Height =240
                            Name ="Label16"
                            Caption ="نرخ"
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
                    Left =1984
                    Top =3458
                    ColumnWidth =1260
                    ColumnOrder =3
                    TabIndex =9
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =3458
                            Width =915
                            Height =240
                            Name ="Label17"
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
                    Left =2551
                    Top =4081
                    TabIndex =10
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =737
                            Top =4081
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
                    Left =2494
                    Top =4365
                    TabIndex =11
                    Name ="Text19"
                    ControlSource ="grp"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =680
                            Top =4365
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
' See "ANBGRD_LST_SUB3.cls"
