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
    Width =4940
    RowHeight =315
    ItemSuffix =134
    Left =390
    Top =1470
    Right =14880
    Bottom =6600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x82530b95d2bde240
    End
    RecordSource ="KAR_BARG"
    Caption ="KAR_BARG"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =12038
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
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnOrder =32
                    Name ="KNUM"
                    ControlSource ="KNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="KNUM"
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
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnOrder =33
                    TabIndex =1
                    Name ="KMON"
                    ControlSource ="KMON"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =510
                            Height =240
                            Name ="Label3"
                            Caption ="KMON"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =480
                    ColumnOrder =31
                    TabIndex =3
                    Name ="D31"
                    ControlSource ="D31"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =360
                            Height =240
                            Name ="Label7"
                            Caption ="31"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =480
                    ColumnOrder =30
                    TabIndex =4
                    Name ="D30"
                    ControlSource ="D30"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =360
                            Height =240
                            Name ="Label9"
                            Caption ="30"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    ColumnWidth =480
                    ColumnOrder =29
                    TabIndex =5
                    Name ="D29"
                    ControlSource ="D29"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =360
                            Height =240
                            Name ="Label11"
                            Caption ="29"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1984
                    Top =2097
                    ColumnWidth =480
                    ColumnOrder =28
                    TabIndex =6
                    Name ="D28"
                    ControlSource ="D28"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =170
                            Top =2097
                            Width =255
                            Height =240
                            Name ="Label13"
                            Caption ="28"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2494
                    ColumnWidth =480
                    ColumnOrder =27
                    TabIndex =7
                    Name ="D27"
                    ControlSource ="D27"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2494
                            Width =360
                            Height =240
                            Name ="Label15"
                            Caption ="27"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =2834
                    ColumnWidth =480
                    ColumnOrder =26
                    TabIndex =8
                    Name ="D26"
                    ControlSource ="D26"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =2834
                            Width =360
                            Height =240
                            Name ="Label17"
                            Caption ="26"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =3174
                    ColumnWidth =480
                    ColumnOrder =25
                    TabIndex =9
                    Name ="D25"
                    ControlSource ="D25"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3174
                            Width =360
                            Height =240
                            Name ="Label19"
                            Caption ="25"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =3514
                    ColumnWidth =480
                    ColumnOrder =24
                    TabIndex =10
                    Name ="D24"
                    ControlSource ="D24"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3514
                            Width =360
                            Height =240
                            Name ="Label21"
                            Caption ="24"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =3855
                    ColumnWidth =480
                    ColumnOrder =23
                    TabIndex =11
                    Name ="D23"
                    ControlSource ="D23"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =3855
                            Width =360
                            Height =240
                            Name ="Label23"
                            Caption ="23"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =4195
                    ColumnWidth =480
                    ColumnOrder =22
                    TabIndex =12
                    Name ="D22"
                    ControlSource ="D22"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4195
                            Width =360
                            Height =240
                            Name ="Label25"
                            Caption ="22"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =4535
                    ColumnWidth =480
                    ColumnOrder =21
                    TabIndex =13
                    Name ="D21"
                    ControlSource ="D21"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4535
                            Width =360
                            Height =240
                            Name ="Label27"
                            Caption ="21"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =4875
                    ColumnWidth =480
                    ColumnOrder =20
                    TabIndex =14
                    Name ="D20"
                    ControlSource ="D20"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4875
                            Width =360
                            Height =240
                            Name ="Label29"
                            Caption ="20"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =5215
                    ColumnWidth =480
                    ColumnOrder =19
                    TabIndex =15
                    Name ="D19"
                    ControlSource ="D19"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5215
                            Width =360
                            Height =240
                            Name ="Label31"
                            Caption ="19"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =5555
                    ColumnWidth =480
                    ColumnOrder =18
                    TabIndex =16
                    Name ="D18"
                    ControlSource ="D18"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5555
                            Width =360
                            Height =240
                            Name ="Label33"
                            Caption ="18"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =5896
                    ColumnWidth =480
                    ColumnOrder =17
                    TabIndex =17
                    Name ="D17"
                    ControlSource ="D17"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =5896
                            Width =360
                            Height =240
                            Name ="Label35"
                            Caption ="17"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =6236
                    ColumnWidth =480
                    ColumnOrder =16
                    TabIndex =18
                    Name ="D16"
                    ControlSource ="D16"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =6236
                            Width =360
                            Height =240
                            Name ="Label37"
                            Caption ="16"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =6576
                    ColumnWidth =480
                    ColumnOrder =15
                    TabIndex =19
                    Name ="D15"
                    ControlSource ="D15"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =6576
                            Width =360
                            Height =240
                            Name ="Label39"
                            Caption ="15"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =6916
                    ColumnWidth =480
                    ColumnOrder =14
                    TabIndex =20
                    Name ="D14"
                    ControlSource ="D14"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =6916
                            Width =360
                            Height =240
                            Name ="Label41"
                            Caption ="14"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =7256
                    ColumnWidth =480
                    ColumnOrder =13
                    TabIndex =21
                    Name ="D13"
                    ControlSource ="D13"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =7256
                            Width =360
                            Height =240
                            Name ="Label43"
                            Caption ="13"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =7596
                    ColumnWidth =480
                    ColumnOrder =12
                    TabIndex =22
                    Name ="D12"
                    ControlSource ="D12"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =7596
                            Width =360
                            Height =240
                            Name ="Label45"
                            Caption ="12"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =7937
                    ColumnWidth =480
                    ColumnOrder =11
                    TabIndex =23
                    Name ="D11"
                    ControlSource ="D11"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =7937
                            Width =360
                            Height =240
                            Name ="Label47"
                            Caption ="11"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =8277
                    ColumnWidth =480
                    ColumnOrder =10
                    TabIndex =24
                    Name ="D10"
                    ControlSource ="D10"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =8277
                            Width =360
                            Height =240
                            Name ="Label49"
                            Caption ="10"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =8617
                    ColumnWidth =480
                    ColumnOrder =9
                    TabIndex =25
                    Name ="D9"
                    ControlSource ="D9"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =8617
                            Width =270
                            Height =240
                            Name ="Label51"
                            Caption ="9"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =8957
                    ColumnWidth =480
                    ColumnOrder =8
                    TabIndex =26
                    Name ="D8"
                    ControlSource ="D8"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =8957
                            Width =270
                            Height =240
                            Name ="Label53"
                            Caption ="8"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =9297
                    ColumnWidth =480
                    ColumnOrder =7
                    TabIndex =27
                    Name ="D7"
                    ControlSource ="D7"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =9297
                            Width =270
                            Height =240
                            Name ="Label55"
                            Caption ="7"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =9637
                    ColumnWidth =480
                    ColumnOrder =6
                    TabIndex =28
                    Name ="D6"
                    ControlSource ="D6"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =9637
                            Width =270
                            Height =240
                            Name ="Label57"
                            Caption ="6"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =9977
                    ColumnWidth =480
                    ColumnOrder =5
                    TabIndex =29
                    Name ="D5"
                    ControlSource ="D5"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =9977
                            Width =270
                            Height =240
                            Name ="Label59"
                            Caption ="5"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =10318
                    ColumnWidth =480
                    ColumnOrder =4
                    TabIndex =30
                    Name ="D4"
                    ControlSource ="D4"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =10318
                            Width =270
                            Height =240
                            Name ="Label61"
                            Caption ="4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =10658
                    ColumnWidth =480
                    ColumnOrder =3
                    TabIndex =31
                    Name ="D3"
                    ControlSource ="D3"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =10658
                            Width =270
                            Height =240
                            Name ="Label63"
                            Caption ="3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =10998
                    ColumnWidth =450
                    ColumnOrder =2
                    TabIndex =32
                    Name ="D2"
                    ControlSource ="D2"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =10998
                            Width =270
                            Height =240
                            Name ="Label65"
                            Caption ="2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =11338
                    ColumnWidth =450
                    ColumnOrder =1
                    TabIndex =33
                    Name ="D1"
                    ControlSource ="D1"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =11338
                            Width =270
                            Height =240
                            Name ="Label67"
                            Caption ="1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =11678
                    ColumnWidth =1755
                    TabIndex =34
                    Name ="ZD"
                    ControlSource ="ZD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =11678
                            Width =270
                            Height =240
                            Name ="Label69"
                            Caption ="جمع ساعت كاركرد"
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
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1927
                    Top =793
                    ColumnWidth =5460
                    ColumnOrder =0
                    TabIndex =2
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CAST(HEAD_MANF.FNUMB AS nvarchar) AS fnumb, ISNULL(HEAD_MANF.NAMES, STUF_"
                        "DEF.NAME) AS NAM, CAST(HEAD_MANF.FNUMB AS nvarchar) AS Expr1 FROM STUF_DEF RIGHT"
                        " OUTER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE UNION SELECT hes, NAME A"
                        "S nam, hes AS Expr1 FROM CUST_HESAb"
                    ColumnWidths ="0;3969"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =1080
                            Height =240
                            Name ="Label5"
                            Caption ="نام محل كاركرد"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =10785
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Name ="SD31"
                    ControlSource ="=Sum([D31])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =341
                    TabIndex =1
                    Name ="SD30"
                    ControlSource ="=Sum([D30])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =681
                    TabIndex =2
                    Name ="SD29"
                    ControlSource ="=Sum([D29])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1871
                    Top =964
                    TabIndex =3
                    Name ="SD28"
                    ControlSource ="=Sum([D28])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =1361
                    TabIndex =4
                    Name ="SD27"
                    ControlSource ="=Sum([D27])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =1701
                    TabIndex =5
                    Name ="SD26"
                    ControlSource ="=Sum([D26])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =2041
                    TabIndex =6
                    Name ="SD25"
                    ControlSource ="=Sum([D25])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =2381
                    TabIndex =7
                    Name ="SD24"
                    ControlSource ="=Sum([D24])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =2722
                    TabIndex =8
                    Name ="SD23"
                    ControlSource ="=Sum([D23])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =3062
                    TabIndex =9
                    Name ="SD22"
                    ControlSource ="=Sum([D22])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =3402
                    TabIndex =10
                    Name ="SD21"
                    ControlSource ="=Sum([D21])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =3742
                    TabIndex =11
                    Name ="SD20"
                    ControlSource ="=Sum([D20])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =4082
                    TabIndex =12
                    Name ="SD19"
                    ControlSource ="=Sum([D19])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =4422
                    TabIndex =13
                    Name ="SD18"
                    ControlSource ="=Sum([D18])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =4763
                    TabIndex =14
                    Name ="SD17"
                    ControlSource ="=Sum([D17])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =5103
                    TabIndex =15
                    Name ="SD16"
                    ControlSource ="=Sum([D16])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =5443
                    TabIndex =16
                    Name ="SD15"
                    ControlSource ="=Sum([D15])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =5783
                    TabIndex =17
                    Name ="SD14"
                    ControlSource ="=Sum([D14])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =6123
                    TabIndex =18
                    Name ="SD13"
                    ControlSource ="=Sum([D13])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =6463
                    TabIndex =19
                    Name ="SD12"
                    ControlSource ="=Sum([D12])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =6804
                    TabIndex =20
                    Name ="SD11"
                    ControlSource ="=Sum([D11])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =7144
                    TabIndex =21
                    Name ="SD10"
                    ControlSource ="=Sum([D10])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =7484
                    TabIndex =22
                    Name ="SD9"
                    ControlSource ="=Sum([D9])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =7824
                    TabIndex =23
                    Name ="SD8"
                    ControlSource ="=Sum([D8])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =8164
                    TabIndex =24
                    Name ="SD7"
                    ControlSource ="=Sum([D7])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =8504
                    TabIndex =25
                    Name ="SD6"
                    ControlSource ="=Sum([D6])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =8844
                    TabIndex =26
                    Name ="SD5"
                    ControlSource ="=Sum([D5])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =9185
                    TabIndex =27
                    Name ="SD4"
                    ControlSource ="=Sum([D4])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =9525
                    TabIndex =28
                    Name ="SD3"
                    ControlSource ="=Sum([D3])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =9865
                    TabIndex =29
                    Name ="SD2"
                    ControlSource ="=Sum([D2])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =10205
                    TabIndex =30
                    Name ="SD1"
                    ControlSource ="=Sum([D1])"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1814
                    Top =10545
                    TabIndex =31
                    Name ="sZD"
                    ControlSource ="=Sum([ZD])"
                End
            End
        End
    End
End
CodeBehindForm
' See "KAR_BARG_SUB.cls"
