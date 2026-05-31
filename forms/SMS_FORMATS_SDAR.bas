Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =285
    ItemSuffix =8
    Left =600
    Top =930
    Right =10215
    Bottom =9165
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xdf45ad0d5f1be440
    End
    RecordSource ="SELECT SMS_FORMATS.* FROM SMS_FORMATS WHERE (SMSTY = 5)"
    Caption ="SMS_FORMATS"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
        Begin Section
            Height =1493
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
                    Name ="SMSTY"
                    ControlSource ="SMSTY"
                    DefaultValue ="5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =555
                            Height =240
                            Name ="Label1"
                            Caption ="SMSTY"
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
                    Left =1927
                    Top =453
                    ColumnWidth =510
                    ColumnOrder =1
                    TabIndex =1
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =540
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
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
                    Left =1927
                    Top =1133
                    Width =2490
                    ColumnWidth =4335
                    ColumnOrder =2
                    TabIndex =3
                    Name ="SMS_CAPTION"
                    ControlSource ="SMS_CAPTION"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =1140
                            Height =240
                            Name ="Label7"
                            Caption ="عنوان  فيلد"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =4290
                    ColumnOrder =3
                    TabIndex =2
                    Name ="SMS_FIELS"
                    ControlSource ="SMS_FIELS"
                    RowSourceType ="Value List"
                    RowSource ="\"FHES1\";\"نام پرداخت كننده\";\"FHES\";\"كد پرداخت كننده \";\"MOLAH\";\"ملاحظات"
                        "\";\"NUMBER\";\"شماره خزانه\";\"DATE\";\"تاريخ \";\"SHARH\";\"شرح دريافت \";\"MA"
                        "NH\";\"مانده حساب\";\"NULL\";\"هيچي(خالي)\";\"ENTER\";\"رفتن به اول سطر بعد(اينت"
                        "ر)\";\"SPC\";\"فاصله\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =1035
                            Height =240
                            Name ="Label5"
                            Caption ="فيلد (انتخابي)"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SMS_FORMATS_SDAR.cls"
