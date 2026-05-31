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
    Width =9808
    RowHeight =270
    ItemSuffix =58
    Left =405
    Top =3690
    Right =11790
    Bottom =5580
    DatasheetForeColor =8388608
    DatasheetBackColor =13434828
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5fc838f3aa9ee240
    End
    RecordSource ="MEHMAN"
    Caption ="MEHMAN"
    BeforeInsert ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
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
        Begin ToggleButton
            TextFontCharSet =178
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin Section
            Height =4818
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7061
                    Top =57
                    Width =1236
                    Height =255
                    ColumnOrder =0
                    FontWeight =700
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =64
                            Width =585
                            Height =240
                            FontWeight =700
                            Name ="Label1"
                            Caption ="شماره"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5807
                    Top =397
                    Width =2490
                    Height =255
                    ColumnWidth =1005
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =1
                    Name ="MNAME"
                    ControlSource ="MNAME"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =397
                            Width =300
                            Height =240
                            FontWeight =700
                            Name ="Label3"
                            Caption ="نام"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5807
                    Top =737
                    Width =2490
                    Height =255
                    ColumnWidth =1650
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =2
                    Name ="MFAMILY"
                    ControlSource ="MFAMILY"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =737
                            Width =1095
                            Height =240
                            FontWeight =700
                            Name ="Label5"
                            Caption ="نام خانوادگي"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5807
                    Top =1077
                    Width =2490
                    Height =255
                    ColumnWidth =1050
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =3
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =1077
                            Width =570
                            Height =240
                            FontWeight =700
                            Name ="Label7"
                            Caption ="نام پدر"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5807
                    Top =1418
                    Width =2490
                    Height =255
                    ColumnWidth =1290
                    ColumnOrder =5
                    FontWeight =700
                    TabIndex =4
                    Name ="KNOW"
                    ControlSource ="KNOW"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =1418
                            Width =1440
                            Height =240
                            FontWeight =700
                            Name ="Label9"
                            Caption ="شماره شناسنامه"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5807
                    Top =1758
                    Width =2490
                    Height =255
                    ColumnWidth =960
                    ColumnOrder =6
                    FontWeight =700
                    TabIndex =5
                    Name ="MSODUR"
                    ControlSource ="MSODUR"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =1758
                            Width =885
                            Height =240
                            FontWeight =700
                            Name ="Label11"
                            Caption ="محل صدور"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6752
                    Top =2098
                    Width =1545
                    Height =255
                    ColumnWidth =1065
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =6
                    Name ="BDATE"
                    ControlSource ="BDATE"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma (Arabic)"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =2098
                            Width =810
                            Height =240
                            FontWeight =700
                            Name ="Label13"
                            Caption ="تاريخ تولد"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3332
                    Top =4479
                    Width =4965
                    Height =255
                    ColumnWidth =1755
                    ColumnOrder =11
                    FontWeight =700
                    TabIndex =7
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =4479
                            Width =525
                            Height =240
                            FontWeight =700
                            Name ="Label27"
                            Caption ="آدرس"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1617
                    Top =114
                    Width =1710
                    Height =255
                    ColumnWidth =1365
                    ColumnOrder =12
                    FontWeight =700
                    TabIndex =8
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma (Arabic)"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3398
                            Top =114
                            Width =1545
                            Height =240
                            FontWeight =700
                            Name ="Label29"
                            Caption ="تاريخ  ورود"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1617
                    Top =454
                    Width =1710
                    Height =255
                    ColumnWidth =1350
                    ColumnOrder =14
                    FontWeight =700
                    TabIndex =9
                    Name ="KDATE"
                    ControlSource ="KDATE"
                    BeforeUpdate ="[Event Procedure]"
                    FontName ="Tahoma (Arabic)"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3398
                            Top =454
                            Width =1575
                            Height =240
                            FontWeight =700
                            Name ="Label31"
                            Caption ="تاريخ  خروج"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1626
                    Top =794
                    Height =255
                    ColumnOrder =15
                    FontWeight =700
                    TabIndex =10
                    Name ="HAMRAH"
                    ControlSource ="HAMRAH"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3405
                            Top =799
                            Width =1230
                            Height =240
                            FontWeight =700
                            Name ="Label33"
                            Caption ="تعداد همراهان"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1626
                    Top =1475
                    Height =255
                    ColumnOrder =16
                    FontWeight =700
                    TabIndex =11
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3398
                            Top =1475
                            Width =930
                            Height =240
                            FontWeight =700
                            Name ="Label37"
                            Caption ="تخت اضافه"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1626
                    Top =1815
                    Height =255
                    ColumnOrder =17
                    FontWeight =700
                    TabIndex =12
                    Name ="MABLPISH"
                    ControlSource ="MABLPISH"
                    Format ="#,###"
                    FontName ="Tahoma (Arabic)"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3398
                            Top =1815
                            Width =1065
                            Height =240
                            FontWeight =700
                            Name ="Label39"
                            Caption ="پيش پرداخت"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =837
                    Top =2155
                    Width =2490
                    Height =255
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =13
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3398
                            Top =2155
                            Width =765
                            Height =240
                            FontWeight =700
                            Name ="Label41"
                            Caption ="ملاحظات"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =441
                    Height =255
                    ColumnOrder =18
                    FontWeight =700
                    TabIndex =14
                    Name ="TAFZ"
                    ControlSource ="TAFZ"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =512
                            Width =690
                            Height =240
                            FontWeight =700
                            Name ="Label47"
                            Caption ="تفصيلي"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5831
                    Top =2438
                    Width =2466
                    ColumnOrder =20
                    FontWeight =700
                    TabIndex =15
                    Name ="TAHSIL"
                    ControlSource ="TAHSIL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCODE_MADRAK.CODE, TCODE_MADRAK.MADRAK_NAME FROM TCODE_MADRAK;"
                    ColumnWidths ="0"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =2438
                            Width =750
                            Height =240
                            FontWeight =700
                            Name ="Label15"
                            Caption ="تحصيلات"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5837
                    Top =2778
                    Width =2460
                    Height =255
                    ColumnOrder =8
                    FontWeight =700
                    TabIndex =16
                    Name ="JOB"
                    ControlSource ="JOB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.JOB FROM MEHMAN GROUP BY MEHMAN.JOB ORDER BY MEHMAN.JOB;"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =2778
                            Width =480
                            Height =240
                            FontWeight =700
                            Name ="Label17"
                            Caption ="شغل"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5837
                    Top =3118
                    Width =2460
                    ColumnOrder =21
                    FontWeight =700
                    TabIndex =17
                    Name ="GHASD"
                    ControlSource ="GHASD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.GHASD FROM MEHMAN GROUP BY MEHMAN.GHASD ORDER BY MEHMAN.GHASD;"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =3118
                            Width =1155
                            Height =240
                            FontWeight =700
                            Name ="Label19"
                            Caption ="قصد مسافرت"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5837
                    Top =3458
                    Width =2460
                    Height =255
                    ColumnOrder =22
                    FontWeight =700
                    TabIndex =18
                    Name ="SCITY"
                    ControlSource ="SCITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.SCITY FROM MEHMAN GROUP BY MEHMAN.SCITY ORDER BY MEHMAN.SCITY;"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =3458
                            Width =795
                            Height =240
                            FontWeight =700
                            Name ="Label21"
                            Caption ="شهر مبدا"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5837
                    Top =3799
                    Width =2460
                    Height =255
                    ColumnOrder =23
                    FontWeight =700
                    TabIndex =19
                    Name ="DCITY"
                    ControlSource ="DCITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.DCITY FROM MEHMAN GROUP BY MEHMAN.DCITY ORDER BY MEHMAN.DCITY;"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =3799
                            Width =930
                            Height =240
                            FontWeight =700
                            Name ="Label23"
                            Caption ="شهر مقصد"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5837
                    Top =4139
                    Width =2460
                    ColumnOrder =24
                    FontWeight =700
                    TabIndex =20
                    Name ="ENGIN"
                    ControlSource ="ENGIN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MEHMAN.ENGIN FROM MEHMAN GROUP BY MEHMAN.ENGIN ORDER BY MEHMAN.ENGIN;"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =8368
                            Top =4139
                            Width =975
                            Height =240
                            FontWeight =700
                            Name ="Label25"
                            Caption ="وسيله سفر"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1626
                    Top =1135
                    Height =255
                    ColumnWidth =975
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =21
                    Name ="ROOMN"
                    ControlSource ="ROOMN"
                    RowSourceType ="Table/View/StoredProc"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3398
                            Top =1135
                            Width =975
                            Height =240
                            FontWeight =700
                            Name ="Label35"
                            Caption ="شماره اتاق"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2145
                    Top =2505
                    Width =1176
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =9
                    FontWeight =700
                    TabIndex =22
                    Name ="Combo51"
                    ControlSource ="SEX"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SEX.SEX_COD, SEX.SEX FROM SEX;"
                    ColumnWidths ="0"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3392
                            Top =2505
                            Width =975
                            Height =240
                            FontWeight =700
                            Name ="Label52"
                            Caption ="جنسيت"
                            FontName ="Tahoma (Arabic)"
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
                    Left =906
                    Top =3344
                    ColumnOrder =25
                    TabIndex =23
                    Name ="SAZMANM"
                    ControlSource ="SAZMANM"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =3630
                    Width =1710
                    Height =255
                    ColumnWidth =990
                    ColumnOrder =13
                    FontWeight =700
                    TabIndex =24
                    Name ="VTIME"
                    ControlSource ="VTIME"
                    FontName ="Tahoma (Arabic)"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3266
                            Top =3630
                            Width =1545
                            Height =240
                            FontWeight =700
                            Name ="Label55"
                            Caption =" ساعت ورود"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =4050
                    Width =1710
                    Height =255
                    ColumnWidth =1140
                    ColumnOrder =19
                    FontWeight =700
                    TabIndex =25
                    Name ="KTIME"
                    ControlSource ="KTIME"
                    FontName ="Tahoma (Arabic)"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =42
                            Left =3266
                            Top =4050
                            Width =1545
                            Height =240
                            FontWeight =700
                            Name ="Label57"
                            Caption ="ساعت خروج"
                            FontName ="Tahoma (Arabic)"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MEHMAN_sub.cls"
