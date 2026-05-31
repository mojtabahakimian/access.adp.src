Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7035
    RowHeight =345
    ItemSuffix =15
    Left =5775
    Top =885
    Right =12810
    Bottom =4110
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    OrderBy ="PGHARAR.GHNUM DESC"
    RecSrcDt = Begin
        0x7424972f0f8ce240
    End
    RecordSource ="PGHARAR"
    Caption ="قرارداد ها"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
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
            Height =615
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =2
                    Left =2834
                    Top =15
                    Width =1635
                    Height =600
                    FontSize =18
                    ForeColor =10092543
                    Name ="Label8"
                    Caption ="ثبت قرارداد"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =2
                    Left =2849
                    Width =1635
                    Height =600
                    FontSize =18
                    Name ="Label10"
                    Caption ="ثبت قرارداد"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =2628
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2636
                    Top =498
                    Width =1956
                    Height =315
                    ColumnOrder =1
                    TabIndex =1
                    Name ="GHNUM"
                    ControlSource ="GHNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =4678
                            Top =498
                            Width =1110
                            Height =300
                            Name ="Label3"
                            Caption ="شماره قرارداد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2640
                    Top =855
                    Width =1956
                    Height =315
                    ColumnWidth =1605
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =16777164
                    Name ="GHSTART"
                    ControlSource ="GHSTART"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =4682
                            Top =855
                            Width =1530
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ شروع قرارداد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2640
                    Top =1215
                    Width =1956
                    Height =315
                    ColumnWidth =1605
                    ColumnOrder =3
                    TabIndex =3
                    BackColor =16777164
                    Name ="GHEND"
                    ControlSource ="GHEND"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =4682
                            Top =1215
                            Width =1530
                            Height =300
                            Name ="Label7"
                            Caption ="تاريخ پايان قرار داد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2640
                    Top =2280
                    Width =1956
                    Height =315
                    ColumnWidth =2025
                    ColumnOrder =4
                    TabIndex =7
                    BackColor =13434828
                    Name ="MORAKH"
                    ControlSource ="MORAKH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =4680
                            Top =2280
                            Width =2130
                            Height =300
                            Name ="Label9"
                            Caption ="مانده مرخصي تا اين قرارداد"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =450
                    Top =90
                    Width =4821
                    Height =315
                    ColumnWidth =4260
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5372
                            Top =90
                            Width =675
                            Height =300
                            Name ="Label1"
                            Caption ="پرسنل"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3929
                    Top =1908
                    Width =666
                    Height =315
                    TabIndex =6
                    BackColor =16777164
                    Name ="MMINU"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="دقيقه"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4694
                            Top =1908
                            Width =1080
                            Height =300
                            Name ="Label13"
                            Caption ="مانده مرخصي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3266
                    Top =1908
                    Width =606
                    Height =315
                    ColumnWidth =630
                    TabIndex =5
                    BackColor =16777164
                    Name ="MHOU"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="ساعت"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2640
                    Top =1905
                    Width =576
                    Height =315
                    TabIndex =4
                    BackColor =16777164
                    Name ="MDA"
                    Format ="Fixed"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ControlTipText ="روز"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =2715
                    Top =1590
                    Width =405
                    Height =300
                    Name ="Label11"
                    Caption ="روز"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    Left =3300
                    Top =1590
                    Width =540
                    Height =300
                    Name ="Label12"
                    Caption ="ساعت"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3995
                    Top =1587
                    Width =540
                    Height =300
                    Name ="Label14"
                    Caption ="دقيقه"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =396
                    Top =1193
                    ColumnWidth =2070
                    TabIndex =8
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2601
                            Top =1190
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =340
                    Top =1473
                    TabIndex =9
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2545
                            Top =1475
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "PGHARAR.cls"
