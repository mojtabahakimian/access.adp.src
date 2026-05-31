Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13889
    RowHeight =330
    ItemSuffix =77
    Left =4875
    Top =390
    Right =18510
    Bottom =10530
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa65ea3dddb81e240
    End
    RecordSource ="PERSONEL"
    Caption ="مشخصات شخصي"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            Height =690
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =10
                    Left =4932
                    Top =15
                    Width =2985
                    Height =675
                    FontSize =18
                    ForeColor =16777215
                    Name ="Label62"
                    Caption ="تعريف مشخصات پرسنل"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =10
                    Left =4947
                    Width =2985
                    Height =675
                    FontSize =18
                    Name ="Label63"
                    Caption ="تعريف مشخصات پرسنل"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            SpecialEffect =1
            Height =3911
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9466
                    Top =472
                    Width =2016
                    Height =315
                    ColumnOrder =3
                    TabIndex =1
                    Name ="BIMEH_NUM"
                    ControlSource ="BIMEH_NUM"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11583
                            Top =472
                            Width =885
                            Height =300
                            Name ="Label3"
                            Caption ="شماره بيمه"
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
                    Left =8677
                    Top =832
                    Width =2805
                    Height =315
                    ColumnOrder =4
                    TabIndex =2
                    Name ="PNAME"
                    ControlSource ="PNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11583
                            Top =832
                            Width =300
                            Height =300
                            Name ="Label5"
                            Caption ="نام"
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
                    Left =8677
                    Top =1192
                    Width =2805
                    Height =315
                    ColumnOrder =5
                    TabIndex =3
                    Name ="PFAMILY"
                    ControlSource ="PFAMILY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11583
                            Top =1192
                            Width =1020
                            Height =300
                            Name ="Label7"
                            Caption ="نام خانوادگي"
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
                    Left =8677
                    Top =1552
                    Width =2805
                    Height =315
                    ColumnWidth =1305
                    ColumnOrder =6
                    TabIndex =4
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11583
                            Top =1552
                            Width =1335
                            Height =300
                            Name ="Label9"
                            Caption ="شماره شناسنامه"
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
                    Left =8677
                    Top =1912
                    Width =2805
                    Height =315
                    ColumnWidth =840
                    ColumnOrder =7
                    TabIndex =5
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11583
                            Top =1912
                            Width =660
                            Height =300
                            Name ="Label11"
                            Caption ="نام پدر"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =5553
                    Top =120
                    Width =5931
                    Height =315
                    ColumnWidth =3660
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR)) AS PER, NUMBER FR"
                        "OM TDETA_HES WHERE (NUMBER = 1) AND (N_KOL = 124) ORDER BY NAME + N' ' + RTRIM(C"
                        "AST(TNUMBER AS NVARCHAR))"
                    ColumnWidths ="0"
                    ValidationRule ="Not Is Null"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11585
                            Top =120
                            Width =705
                            Height =300
                            Name ="Label1"
                            Caption ="كد معين"
                        End
                    End
                End
                Begin OptionGroup
                    ColumnHidden = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =10511
                    Top =2789
                    Width =999
                    Height =743
                    ColumnOrder =8
                    TabIndex =7
                    Name ="Frame28"
                    ControlSource ="SEX"
                    DefaultValue ="-1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =11215
                            Top =2859
                            Name ="Option31"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =10810
                                    Top =2831
                                    Width =435
                                    Height =300
                                    Name ="Label32"
                                    Caption ="مذكر"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =11212
                            Top =3232
                            OptionValue =0
                            Name ="Option33"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =10787
                                    Top =3232
                                    Width =465
                                    Height =300
                                    Name ="Label34"
                                    Caption ="مونث"
                                End
                            End
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =11576
                    Top =2945
                    Width =765
                    Height =300
                    Name ="Label35"
                    Caption ="جنسيت"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =7077
                    Top =1215
                    ColumnOrder =13
                    TabIndex =13
                    Name ="SEX"
                    ControlSource ="SEX"
                    DefaultValue ="-1"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =6477
                            Top =1185
                            Width =630
                            Height =240
                            Name ="Label36"
                            Caption ="جنسيت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2105
                    Top =1200
                    Width =2016
                    Height =315
                    ColumnWidth =1545
                    ColumnOrder =14
                    TabIndex =14
                    Name ="WSDATE"
                    ControlSource ="WSDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4220
                            Top =1200
                            Width =1425
                            Height =300
                            Name ="Label38"
                            Caption ="تاريخ شروع به كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2105
                    Top =1560
                    Width =2016
                    Height =315
                    ColumnWidth =1290
                    ColumnOrder =15
                    TabIndex =15
                    Name ="WEDATE"
                    ControlSource ="WEDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4220
                            Top =1560
                            Width =1185
                            Height =300
                            Name ="Label40"
                            Caption ="تاريخ ترك كار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9872
                    Top =2861
                    ColumnOrder =9
                    TabIndex =8
                    Name ="TAB56"
                    ControlSource ="TAB56"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextFontCharSet =0
                            TextAlign =3
                            Left =9110
                            Top =2775
                            Width =765
                            Height =300
                            Name ="Label43"
                            Caption ="تبصره اي"
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
                    Left =2115
                    Top =840
                    Width =2010
                    Height =315
                    ColumnOrder =12
                    TabIndex =12
                    Name ="DSW_BDATE"
                    ControlSource ="DSW_BDATE"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4226
                            Top =840
                            Width =1215
                            Height =300
                            Name ="Label53"
                            Caption ="تاريخ تولد"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1320
                    Top =120
                    Width =2805
                    Height =315
                    ColumnOrder =10
                    TabIndex =10
                    Name ="DSW_IDPLC"
                    ControlSource ="DSW_IDPLC"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PERSONEL.DSW_IDPLC FROM PERSONEL GROUP BY PERSONEL.DSW_IDPLC ORDER BY PER"
                        "SONEL.DSW_IDPLC;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4230
                            Top =120
                            Width =855
                            Height =300
                            Name ="Label49"
                            Caption ="محل صدور"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1320
                    Top =480
                    Width =2805
                    Height =315
                    ColumnOrder =11
                    TabIndex =11
                    Name ="DSW_NAT"
                    ControlSource ="DSW_NAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PERSONEL.DSW_NAT FROM PERSONEL GROUP BY PERSONEL.DSW_NAT ORDER BY PERSONE"
                        "L.DSW_NAT;"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4226
                            Top =480
                            Width =1020
                            Height =300
                            Name ="Label51"
                            Caption ="مليت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =75
                    Top =1905
                    Width =4056
                    Height =315
                    ColumnOrder =16
                    TabIndex =16
                    Name ="JOB"
                    ControlSource ="JOB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Job_Code, Job_Desc FROM tab_job ORDER BY Job_Desc"
                    ColumnWidths ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4215
                            Top =1905
                            Width =1185
                            Height =300
                            Name ="Label42"
                            Caption ="شغل"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1320
                    Top =2310
                    Width =2811
                    Height =315
                    ColumnOrder =1
                    TabIndex =17
                    Name ="PUNIT"
                    ControlSource ="PUNIT"
                    RowSourceType ="Value List"
                    RowSource ="1;\"توليد\";2;\"فروش\";3;\"خدمات\";4;\"اداري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4215
                            Top =2310
                            Width =1185
                            Height =300
                            Name ="Label65"
                            Caption ="محل  كاركرد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1335
                    Top =2745
                    Width =2796
                    Height =315
                    ColumnOrder =2
                    TabIndex =18
                    Name ="COT"
                    ControlSource ="COT"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4215
                            Top =2745
                            Width =1575
                            Height =405
                            Name ="Label67"
                            Caption ="شماره كارت ساعت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1335
                    Top =3105
                    Width =2796
                    Height =315
                    ColumnWidth =2340
                    TabIndex =19
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextFontCharSet =0
                            Left =4215
                            Top =3105
                            Width =1575
                            Height =405
                            Name ="Label69"
                            Caption ="شماره حساب بانكي"
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
                    Left =8675
                    Top =2310
                    Width =2805
                    Height =315
                    TabIndex =6
                    Name ="MELLICOD"
                    ControlSource ="MELLICOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11581
                            Top =2310
                            Width =660
                            Height =300
                            Name ="Label71"
                            Caption ="كد ملي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =9872
                    Top =3236
                    TabIndex =9
                    Name ="DIHARD"
                    ControlSource ="DIHARD"
                    Begin
                        Begin Label
                            OverlapFlags =119
                            TextFontCharSet =0
                            TextAlign =3
                            Left =9110
                            Top =3150
                            Width =765
                            Height =300
                            Name ="Label74"
                            Caption ="جانباز"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =12699
                    Top =56
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =20
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =849
                    Top =3291
                    ColumnWidth =2070
                    TabIndex =21
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3054
                            Top =3288
                            Width =1515
                            Height =240
                            Name ="Label75"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =793
                    Top =3571
                    TabIndex =22
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2998
                            Top =3573
                            Width =1365
                            Height =240
                            Name ="Label76"
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
' See "PERSON.cls"
