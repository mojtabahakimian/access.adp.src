Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =10
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4860
    DatasheetFontWeight =700
    ItemSuffix =58
    Left =825
    Top =720
    Right =5640
    Bottom =6165
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x207e12471387e240
    End
    RecordSource ="list_salary"
    Caption ="ليست كامل حقوق و دستمزد پرسنل"
    DatasheetFontName ="Arabic Transparent"
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
        Begin Section
            Height =10388
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =113
                    ColumnWidth =750
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3860
                            Top =113
                            Width =645
                            Height =240
                            Name ="Label1"
                            Caption ="كد معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =453
                    Width =2490
                    ColumnWidth =585
                    TabIndex =1
                    Name ="PNAME"
                    ControlSource ="PNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4235
                            Top =453
                            Width =270
                            Height =240
                            Name ="Label3"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =793
                    Width =2490
                    ColumnWidth =1095
                    TabIndex =2
                    Name ="PFAMILY"
                    ControlSource ="PFAMILY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3560
                            Top =793
                            Width =945
                            Height =240
                            Name ="Label5"
                            Caption ="نام خانوادگي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1133
                    Width =2490
                    ColumnWidth =1290
                    TabIndex =3
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3215
                            Top =1133
                            Width =1290
                            Height =240
                            Name ="Label7"
                            Caption ="شماره شناسنامه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =1474
                    Width =2490
                    ColumnWidth =645
                    TabIndex =4
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3995
                            Top =1474
                            Width =510
                            Height =240
                            Name ="Label9"
                            Caption ="نام پدر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =1814
                    ColumnWidth =960
                    TabIndex =5
                    Name ="BIMEH_NUM"
                    ControlSource ="BIMEH_NUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3635
                            Top =1814
                            Width =870
                            Height =240
                            Name ="Label11"
                            Caption ="شماره بيمه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =2154
                    Width =2490
                    ColumnWidth =780
                    TabIndex =6
                    Name ="JOB"
                    ControlSource ="JOB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4055
                            Top =2154
                            Width =450
                            Height =240
                            Name ="Label13"
                            Caption ="شغل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2374
                    Top =2494
                    ColumnWidth =615
                    TabIndex =7
                    Name ="SEX"
                    ControlSource ="SEX"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3875
                            Top =2494
                            Width =630
                            Height =240
                            Name ="Label15"
                            Caption ="جنسيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =2834
                    ColumnWidth =1500
                    TabIndex =8
                    Name ="WSDATE"
                    ControlSource ="WSDATE"
                    Format ="yyyy/mm/dd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3245
                            Top =2834
                            Width =1260
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ شروع به كار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =3174
                    ColumnWidth =1170
                    TabIndex =9
                    Name ="WEDATE"
                    ControlSource ="WEDATE"
                    Format ="yyyy/mm/dd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3575
                            Top =3174
                            Width =930
                            Height =240
                            Name ="Label19"
                            Caption ="تاريخ ترك كار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =3514
                    ColumnWidth =1110
                    TabIndex =10
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    Format ="yyyy/mm/dd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4110
                            Top =3510
                            Width =630
                            Height =240
                            Name ="Label21"
                            Caption ="تاريخ ثبت"
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
                    Left =933
                    Top =3855
                    ColumnWidth =1290
                    TabIndex =11
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3470
                            Top =3855
                            Width =1035
                            Height =240
                            Name ="Label23"
                            Caption ="روزهاي كاركرد"
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
                    Left =933
                    Top =4195
                    ColumnWidth =1110
                    TabIndex =12
                    Name ="SALARY_DAYLY"
                    ControlSource ="SALARY_DAYLY"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3585
                            Top =4200
                            Width =1275
                            Height =405
                            Name ="Label25"
                            Caption ="دستمزد روزانه اسمي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =4535
                    ColumnWidth =1200
                    TabIndex =13
                    Name ="mosalary"
                    ControlSource ="mosalary"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3510
                            Top =4530
                            Width =1260
                            Height =405
                            Name ="Label27"
                            Caption ="دستمزد ماهانه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =4875
                    ColumnWidth =1800
                    TabIndex =14
                    Name ="mazmash"
                    ControlSource ="mazmash"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3765
                            Top =4875
                            Width =1020
                            Height =405
                            Name ="Label29"
                            Caption ="مزاياي ماهانه مشمول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =5215
                    ColumnWidth =2325
                    TabIndex =15
                    Name ="jmazsalmash"
                    ControlSource ="jmazsalmash"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3525
                            Top =5220
                            Width =1275
                            Height =405
                            Name ="Label31"
                            Caption ="جمع دستمزد و مزاياي مشمول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =5555
                    ColumnWidth =3885
                    TabIndex =16
                    Name ="jamazsalmash"
                    ControlSource ="jamazsalmash"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3435
                            Top =5550
                            Width =1380
                            Height =795
                            Name ="Label33"
                            Caption ="جمع دستمزد و مزاياي ماهانه مشمول و غير مشمول"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =5896
                    ColumnWidth =1800
                    TabIndex =17
                    Name ="bimper"
                    ControlSource ="bimper"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3945
                            Top =5895
                            Width =885
                            Height =600
                            Name ="Label35"
                            Caption ="حق بيمه سهم بيمه شده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =6236
                    ColumnWidth =1500
                    TabIndex =18
                    Name ="ghabel"
                    ControlSource ="ghabel"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3945
                            Top =6240
                            Width =900
                            Height =405
                            Name ="Label37"
                            Caption ="مانده قابل پرداخت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =933
                    Top =6576
                    ColumnWidth =1620
                    TabIndex =19
                    Name ="HOME"
                    ControlSource ="HOME"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3050
                            Top =6576
                            Width =1455
                            Height =240
                            Name ="Label39"
                            Caption ="حق خواربار و مسكن"
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
                    Left =933
                    Top =6916
                    ColumnWidth =780
                    TabIndex =20
                    Name ="CHILDREN"
                    ControlSource ="CHILDREN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3830
                            Top =6916
                            Width =675
                            Height =240
                            Name ="Label41"
                            Caption ="حق اولاد"
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
                    Left =933
                    Top =7256
                    ColumnWidth =1335
                    TabIndex =21
                    Name ="CONDITIONS"
                    ControlSource ="CONDITIONS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3335
                            Top =7256
                            Width =1170
                            Height =240
                            Name ="Label43"
                            Caption ="شرايط محيط كار"
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
                    Left =933
                    Top =7596
                    ColumnWidth =720
                    TabIndex =22
                    Name ="JAZB"
                    ControlSource ="JAZB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3800
                            Top =7596
                            Width =705
                            Height =240
                            Name ="Label45"
                            Caption ="حق جذب"
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
                    Left =933
                    Top =7937
                    ColumnWidth =465
                    TabIndex =23
                    Name ="SAYER"
                    ControlSource ="SAYER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4100
                            Top =7937
                            Width =405
                            Height =240
                            Name ="Label47"
                            Caption ="ساير"
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
                    Left =933
                    Top =8277
                    ColumnWidth =795
                    TabIndex =24
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3815
                            Top =8277
                            Width =690
                            Height =240
                            Name ="Label49"
                            Caption ="اضافه كار"
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
                    Left =933
                    Top =8617
                    ColumnWidth =675
                    TabIndex =25
                    Name ="PADASH"
                    ControlSource ="PADASH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4025
                            Top =8617
                            Width =480
                            Height =240
                            Name ="Label51"
                            Caption ="پاداش"
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
                    Left =933
                    Top =8957
                    ColumnWidth =780
                    TabIndex =26
                    Name ="KASR_VAM"
                    ControlSource ="KASR_VAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3830
                            Top =8957
                            Width =675
                            Height =240
                            Name ="Label53"
                            Caption ="كسر وام"
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
                    Left =933
                    Top =9317
                    TabIndex =27
                    Name ="Text54"
                    ControlSource ="SALARY_DAYLY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3585
                            Top =9322
                            Width =1275
                            Height =405
                            Name ="Label55"
                            Caption ="دستمزد روزانه رسمي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =9807
                    TabIndex =28
                    Name ="DSW_BDATE"
                    ControlSource ="DSW_BDATE"
                    Format ="yy/mm/dd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2878
                            Top =9807
                            Width =750
                            Height =240
                            Name ="Label56"
                            Caption ="تاريخ تولد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =10148
                    TabIndex =29
                    Name ="DSW_NAT"
                    ControlSource ="DSW_NAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3133
                            Top =10148
                            Width =495
                            Height =240
                            Name ="Label57"
                            Caption ="مليت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "list_salary.cls"
