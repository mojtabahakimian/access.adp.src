Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14966
    RowHeight =420
    ItemSuffix =26
    Left =1395
    Top =780
    Right =16380
    Bottom =8265
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x5406362e3462e440
    End
    RecordSource ="visitgol_head"
    Caption ="تعيين اهداف براي ويزيتورها"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8957
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7879
                    Top =510
                    Width =1266
                    Height =317
                    ColumnWidth =1980
                    ColumnOrder =4
                    TabIndex =3
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =9241
                            Top =510
                            Width =750
                            Height =240
                            Name ="Label11"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =975
                    Top =120
                    Width =1905
                    Height =317
                    ColumnOrder =3
                    TabIndex =4
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2940
                            Top =120
                            Width =345
                            Height =330
                            Name ="Label13"
                            Caption ="كاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =6237
                    Top =113
                    Width =5556
                    Height =317
                    ColumnWidth =6255
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.HES, CUST_HESAB.NAME, CUST_HESAB.hes AS Expr1 FROM Visit_rout"
                        "e INNER JOIN CUST_HESAB ON Visit_route.HES = CUST_HESAB.hes GROUP BY Visit_route"
                        ".HES, CUST_HESAB.NAME, CUST_HESAB.hes ORDER BY CUST_HESAB.NAME"
                    ColumnWidths ="0;3969"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11877
                            Top =113
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label19"
                            Caption ="ويزيتور"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3411
                    Top =113
                    Width =2766
                    Height =317
                    ColumnWidth =2025
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES2"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT HES FROM Visit_route GROUP BY HES ORDER BY HES"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4821
                            Top =113
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =61
                    Top =982
                    Width =14850
                    Height =7875
                    TabIndex =5
                    Name ="Visit_gol_dtl_FORM"
                    SourceObject ="Form.Visit_gol_dtl_FORM"
                    LinkChildFields ="MAH;HES"
                    LinkMasterFields ="MAH;HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12991
                            Top =623
                            Width =1155
                            Height =240
                            Name ="Visit_route_dtl_FORM Label"
                            Caption ="اهداف مورد نظر"
                            EventProcPrefix ="Visit_route_dtl_FORM_Label"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =12699
                    Top =113
                    Height =330
                    ColumnOrder =0
                    Name ="MAH"
                    ControlSource ="MAH"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فروردين\";2;\"ارديبهشت\";3;\"خرداد\";4;\"تير\";5;\"مرداد\";6;\"شهريور\";7;\""
                        "مهر\";8;\"آبان\";9;\"آذر\";10;\"دي\";11;\"بهمن\";12;\"اسفند\""
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14452
                            Top =113
                            Width =345
                            Height =240
                            Name ="Label21"
                            Caption ="ماه"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2385
                    Top =480
                    Width =801
                    TabIndex =6
                    Name ="FDATE"
                    ControlSource ="FDATE"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =3312
                            Top =480
                            Width =600
                            Height =240
                            Name ="Label23"
                            Caption ="از تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4024
                    Top =510
                    TabIndex =7
                    Name ="TODATE"
                    ControlSource ="TODATE"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =6841
                            Top =510
                            Width =585
                            Height =240
                            Name ="Label24"
                            Caption ="تا تاريخ"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =10265
                    Top =510
                    Width =2490
                    Height =405
                    TabIndex =8
                    Name ="Command25"
                    Caption ="اضافه كردن كالاها بر اساس  گروه"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =134
                    TextFontCharSet =0
                    Left =396
                    Top =510
                    Width =1020
                    Height =405
                    FontWeight =700
                    TabIndex =9
                    Name ="Command100"
                    Caption ="&چاپ "
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1670
                End
            End
        End
    End
End
CodeBehindForm
' See "Visit_GOL_FORM.cls"
