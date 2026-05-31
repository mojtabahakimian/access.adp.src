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
    Width =14985
    RowHeight =420
    ItemSuffix =25
    Left =1620
    Top =2160
    Right =21795
    Bottom =11250
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x827d26ded075e440
    End
    RecordSource ="VISITORS_DAY"
    Caption ="تعيين ويزيت روزانه"
    OnCurrent ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4025
                    Top =510
                    Width =1266
                    Height =317
                    ColumnWidth =1980
                    ColumnOrder =7
                    TabIndex =2
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =5387
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =113
                    Top =510
                    Width =1515
                    Height =317
                    ColumnOrder =5
                    TabIndex =3
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1688
                            Top =510
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =2034
                    Top =113
                    Width =9531
                    Height =317
                    ColumnWidth =5325
                    ColumnOrder =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.HES, CUST_HESAB.NAME, CUST_HESAB.MOBILE FROM Visit_route INNE"
                        "R JOIN CUST_HESAB ON Visit_route.HES = CUST_HESAB.hes GROUP BY Visit_route.HES, "
                        "CUST_HESAB.NAME, CUST_HESAB.MOBILE"
                    ColumnWidths ="0;3969"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11649
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =120
                    Width =1926
                    Height =317
                    ColumnOrder =8
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES2"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT HES FROM Visit_route GROUP BY HES"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =600
                            Top =120
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =61
                    Top =1027
                    Width =14850
                    Height =7830
                    TabIndex =4
                    Name ="VISITORS_DAY_DTL_SUB"
                    SourceObject ="Form.VISITORS_DAY_DTL_SUB"
                    LinkChildFields ="HES;VDATE"
                    LinkMasterFields ="HES;VDATE"
                    OnEnter ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =12991
                            Top =1177
                            Width =1635
                            Height =240
                            Name ="Visit_route_dtl_FORM Label"
                            Caption ="مشتريان زير مجموعه"
                            EventProcPrefix ="Visit_route_dtl_FORM_Label"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =134
                    TextFontCharSet =0
                    Left =2211
                    Top =510
                    Width =1020
                    Height =465
                    FontWeight =700
                    TabIndex =5
                    Name ="Command100"
                    Caption ="&چاپ "
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =1670
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12472
                    Top =113
                    Width =1266
                    Height =317
                    TabIndex =6
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    DefaultValue ="=FARSIDATE(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13832
                            Top =120
                            Width =870
                            Height =240
                            Name ="Label22"
                            Caption ="تاريخ ويزيت:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =10374
                    Top =510
                    Width =4185
                    Height =465
                    FontWeight =700
                    TabIndex =7
                    Name ="Command23"
                    Caption ="دريافت ليست مشتريان از مسيرهاي تعيين شده "
                    OnClick ="[Event Procedure]"
                    OnLostFocus ="[Event Procedure]"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =8670
                    Top =630
                    TabIndex =8
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9020
                            Top =600
                            Width =570
                            Height =240
                            Name ="Label17"
                            Caption ="فعال"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITORS_DAY_HEAD.cls"
