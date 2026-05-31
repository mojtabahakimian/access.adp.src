Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14966
    RowHeight =420
    ItemSuffix =21
    Left =1395
    Top =990
    Right =16380
    Bottom =9930
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    AfterInsert ="[Event Procedure]"
    Filter ="((((((((Not USERNAME=\"modir-mali\"))) And ((Not USERNAME=\"آقاي شيخيان\"))) And"
        " ((Not USERNAME=\"فروش جنوب\"))) And ((Not USERNAME=\"آقاي دکتر حکيميان\"))) And"
        " ((Not USERNAME=\"MODIR-BAZARGANI\"))) And ((Not USERNAME=\"Controller\"))) And "
        "((Not USERNAME=\"Mr-kazemi\"))"
    OrderBy ="Visit_route.CDATE"
    RecSrcDt = Begin
        0xb4fbb71f1053e440
    End
    RecordSource ="Visit_route"
    Caption ="تعيين مسير وزيت براي مشتريان"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8445
                    Top =165
                    Width =4995
                    Height =317
                    ColumnWidth =2235
                    ColumnOrder =0
                    BackColor =13408767
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =13502
                            Top =165
                            Width =1200
                            Height =317
                            Name ="Label1"
                            Caption ="نام مسير  ويزيت"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =170
                    Top =680
                    Width =1266
                    Height =317
                    ColumnWidth =1980
                    ColumnOrder =7
                    TabIndex =6
                    Name ="CDATE"
                    ControlSource ="CDATE"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =1532
                            Top =680
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
                    Left =4305
                    Top =660
                    Width =1515
                    Height =317
                    ColumnOrder =5
                    TabIndex =7
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5880
                            Top =660
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
                    Left =2016
                    Top =165
                    Width =5556
                    Height =317
                    ColumnWidth =5325
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;3969"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7656
                            Top =165
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
                    Left =30
                    Top =165
                    Width =1926
                    Height =317
                    ColumnOrder =8
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES2"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    BeforeUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =600
                            Top =165
                            Width =735
                            Height =317
                            ForeColor =8388608
                            Name ="Label18"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12021
                    Top =675
                    Width =1980
                    Height =317
                    ColumnWidth =1575
                    ColumnOrder =2
                    TabIndex =3
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IYALAT FROM Visit_route GROUP BY IYALAT ORDER BY IYALAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14122
                            Top =675
                            Width =600
                            Height =317
                            Name ="Label5"
                            Caption ="استان"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9244
                    Top =675
                    Width =1980
                    Height =317
                    ColumnOrder =3
                    TabIndex =4
                    Name ="CITY"
                    ControlSource ="CITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITY FROM Visit_route GROUP BY CITY ORDER BY CITY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11445
                            Top =679
                            Width =525
                            Height =345
                            Name ="Label7"
                            Caption ="شهر"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2664
                    Top =710
                    ColumnOrder =6
                    TabIndex =8
                    Name ="RACTIVE"
                    ControlSource ="RACTIVE"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2894
                            Top =680
                            Width =690
                            Height =240
                            Name ="Label17"
                            Caption ="فعال:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6352
                    Top =675
                    Width =1980
                    Height =300
                    ColumnOrder =4
                    TabIndex =5
                    Name ="District"
                    ControlSource ="District"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT District FROM Visit_route GROUP BY District ORDER BY District"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8565
                            Top =679
                            Width =570
                            Height =345
                            Name ="Label9"
                            Caption ="منطقه"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =61
                    Top =1417
                    Width =14850
                    Height =7440
                    TabIndex =9
                    Name ="Visit_route_dtl_FORM"
                    SourceObject ="Form.Visit_route_dtl_FORM"
                    LinkChildFields ="ROUTE_NAME"
                    LinkMasterFields ="ROUTE_NAME"
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
                    OverlapFlags =247
                    AccessKey =134
                    TextFontCharSet =0
                    Left =340
                    Top =623
                    Width =1020
                    Height =465
                    FontWeight =700
                    TabIndex =10
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
' See "Visit_route_FORM.cls"
