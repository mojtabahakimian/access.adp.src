Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15250
    RowHeight =315
    ItemSuffix =23
    Top =15
    Right =14490
    Bottom =7560
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x8b23727b9cf4e240
    End
    RecordSource ="AZAE"
    Caption ="مشخصات اعضاء"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin FormHeader
            SpecialEffect =1
            Height =737
            BackColor =13421619
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =5273
                    Top =76
                    Width =5940
                    Height =585
                    FontSize =18
                    ForeColor =16777215
                    Name ="Label4"
                    Caption ="مشخصات مشتريان - اعضاء  و اعتبارات"
                    FontName ="Sina"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =5265
                    Top =60
                    Width =5940
                    Height =585
                    FontSize =18
                    ForeColor =13056
                    Name ="Label5"
                    Caption ="مشخصات مشتريان - اعضاء  و اعتبارات"
                    FontName ="Sina"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            SpecialEffect =1
            Height =9014
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11646
                    Top =630
                    Width =1935
                    Height =345
                    ColumnWidth =1425
                    ColumnOrder =2
                    TabIndex =2
                    Name ="TOPETEB"
                    ControlSource ="TOPETEB"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13620
                            Top =634
                            Width =1305
                            Height =240
                            Name ="Label3"
                            Caption ="سقف اعتبار ريالي"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =4629
                    Top =113
                    Width =9606
                    Height =315
                    ColumnWidth =4710
                    ColumnOrder =0
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME, hes AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0;6804;567"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =14319
                            Top =113
                            Width =735
                            Height =300
                            ForeColor =8388608
                            Name ="Label19"
                            Caption =" مشتري"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    AutoExpand = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1245
                    Top =120
                    Width =3111
                    Height =315
                    ColumnWidth =1417
                    ColumnOrder =1
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="حساب"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes FROM CUST_HESAB"
                    OnNotInList ="[Event Procedure]"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =11642
                    Top =1286
                    ColumnWidth =2640
                    TabIndex =4
                    Name ="TOPASN"
                    ControlSource ="TOPASN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12086
                            Top =1254
                            Width =2775
                            Height =240
                            Name ="Label6"
                            Caption ="اسناد پاس نشده از اعتبار كسر شود"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =11640
                    Top =1620
                    ColumnWidth =1785
                    TabIndex =5
                    Name ="ENDIS"
                    ControlSource ="ENDIS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =12072
                            Top =1594
                            Width =1770
                            Height =240
                            Name ="Label7"
                            Caption ="اعتبار سنجي كنترل شود"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8115
                    Top =630
                    Width =1935
                    Height =345
                    ColumnWidth =1575
                    TabIndex =3
                    Name ="TOPMEGH"
                    ControlSource ="TOPMEGH"
                    Format ="Standard"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10087
                            Top =637
                            Width =1470
                            Height =240
                            Name ="Label9"
                            Caption ="سقف اعتبار مقداري"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =2955
                    Top =623
                    Width =4740
                    Height =435
                    ForeColor =255
                    Name ="Label10"
                    Caption ="اگر سقف اعتبار مقداري را صفر بزنيد اعتبار مقداري كنترل نمي شود"
                End
                Begin Subform
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =180
                    Top =1995
                    Width =14853
                    Height =6967
                    TabIndex =6
                    Name ="GRADE_TAB_FORM"
                    SourceObject ="Form.GRADE_TAB_FORM"
                    LinkChildFields ="GCCUST_HES"
                    LinkMasterFields ="HES"
                    OnExit ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =226
                    Top =850
                    Width =1695
                    Height =405
                    TabIndex =7
                    Name ="ADDGRADE"
                    Caption ="اضافه کردن  رده بندي"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =226
                    Top =1360
                    Width =1695
                    Height =405
                    TabIndex =8
                    Name ="Command14"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8107
                    Top =1077
                    Width =1935
                    Height =345
                    TabIndex =9
                    BackColor =13408767
                    Name ="JAMZARIB"
                    ControlSource ="JAMZARIB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10081
                            Top =1081
                            Width =1305
                            Height =240
                            Name ="Label16"
                            Caption ="جمع ضرائب گروه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3004
                    Top =1360
                    Width =1935
                    Height =345
                    TabIndex =10
                    BackColor =13408767
                    Name ="EMTIAZ"
                    ControlSource ="EMTIAZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4978
                            Top =1364
                            Width =675
                            Height =240
                            Name ="Label18"
                            Caption ="امتياز کل"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5850
                    Top =1365
                    Width =1905
                    Height =330
                    TabIndex =11
                    BackColor =13408767
                    Name ="gradeid"
                    ControlSource ="gradeid"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GID, GNAME FROM GRADE_RANGE"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =7824
                            Top =1369
                            Width =675
                            Height =240
                            Name ="Label20"
                            Caption ="گريد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =171
                    Top =226
                    Width =225
                    Height =345
                    ColumnWidth =2085
                    TabIndex =12
                    BackColor =13408767
                    Name ="DTTIM"
                    ControlSource ="DTTIM"
                    DefaultValue ="Now()"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =8864
                    Top =1474
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =13
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
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
' See "AZAE.cls"
