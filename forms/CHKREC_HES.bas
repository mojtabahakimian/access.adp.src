Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14683
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =210
    Top =780
    Right =14895
    Bottom =7890
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xb92b93d2b289e440
    End
    RecordSource ="CHREC_HES"
    Caption ="ثبت به حساب گذاشتن چكهاي دريافتي"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =8917
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
                    Left =12546
                    Top =120
                    Width =1131
                    Height =300
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13793
                            Top =120
                            Width =750
                            Height =285
                            Name ="Label1"
                            Caption ="تاريخ:"
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
                    Left =9935
                    Top =460
                    Width =3735
                    Height =300
                    TabIndex =1
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13803
                            Top =457
                            Width =780
                            Height =285
                            Name ="Label3"
                            Caption ="ملاحظات:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =95
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3605
                    Top =585
                    Width =1641
                    Height =286
                    TabIndex =2
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            ReadingOrder =1
                            NumeralShapes =2
                            Left =5255
                            Top =585
                            Width =1425
                            Height =285
                            BackColor =32768
                            ForeColor =65535
                            Name ="lsanad"
                            Caption ="ثبت در سند شماره:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =90
                    Top =960
                    Width =14460
                    Height =7530
                    TabIndex =3
                    Name ="CHRE_LSTH_SUB"
                    SourceObject ="Form.CHRE_LSTH_SUB"
                    LinkChildFields ="DATE"
                    LinkMasterFields ="DATE"
                    OnExit ="[Event Procedure]"
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6918
                    Top =8617
                    Width =1983
                    Height =300
                    TabIndex =4
                    Name ="Text10"
                    ControlSource ="=[CHRE_LSTH_SUB].[Form]![smab]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8994
                            Top =8617
                            Width =780
                            Height =285
                            Name ="Label11"
                            Caption ="جمع مبالغ:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =141
                    TextFontCharSet =0
                    Left =113
                    Top =56
                    Width =1020
                    Height =450
                    FontWeight =700
                    TabIndex =5
                    Name ="PRINTC"
                    Caption ="&چاپ برگه"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1417
                    Top =510
                    Width =411
                    Height =345
                    TabIndex =6
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =255
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1474
                    Top =623
                    TabIndex =7
                    Name ="IDH"
                    ControlSource ="IDH"
                End
                Begin TextBox
                    OverlapFlags =247
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =566
                    TabIndex =8
                    Name ="CHK"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =1943
                            Top =510
                            Width =1515
                            Height =240
                            Name ="Label15"
                            Caption ="شماره فرمت  چاپ:"
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
                    Left =8460
                    Top =90
                    Width =2256
                    Height =315
                    TabIndex =9
                    Name ="CITY"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =10710
                    Top =105
                    Width =1755
                    Height =300
                    Name ="Label9"
                    Caption ="شهرستان محل پرداخت:"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1365
                    Top =120
                    Width =5181
                    Height =315
                    TabIndex =10
                    Name ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6758
                            Top =127
                            Width =1560
                            Height =240
                            Name ="Label14"
                            Caption ="انتخاب بانك براي چاپ "
                        End
                    End
                End
                Begin Label
                    OverlapFlags =247
                    Left =4005
                    Top =570
                    Width =5130
                    Height =240
                    ForeColor =255
                    Name ="Label16"
                    Caption ="براي حذف به حساب گذاشتن چک, روي شماره سريال چک دابل کليک کنيد"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHKREC_HES.cls"
