Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14740
    RowHeight =345
    ItemSuffix =15
    Left =960
    Top =165
    Right =15450
    Bottom =9240
    HelpContextId =47
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbdbfa3324e5ee240
    End
    RecordSource ="TOTA_HES"
    Caption ="حسابهاي كل و معين"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =9070
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
                    Left =11936
                    Top =435
                    Width =1566
                    Height =315
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13628
                            Top =442
                            Width =630
                            Height =300
                            Name ="Label1"
                            Caption ="شماره :"
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
                    Left =9308
                    Top =775
                    Width =4200
                    Height =315
                    ColumnWidth =3450
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13628
                            Top =772
                            Width =1080
                            Height =300
                            Name ="Label3"
                            Caption ="نام حساب كل:"
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
                    Left =6646
                    Top =435
                    Width =1566
                    Height =315
                    TabIndex =3
                    Name ="NO_HES"
                    ControlSource ="NO_HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_HESKIND.CODE, TCOD_HESKIND.NAMES FROM TCOD_HESKIND ORDER BY TCOD_HES"
                        "KIND.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8338
                            Top =437
                            Width =870
                            Height =300
                            Name ="Label5"
                            Caption ="نوع حساب:"
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
                    Left =6646
                    Top =775
                    Width =1566
                    Height =315
                    ColumnWidth =960
                    TabIndex =4
                    Name ="M_D"
                    ControlSource ="M_D"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_HESVAZ.CODE, TCOD_HESVAZ.NAMES FROM TCOD_HESVAZ ORDER BY TCOD_HESVAZ"
                        ".NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8338
                            Top =782
                            Width =900
                            Height =300
                            Name ="Label7"
                            Caption ="موقت دائم:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =67
                    Top =1477
                    Width =14580
                    Height =2505
                    TabIndex =5
                    Name ="DETA_HES_SUB"
                    SourceObject ="Form.DETA_HES_SUB"
                    LinkChildFields ="N_KOL"
                    LinkMasterFields ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13432
                            Top =1177
                            Width =1155
                            Height =300
                            Name ="DETA_HES_SUB Label"
                            Caption ="حسابهاي معين"
                            EventProcPrefix ="DETA_HES_SUB_Label"
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
                    ColumnCount =3
                    ListWidth =1701
                    Left =10642
                    Top =75
                    Width =2871
                    Height =315
                    ColumnWidth =2895
                    Name ="GROUP"
                    ControlSource ="GROUP"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_HESGROUP.CODE, TCOD_HESGROUP.NAMES, TCOD_HESGROUP.CODE FROM TCOD_HES"
                        "GROUP ORDER BY TCOD_HESGROUP.NAMES;"
                    ColumnWidths ="0;2268;567"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13642
                            Top =75
                            Width =945
                            Height =300
                            Name ="Label11"
                            Caption ="گروه حساب:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =67
                    Top =4305
                    Width =14580
                    Height =4710
                    TabIndex =6
                    Name ="Child12"
                    SourceObject ="Form.TDETA_HES_SUB"
                    LinkChildFields ="N_KOL;NUMBER"
                    LinkMasterFields ="[DETA_HES_SUB].Form![N_KOL];[DETA_HES_SUB].Form![NUMBER]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =13312
                            Top =4005
                            Width =1425
                            Height =300
                            Name ="Label13"
                            Caption ="حسابهاي تفضيلي"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =141
                    TextFontCharSet =0
                    Left =1512
                    Top =566
                    Width =1200
                    Height =405
                    FontWeight =700
                    TabIndex =7
                    Name ="Command14"
                    Caption ="&چاپ حساب"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =283
                    Top =567
                    Width =1201
                    Height =405
                    FontWeight =700
                    TabIndex =8
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "TOTA_HES.cls"
