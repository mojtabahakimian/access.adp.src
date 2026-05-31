Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8031
    RowHeight =270
    ItemSuffix =28
    Left =405
    Top =1590
    Right =11295
    Bottom =6375
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7bf9624531cce240
    End
    RecordSource ="CHK_SODUR"
    Caption ="CHK_SODUR"
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
        Begin PageHeader
            DisplayWhen =1
            Height =0
            BackColor =-2147483633
            Name ="PageHeader"
        End
        Begin Section
            Height =3194
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4208
                    Top =75
                    Height =315
                    ColumnWidth =1215
                    ColumnOrder =0
                    FontWeight =700
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    OnExit ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7270
                            Top =75
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="سريال"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3419
                    Top =415
                    Width =2490
                    Height =315
                    ColumnWidth =2985
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =1
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TDETA_HES.TNUMBER, RTRIM(CAST(COD_HESAB.CODE AS NVARCHAR)) + '-' + TDETA_"
                        "HES.NAME + ' شعبه ' + RTRIM(CAST(COD_HESAB.SHOBEH AS NVARCHAR)) AS HESAB FROM TD"
                        "ETA_HES INNER JOIN COD_HESAB ON TDETA_HES.TNUMBER = COD_HESAB.MOIN WHERE (TDETA_"
                        "HES.N_KOL = 122)"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7030
                            Top =415
                            Width =750
                            Height =240
                            Name ="Label3"
                            Caption ="كد حساب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4208
                    Top =780
                    Height =315
                    ColumnWidth =0
                    ColumnOrder =3
                    FontWeight =700
                    TabIndex =2
                    Name ="SDATE"
                    ControlSource ="DATE"
                    FontName ="Arabic Transparent"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6985
                            Top =780
                            Width =795
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ صدور"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4208
                    Top =1120
                    Height =315
                    ColumnWidth =1050
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =3
                    Name ="DATE_RASID"
                    ControlSource ="DATE_S"
                    FontName ="Arabic Transparent"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6685
                            Top =1120
                            Width =1095
                            Height =240
                            Name ="Label7"
                            Caption ="سر رسيد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4208
                    Top =1461
                    Height =315
                    ColumnWidth =1425
                    ColumnOrder =5
                    FontWeight =700
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###\"ريال\""
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7405
                            Top =1461
                            Width =375
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =944
                    Top =1801
                    Width =4965
                    Height =315
                    ColumnWidth =3090
                    ColumnOrder =6
                    FontWeight =700
                    TabIndex =5
                    Name ="PAYMENT"
                    ControlSource ="NAME_TAH"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7255
                            Top =1801
                            Width =525
                            Height =240
                            Name ="Label11"
                            Caption ="در وجه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4208
                    Top =2141
                    Height =315
                    ColumnWidth =885
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =6
                    Name ="SND_NUM"
                    ControlSource ="N_S"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6895
                            Top =2141
                            Width =885
                            Height =240
                            Name ="Label13"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4308
                    Top =2664
                    TabIndex =7
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6705
                            Top =2664
                            Width =1005
                            Height =240
                            Name ="Label27"
                            Caption ="حساب معين:"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =1134
            BackColor =-2147483633
            Name ="PageFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1871
                    Height =315
                    FontWeight =700
                    Name ="SMABL"
                    ControlSource ="=Sum([MABL])"
                    Format ="#,###\"ريال\""
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1531
                    Top =566
                    Width =2490
                    Height =315
                    FontWeight =700
                    TabIndex =1
                    Name ="CCODH"
                    ControlSource ="=Count([N_MOIN])"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5329
                    Top =340
                    Width =510
                    Height =226
                    TabIndex =2
                    Name ="MOGUDI"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHK_SODUR.cls"
