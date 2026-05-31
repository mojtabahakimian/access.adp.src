Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
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
    Width =14796
    DatasheetFontHeight =10
    ItemSuffix =29
    Top =1170
    Right =11760
    Bottom =6930
    DatasheetGridlinesColor =12632256
    Tag ="SELECT     dbo.DEED_HED.DATE_S, dbo.MO_DTL.*\015\012FROM         dbo.MO_DTL INNE"
        "R JOIN\015\012                      dbo.DEED_HED ON dbo.MO_DTL.N_S = dbo.DEED_HE"
        "D.N_S"
    RecSrcDt = Begin
        0x1a760538f19fe240
    End
    RecordSource ="MOGHHEAD"
    Caption ="صورت مغايرتها"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
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
            Height =8787
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12172
                    Top =61
                    Height =315
                    Name ="MONUM"
                    ControlSource ="MONUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13959
                            Top =61
                            Width =525
                            Height =300
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12170
                    Top =424
                    Height =315
                    TabIndex =1
                    Name ="MODATE"
                    ControlSource ="MODATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13957
                            Top =424
                            Width =480
                            Height =300
                            Name ="Label3"
                            Caption ="تاريخ"
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
                    Left =5961
                    Top =60
                    Width =5535
                    Height =315
                    TabIndex =2
                    Name ="ONVAN"
                    ControlSource ="ONVAN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11582
                            Top =60
                            Width =525
                            Height =300
                            Name ="Label5"
                            Caption ="عنوان"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =966
                    Top =113
                    Height =315
                    TabIndex =4
                    Name ="FDATE"
                    ControlSource ="FDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2753
                            Top =113
                            Width =675
                            Height =300
                            Name ="Label9"
                            Caption ="از تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =963
                    Top =477
                    Height =315
                    TabIndex =5
                    Name ="TDATE"
                    ControlSource ="TDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2750
                            Top =477
                            Width =600
                            Height =300
                            Name ="Label11"
                            Caption ="تاتاريخ"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =53
                    Top =855
                    Width =14625
                    Height =6630
                    TabIndex =6
                    Name ="MO_DTL_SUB"
                    SourceObject ="Form.MO_DTL_SUB"
                    LinkChildFields ="MONUM"
                    LinkMasterFields ="MONUM"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2416
                    Top =7530
                    Width =2010
                    Height =315
                    TabIndex =7
                    BackColor =16764057
                    Name ="SBED"
                    ControlSource ="=[MO_DTL_SUB].[Form]![SBED]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4567
                            Top =7530
                            Width =735
                            Height =300
                            Name ="Label18"
                            Caption ="جمع كل:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =375
                    Top =7530
                    Width =2010
                    Height =315
                    TabIndex =8
                    BackColor =16764057
                    Name ="SBES"
                    ControlSource ="=[MO_DTL_SUB].[Form]![SBES]"
                    Format ="#,###"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2416
                    Top =7890
                    Width =2010
                    Height =315
                    TabIndex =9
                    BackColor =8421376
                    Name ="STBED"
                    Format ="#,###"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4560
                            Top =7890
                            Width =1260
                            Height =300
                            Name ="Label20"
                            Caption ="جمع تيك خورده:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =375
                    Top =7890
                    Width =2010
                    Height =315
                    TabIndex =10
                    BackColor =8421376
                    Name ="STBES"
                    Format ="#,###"
                    DefaultValue ="0"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2416
                    Top =8250
                    Width =2010
                    Height =315
                    TabIndex =11
                    BackColor =10079487
                    Name ="Text22"
                    ControlSource ="=[SBED]-[STBED]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4560
                            Top =8250
                            Width =1380
                            Height =300
                            Name ="Label23"
                            Caption ="مانده تيك نخورده:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =375
                    Top =8250
                    Width =2010
                    Height =315
                    TabIndex =12
                    BackColor =10079487
                    Name ="Text24"
                    ControlSource ="=[SBES]-[STBES]"
                    Format ="#,###"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =8425
                    Top =7591
                    Width =3195
                    Height =465
                    FontWeight =700
                    TabIndex =13
                    Name ="Command25"
                    Caption ="بازخواني حساب باكنترل"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5962
                    Top =457
                    Width =4860
                    Height =345
                    TabIndex =3
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0;3402"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextFontCharSet =0
                            Left =10938
                            Top =457
                            Width =1215
                            Height =300
                            Name ="Label7"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =8430
                    Top =8070
                    Width =3195
                    Height =465
                    FontWeight =700
                    TabIndex =14
                    Name ="Command28"
                    Caption ="چاپ مغايرت"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "MOGHAYERAT.cls"
