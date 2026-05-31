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
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11565
    RowHeight =390
    ItemSuffix =30
    Left =3630
    Top =1035
    Right =15195
    Bottom =8415
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x181d2d454aefe240
    End
    RecordSource ="PVAM"
    Caption ="وام پرسنل"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
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
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
        Begin FormHeader
            Height =660
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    TextFontFamily =2
                    Left =3409
                    Top =15
                    Width =5130
                    Height =645
                    FontSize =18
                    ForeColor =10092543
                    Name ="Label8"
                    Caption ="وام پرسنل"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    TextFontFamily =2
                    Left =3424
                    Width =5130
                    Height =645
                    FontSize =18
                    Name ="Label10"
                    Caption ="وام پرسنل"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6468
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7757
                    Top =2449
                    Width =1956
                    Height =315
                    ColumnWidth =1305
                    ColumnOrder =2
                    TabIndex =2
                    BackColor =16777164
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =1
                            Left =9799
                            Top =2449
                            Width =375
                            Height =240
                            Name ="Label5"
                            Caption ="مبلغ"
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
                    ColumnCount =2
                    Left =3291
                    Top =285
                    Width =4821
                    Height =315
                    ColumnWidth =3975
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVARCHAR)) AS PE"
                        "R FROM PERSONEL ORDER BY PNAME + N' ' + PFAMILY + N' ' + RTRIM(CAST(CODE AS NVAR"
                        "CHAR))"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8213
                            Top =285
                            Width =675
                            Height =300
                            Name ="Label1"
                            Caption ="پرسنل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =4041
                    Top =1140
                    Width =7437
                    Height =4074
                    Name ="Box15"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4202
                    Top =1954
                    Width =5511
                    Height =360
                    ColumnWidth =3315
                    ColumnOrder =1
                    TabIndex =1
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHARH FROM PVAM GROUP BY SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =1
                            Left =9799
                            Top =1954
                            Width =420
                            Height =240
                            Name ="Label3"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8854
                    Top =1476
                    Width =816
                    Height =315
                    ColumnWidth =945
                    TabIndex =3
                    Name ="VAM_ID"
                    ControlSource ="VAM_ID"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =1
                            Left =9752
                            Top =1474
                            Width =840
                            Height =240
                            Name ="Label18"
                            Caption ="شمار ه وام"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7757
                    Top =2914
                    Width =1956
                    Height =315
                    ColumnWidth =1710
                    TabIndex =4
                    BackColor =16777164
                    Name ="DATE_BP"
                    ControlSource ="DATE_BP"
                    OnExit ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =1
                            Left =9797
                            Top =2914
                            Width =1605
                            Height =240
                            Name ="Label20"
                            Caption ="باز پرداخت از تاريخ(ماه)"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =351
                    Top =1140
                    Width =3567
                    Height =4074
                    Name ="Box21"
                End
                Begin Subform
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =517
                    Top =1190
                    Width =3175
                    Height =3908
                    TabIndex =5
                    Name ="PVAM_BAZ_SUB"
                    SourceObject ="Form.PVAM_BAZ_SUB"
                    LinkChildFields ="CODE;VAM_ID"
                    LinkMasterFields ="CODE;VAM_ID"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =621
                    Top =5280
                    Width =1956
                    Height =315
                    TabIndex =6
                    BackColor =16777164
                    Name ="BMAB"
                    ControlSource ="=IIf(IsError([PVAM_BAZ_SUB].[Form]![SMABL]),0,[PVAM_BAZ_SUB].[Form]![SMABL])"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =2661
                            Top =5280
                            Width =1065
                            Height =240
                            Name ="Label25"
                            Caption ="جمع بازپرداخت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7757
                    Top =3349
                    Width =1956
                    Height =315
                    TabIndex =7
                    BackColor =16777164
                    Name ="MABLBZ"
                    ControlSource ="MABLBZ"
                    Format ="Standard"
                    BeforeUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =1
                            Left =9797
                            Top =3349
                            Width =1605
                            Height =240
                            Name ="Label27"
                            Caption ="مبلغ هر قسط"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7757
                    Top =3784
                    Width =1956
                    Height =315
                    TabIndex =8
                    BackColor =16777164
                    Name ="NUM"
                    ControlSource ="NUM"
                    Format ="Standard"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            TextAlign =1
                            Left =9797
                            Top =3784
                            Width =1605
                            Height =240
                            Name ="Label29"
                            Caption ="تعداد اقساط"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1247
                    Top =226
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =9
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4024
                    Top =683
                    ColumnWidth =2070
                    TabIndex =10
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6229
                            Top =680
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3968
                    Top =963
                    TabIndex =11
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6173
                            Top =965
                            Width =1365
                            Height =240
                            Name ="Label36"
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
' See "PVAM_KAR.cls"
