Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
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
    Width =15157
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =405
    Top =1170
    Right =15285
    Bottom =9450
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    OrderBy ="TR_DEED_HED.UP_TIME DESC"
    RecSrcDt = Begin
        0x0c27569353d7e340
    End
    RecordSource ="TR_DEED_HED"
    Caption ="سند حسابداري"
    DatasheetFontName ="Arial (Arabic)"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
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
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            CanGrow = NotDefault
            Height =9660
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12361
                    Top =105
                    Height =315
                    ColumnOrder =11
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14150
                            Top =112
                            Width =990
                            Height =300
                            Name ="Label1"
                            Caption ="شماره سند:"
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
                    Left =12361
                    Top =445
                    Height =315
                    ColumnOrder =10
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    ValidationRule ="Is Not Null And >10101"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14150
                            Top =442
                            Width =825
                            Height =300
                            Name ="Label3"
                            Caption ="تاريخ سند:"
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
                    Left =9246
                    Top =800
                    Width =4815
                    Height =315
                    ColumnOrder =9
                    TabIndex =2
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="SHARH_S"
                    ControlSource ="SHARH_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14149
                            Top =802
                            Width =885
                            Height =300
                            Name ="Label5"
                            Caption ="شرح  سند:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =6292
                    Top =390
                    Width =366
                    Height =300
                    ColumnOrder =8
                    TabIndex =3
                    ForeColor =128
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =6671
                            Top =397
                            Width =705
                            Height =300
                            ForeColor =128
                            Name ="Label7"
                            Caption ="نوع سند:"
                        End
                    End
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =117
                    Top =1310
                    Width =14898
                    Height =7070
                    TabIndex =4
                    Name ="Child14"
                    SourceObject ="Form.TR_DEED_DTL_sub"
                    LinkChildFields ="N_S;UP_DATE;UP_TIME"
                    LinkMasterFields ="N_S;UP_DATE;UP_TIME"
                End
                Begin TextBox
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    BackStyle =0
                    Left =6285
                    Top =45
                    Width =1806
                    Height =300
                    ColumnOrder =7
                    TabIndex =5
                    ForeColor =128
                    Name ="GHATEIt"
                    ControlSource ="=IIf([GHATEI]<>0,\"سند قطعي ميباشد\",\"سند قطعي نيست\")"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5008
                    Top =772
                    Width =351
                    Height =255
                    ColumnOrder =6
                    TabIndex =6
                    Name ="GHATEI"
                    ControlSource ="GHATEI"
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
                    Left =2160
                    Top =8505
                    Height =315
                    ColumnOrder =5
                    TabIndex =7
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="SSBED"
                    ControlSource ="=[Child14].[Form]![SBED]"
                    Format ="#,###"
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
                    Left =435
                    Top =8505
                    Height =315
                    ColumnOrder =4
                    TabIndex =8
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="SSBES"
                    ControlSource ="=[Child14].[Form]![SBES]"
                    Format ="#,###"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9256
                    Top =120
                    Height =315
                    ColumnOrder =3
                    TabIndex =9
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="BASE"
                    ControlSource ="BASE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =11041
                            Top =120
                            Width =915
                            Height =300
                            Name ="Label24"
                            Caption ="شماره مبنا:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10034
                    Top =737
                    Width =351
                    Height =255
                    ColumnOrder =2
                    TabIndex =10
                    Name ="DE"
                    DefaultValue ="False"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3968
                    Top =8505
                    Width =10044
                    Height =300
                    ColumnOrder =1
                    TabIndex =11
                    BackColor =8421376
                    Name ="Text8"
                    ControlSource ="=GETKOLNAME([Child14].[Form]![HES_K]) & \"-\" & GETMOINNAME([Child14].[Form]![HE"
                        "S_K],[Child14].[Form]![HES_M]) & \"-\" & GETTAFNAME([Child14].[Form]![HES_K],[Ch"
                        "ild14].[Form]![HES_M],[Child14].[Form]![HES_T])"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14085
                            Top =8505
                            Width =990
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label9"
                            Caption ="شرح حساب:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3968
                    Top =8850
                    Width =10045
                    Height =300
                    ColumnOrder =12
                    TabIndex =12
                    BackColor =8421376
                    Name ="Text10"
                    ControlSource ="=[Child14].[Form]![SHARH]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14078
                            Top =8850
                            Width =990
                            Height =300
                            BackColor =8421376
                            ForeColor =65535
                            Name ="Label11"
                            Caption ="شرح رديف:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =6292
                    Top =730
                    Width =1686
                    Height =345
                    ColumnOrder =0
                    TabIndex =13
                    BackColor =16777164
                    ForeColor =8388608
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8002
                            Top =730
                            Width =465
                            Height =300
                            Name ="Label2"
                            Caption ="كاربر:"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2595
                    Top =8907
                    TabIndex =14
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2850
                            Top =8850
                            Width =885
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="تنظيم كننده"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2565
                    Top =9161
                    Width =290
                    TabIndex =15
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2849
                            Top =9105
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="مدير مالي"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =2595
                    Top =9420
                    TabIndex =16
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2849
                            Top =9357
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10601
                    Top =564
                    Width =340
                    Height =218
                    TabIndex =17
                    Name ="OKF"
                    ControlSource ="OKF"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10991
                            Top =510
                            Width =975
                            Height =240
                            ForeColor =255
                            Name ="Label146"
                            Caption ="تائيد سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =120
                    Top =78
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1652
                            Top =75
                            Width =900
                            Height =240
                            Name ="Label148"
                            Caption ="تاريخ اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =137
                    Top =485
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3094
                            Top =487
                            Width =1050
                            Height =240
                            Name ="Label149"
                            Caption ="ساعت اصلاح:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =120
                    Top =885
                    Height =345
                    TabIndex =20
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1872
                            Top =885
                            Width =1290
                            Height =240
                            Name ="Label150"
                            Caption ="كاربر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2707
                    Top =80
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4238
                            Top =75
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    TextFontFamily =18
                    ReadingOrder =2
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =3644
                    Top =828
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =22
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3295
                            Top =880
                            Width =285
                            Height =240
                            Name ="Label155"
                            Caption ="IP:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_DEED_HEAD.cls"
