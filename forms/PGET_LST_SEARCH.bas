Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5505
    RowHeight =300
    ItemSuffix =47
    Left =75
    Top =855
    Right =11520
    Bottom =5760
    HelpContextId =5003
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x93cead1d70efe440
    End
    RecordSource ="SELECT PGET_LST.ID, PGET_LST.DATE, PGET_LST.RADIF, PGET_LST.NO_AM, PGET_LST.NAHV"
        "A, PGET_LST.FHES_K, PGET_LST.FHES_M, PGET_LST.THES_K, PGET_LST.THES_M, PGET_LST."
        "SHARH, PGET_LST.MABL, PGET_LST.N_SERI, PGET_LST.BANK, ISNULL(TOTA_HES_1.NAME, N'"
        " ') + N'-' + ISNULL(DETA_HES_1.NAME, N' ') + N'-' + ISNULL(TDETA_HES_1.NAME, N' "
        "') AS FHES, ISNULL(TOTA_HES_1.NAME, N' ') + N'-' + ISNULL(DETA_HES_1.NAME, N' ')"
        " + N'-' + ISNULL(TDETA_HES_1.NAME, N' ') AS THES, PGET_LST.FHES_T, PGET_LST.THES"
        "_T, ISNULL(PAY_GETD.DATE_S, PAY_GETP.DATE_S) AS dates, PGET_LST.FHES AS Expr1, P"
        "GET_LST.THES AS Expr2, PGET_LST.ARZD, PGET_LST.FHES_T2, PGET_LST.THES_T2, PGET_L"
        "ST.FHES_T3, PGET_LST.THES_T3, PGET_LST.FHES_T4, PGET_LST.THES_T4 FROM TOTA_HES T"
        "OTA_HES_1 INNER JOIN DETA_HES DETA_HES_1 INNER JOIN TDETA_HES TDETA_HES_1 ON DET"
        "A_HES_1.NUMBER = TDETA_HES_1.NUMBER AND DETA_HES_1.N_KOL = TDETA_HES_1.N_KOL ON "
        "TOTA_HES_1.NUMBER = DETA_HES_1.N_KOL INNER JOIN TOTA_HES TOTA_HES_2 INNER JOIN D"
        "ETA_HES DETA_HES_2 ON TOTA_HES_2.NUMBER = DETA_HES_2.N_KOL INNER JOIN TDETA_HES "
        "TDETA_HES_2 INNER JOIN PGET_LST ON TDETA_HES_2.TNUMBER = PGET_LST.FHES_T AND TDE"
        "TA_HES_2.NUMBER = PGET_LST.FHES_M AND TDETA_HES_2.N_KOL = PGET_LST.FHES_K ON DET"
        "A_HES_2.NUMBER = TDETA_HES_2.NUMBER AND DETA_HES_2.N_KOL = TDETA_HES_2.N_KOL ON "
        "TDETA_HES_1.TNUMBER = PGET_LST.THES_T AND TDETA_HES_1.NUMBER = PGET_LST.THES_M A"
        "ND TDETA_HES_1.N_KOL = PGET_LST.THES_K LEFT OUTER JOIN PAY_GETP ON PGET_LST.N_SE"
        "RI = PAY_GETP.N_SERI AND PGET_LST.BANK = PAY_GETP.BANK LEFT OUTER JOIN PAY_GETD "
        "ON PGET_LST.N_SERI = PAY_GETD.N_SERI AND PGET_LST.BANK = PAY_GETD.BANk"
    Caption ="جستجو در شرح  عملكرد خزانه"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =19672
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =30
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =1
                    Name ="DATE"
                    ControlSource ="DATE"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4667
                            Top =30
                            Width =390
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =370
                    Height =255
                    ColumnWidth =510
                    ColumnOrder =2
                    TabIndex =1
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4652
                            Top =370
                            Width =405
                            Height =240
                            Name ="Label3"
                            Caption ="رديف"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =141
                    Top =2751
                    Width =3045
                    Height =255
                    ColumnWidth =5460
                    ColumnOrder =15
                    TabIndex =8
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4637
                            Top =2751
                            Width =420
                            Height =240
                            Name ="Label17"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =3091
                    Height =255
                    ColumnWidth =1710
                    ColumnOrder =6
                    TabIndex =9
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4682
                            Top =3091
                            Width =375
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =3431
                    Height =255
                    ColumnWidth =1425
                    ColumnOrder =16
                    TabIndex =10
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4052
                            Top =3431
                            Width =1005
                            Height =240
                            Name ="Label21"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1485
                    Top =710
                    Height =255
                    ColumnWidth =960
                    ColumnOrder =3
                    TabIndex =2
                    Name ="NO_AM"
                    ControlSource ="NO_AM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_DPS.CODE, TCOD_DPS.NAMES FROM TCOD_DPS ORDER BY TCOD_DPS.CODE, TCOD_"
                        "DPS.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4202
                            Top =710
                            Width =855
                            Height =240
                            Name ="Label5"
                            Caption ="نوع عمليات"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1485
                    Top =1050
                    Height =255
                    ColumnWidth =1290
                    ColumnOrder =4
                    TabIndex =3
                    Name ="NAHVA"
                    ControlSource ="NAHVA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_DPSKIND.CODE, TCOD_DPSKIND.NAMES FROM TCOD_DPSKIND ORDER BY TCOD_DPS"
                        "KIND.CODE, TCOD_DPSKIND.NAMES;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4682
                            Top =1050
                            Width =375
                            Height =240
                            Name ="Label7"
                            Caption ="نحوه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1500
                    Top =4071
                    Height =255
                    ColumnWidth =4905
                    ColumnOrder =13
                    TabIndex =12
                    Name ="FHES"
                    ControlSource ="FHES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4395
                            Top =4065
                            Width =690
                            Height =240
                            Name ="Label24"
                            Caption ="از حساب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1500
                    Top =4412
                    Height =255
                    ColumnWidth =4905
                    ColumnOrder =14
                    TabIndex =13
                    Name ="THES"
                    ControlSource ="THES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4395
                            Top =4410
                            Width =720
                            Height =240
                            Name ="Label25"
                            Caption ="به حساب"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =226
                    Width =397
                    Height =114
                    ColumnWidth =1417
                    ColumnOrder =18
                    TabIndex =14
                    Name ="upd"
                    DefaultValue ="False"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1530
                    Top =4875
                    ColumnWidth =1605
                    ColumnOrder =0
                    TabIndex =17
                    Name ="ID"
                    ControlSource ="ID"
                    OnDblClick ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3312
                            Top =4875
                            Width =1620
                            Height =240
                            Name ="Label36"
                            Caption ="شماره دريافت پرداخت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =116
                    Top =1216
                    Width =3126
                    Height =255
                    ColumnWidth =885
                    ColumnOrder =9
                    TabIndex =16
                    Name ="THES_T"
                    ControlSource ="THES_T"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3988
                            Top =1216
                            Width =1125
                            Height =240
                            Name ="Label37"
                            Caption ="به تفضيلي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1391
                    Width =3216
                    Height =255
                    ColumnWidth =660
                    ColumnOrder =10
                    TabIndex =4
                    Name ="FHES_K"
                    ControlSource ="FHES_K"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4127
                            Top =1391
                            Width =930
                            Height =240
                            Name ="Label9"
                            Caption ="از  كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1731
                    Width =3186
                    Height =255
                    ColumnWidth =825
                    ColumnOrder =11
                    TabIndex =5
                    Name ="FHES_M"
                    ControlSource ="FHES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3917
                            Top =1731
                            Width =1140
                            Height =240
                            Name ="Label11"
                            Caption =" از  معين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =42
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2071
                    Width =3186
                    Height =255
                    ColumnWidth =660
                    ColumnOrder =7
                    TabIndex =6
                    Name ="THES_K"
                    ControlSource ="THES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4097
                            Top =2071
                            Width =960
                            Height =240
                            Name ="Label13"
                            Caption ="به  كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2411
                    Width =3126
                    Height =255
                    ColumnWidth =780
                    ColumnOrder =8
                    TabIndex =7
                    Name ="THES_M"
                    ControlSource ="THES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3932
                            Top =2411
                            Width =1125
                            Height =240
                            Name ="Label15"
                            Caption ="به  معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =-6
                    Top =396
                    Width =3186
                    Height =255
                    ColumnWidth =885
                    ColumnOrder =12
                    TabIndex =15
                    Name ="FHES_T"
                    ControlSource ="FHES_T"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3917
                            Width =1140
                            Height =240
                            Name ="Label35"
                            Caption =" ازتفضيلي"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1485
                    Top =3772
                    ColumnWidth =1905
                    ColumnOrder =17
                    TabIndex =11
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4682
                            Top =3772
                            Width =375
                            Height =240
                            Name ="Label23"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1133
                    Top =5102
                    ColumnWidth =1155
                    ColumnOrder =5
                    TabIndex =18
                    Name ="dates"
                    ControlSource ="dates"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3990
                            Top =5100
                            Width =1050
                            Height =240
                            Name ="Label38"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =5329
                    TabIndex =19
                    Name ="FHES_T2"
                    ControlSource ="FHES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3570
                            Top =5325
                            Width =915
                            Height =240
                            Name ="Label39"
                            Caption ="از تفصيلي 2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =5669
                    TabIndex =20
                    Name ="THES_T2"
                    ControlSource ="THES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =5670
                            Width =945
                            Height =240
                            Name ="Label40"
                            Caption ="به تفصيلي 2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =6009
                    TabIndex =21
                    Name ="FHES_T3"
                    ControlSource ="FHES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =6015
                            Width =915
                            Height =240
                            Name ="Label41"
                            Caption ="از تفصيلي 3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =6349
                    TabIndex =22
                    Name ="THES_T3"
                    ControlSource ="THES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =6345
                            Width =945
                            Height =240
                            Name ="Label42"
                            Caption ="به تفصيلي 3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =6689
                    TabIndex =23
                    Name ="FHES_T4"
                    ControlSource ="FHES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =6690
                            Width =915
                            Height =240
                            Name ="Label43"
                            Caption ="از تفصيلي 4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =7029
                    TabIndex =24
                    Name ="THES_T4"
                    ControlSource ="THES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3645
                            Top =7035
                            Width =945
                            Height =240
                            Name ="Label44"
                            Caption ="به تفصيلي 4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =7596
                    TabIndex =25
                    Name ="Expr1"
                    ControlSource ="Expr1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3869
                            Top =7599
                            Width =690
                            Height =240
                            Name ="Label45"
                            Caption ="از حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =7936
                    TabIndex =26
                    Name ="Expr2"
                    ControlSource ="Expr2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3869
                            Top =7929
                            Width =720
                            Height =240
                            Name ="Label46"
                            Caption ="به حساب"
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
' See "PGET_LST_SEARCH.cls"
