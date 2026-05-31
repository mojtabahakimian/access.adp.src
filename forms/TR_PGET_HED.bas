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
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15307
    DatasheetFontHeight =10
    ItemSuffix =35
    Right =14490
    Bottom =7995
    HelpContextId =5003
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    Filter ="((TR_PGET_HED.ID=120))"
    OrderBy ="TR_PGET_HED.UP_TIME DESC"
    RecSrcDt = Begin
        0x7bdb9bd159b7e340
    End
    RecordSource ="TR_PGET_HED"
    Caption ="خزانه داري"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Line
            Width =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Section
            CanGrow = NotDefault
            Height =8619
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
                    Left =13250
                    Top =56
                    Width =1101
                    Height =315
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="DATE"
                    ControlSource ="DATE"
                    ValidationRule ="Is Not Null And >10101"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14454
                            Top =56
                            Width =585
                            Height =300
                            Name ="Label1"
                            Caption ="تاريخ :"
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
                    Left =4871
                    Top =430
                    Width =9480
                    Height =315
                    TabIndex =1
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14454
                            Top =427
                            Width =840
                            Height =300
                            Name ="Label3"
                            Caption ="ملاحظات :"
                        End
                    End
                End
                Begin Subform
                    Locked = NotDefault
                    OverlapFlags =87
                    Left =61
                    Top =1257
                    Width =15120
                    Height =6330
                    TabIndex =2
                    Name ="TR_PGET_LST_SUB"
                    SourceObject ="Form.TR_PGET_LST_SUB"
                    LinkChildFields ="ID;DATE;UP_DATE;UP_TIME"
                    LinkMasterFields ="ID;DATE;UP_DATE;UP_TIME"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Top =7657
                    Width =1761
                    Height =315
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="MABL"
                    ControlSource ="=[TR_PGET_LST_SUB].[Form]![Text26]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1975
                            Top =7653
                            Width =915
                            Height =300
                            Name ="Label19"
                            Caption ="جمع مبالغ:"
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =11160
                    Top =930
                    Width =1465
                    Height =309
                    TabIndex =6
                    BackColor =65535
                    Name ="Frame15"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =2
                            Left =11371
                            Top =932
                            Width =1035
                            Height =285
                            BackColor =65535
                            Name ="Label16"
                            Caption ="بدهكار "
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =9480
                    Top =935
                    Width =1645
                    Height =309
                    TabIndex =7
                    BackColor =10079487
                    Name ="Frame17"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =2
                            Left =9856
                            Top =946
                            Width =900
                            Height =270
                            BackColor =10079487
                            Name ="Label18"
                            Caption ="بستانكار "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11405
                    Top =60
                    Width =741
                    Height =315
                    TabIndex =8
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="ID"
                    ControlSource ="ID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12194
                            Top =60
                            Width =990
                            Height =240
                            Name ="Label20"
                            Caption ="شماره خزانه:"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2954
                    Top =7749
                    Width =11394
                    Height =300
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =16711680
                    Name ="Text10"
                    ControlSource ="=[TR_PGET_LST_SUB].[Form]![THES]"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, ISNU"
                        "LL(TOTA_HES.NAME, N' ') + N' - ' + ISNULL(DETA_HES.NAME, N' ') + N' - ' + ISNULL"
                        "(TDETA_HES.NAME, N' ') AS Expr1 FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDE"
                        "TA_HES ON DETA_HES.NUMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KO"
                        "L ON TOTA_HES.NUMBER = DETA_HES.N_KOL"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14354
                            Top =7749
                            Width =825
                            Height =300
                            BackColor =65535
                            Name ="Label11"
                            Caption ="به حساب:"
                        End
                    End
                End
                Begin ComboBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2954
                    Top =8124
                    Width =11394
                    Height =300
                    TabIndex =3
                    BackColor =13434828
                    ForeColor =16711680
                    Name ="Text8"
                    ControlSource ="=[TR_PGET_LST_SUB].[Form]![FHES]"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, ISNU"
                        "LL(TOTA_HES.NAME, N' ') + N' - ' + ISNULL(DETA_HES.NAME, N' ') + N' - ' + ISNULL"
                        "(TDETA_HES.NAME, N' ') AS Expr1 FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDE"
                        "TA_HES ON DETA_HES.NUMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KO"
                        "L ON TOTA_HES.NUMBER = DETA_HES.N_KOL"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =1
                            Left =14361
                            Top =8124
                            Width =825
                            Height =300
                            BackColor =10079487
                            Name ="Label9"
                            Caption ="از حساب:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =0
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5084
                    Top =56
                    Width =1026
                    TabIndex =9
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin ComboBox
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8870
                    Top =60
                    Width =1656
                    Height =330
                    TabIndex =10
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="KIND"
                    ControlSource ="KIND"
                    RowSourceType ="Value List"
                    RowSource ="0;\"عادي\";2;\"سند دريافت\";3;\"سند پرداخت\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =10635
                            Top =60
                            Width =690
                            Height =315
                            Name ="Label26"
                            Caption ="نوع برگه:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7130
                    Top =60
                    Width =696
                    Height =315
                    TabIndex =11
                    BackColor =13434828
                    ForeColor =10040115
                    Name ="IDK"
                    ControlSource ="IDK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7874
                            Top =60
                            Width =915
                            Height =240
                            Name ="Label31"
                            Caption ="شماره برگه:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4869
                    Top =56
                    Width =1686
                    Height =345
                    TabIndex =12
                    BackColor =13434828
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6579
                            Top =56
                            Width =465
                            Height =300
                            Name ="Label2"
                            Caption ="كاربر:"
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
                    Left =105
                    Top =63
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    BackColor =13408767
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1637
                            Top =60
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
                    Left =90
                    Top =450
                    Width =2841
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    BackColor =13408767
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Format ="General Date"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3047
                            Top =452
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
                    Left =90
                    Top =885
                    Height =345
                    TabIndex =15
                    BackColor =13408767
                    Name ="UP_USER_NAME"
                    ControlSource ="UP_USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1842
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
                    Left =5612
                    Top =855
                    Width =1431
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =16
                    BackColor =13408767
                    Name ="PC_NAME"
                    ControlSource ="PC_NAME"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7143
                            Top =850
                            Width =1755
                            Height =240
                            Name ="Label152"
                            Caption ="نام كامپيوتر اصلاح كننده:"
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
                    IMESentenceMode =3
                    Left =3510
                    Top =840
                    Width =1929
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =17
                    BackColor =13408767
                    Name ="IPADD"
                    ControlSource ="IPADD"
                    FontName ="Times New Roman"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3161
                            Top =892
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
' See "TR_PGET_HED.cls"
