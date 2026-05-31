Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =330
    ItemSuffix =16
    Left =7155
    Top =1425
    Right =21615
    Bottom =7590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x82d8e7e62f65e640
    End
    RecordSource ="SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD, HES, PORID, EMZA, menup, erjabe, "
        "DEFAULT_NAHVA FROM SALA_DTL ORDER BY IDD"
    Caption ="گروه کاربران"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    Moveable =0
    AllowFormView =0
    FetchDefaults =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =7633
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2267
                    Top =510
                    ColumnWidth =2085
                    ColumnOrder =1
                    Name ="SAL_NAME1"
                    ControlSource ="=DECODEUN([SAL_NAME])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =565
                            Top =510
                            Width =1230
                            Height =240
                            Name ="Label0"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2267
                    Top =1871
                    ColumnWidth =675
                    ColumnOrder =0
                    TabIndex =2
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =570
                            Top =1875
                            Width =1635
                            Height =240
                            Name ="Label4"
                            Caption ="شماره "
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2267
                    Top =1190
                    ColumnWidth =1395
                    ColumnOrder =2
                    TabIndex =1
                    Name ="GRSAL"
                    ControlSource ="GRSAL"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مديران\";2;\"حسابداران\";3;\"بازرگاني\";4;\"انبار\";5;\"كارگزيني\";6;\"هتل\""
                        ";7;\"تالار\";8;\"ريموت\";9;\"آفلاين\";11;\"تبلت \""
                    ColumnWidths ="0"
                    StatusBarText ="0بدون گروه 1 فرمانده 2 حسابدار 3 فروش 4 انبار5 كارگزيني10 صاحب"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =565
                            Top =1190
                            Width =930
                            Height =240
                            Name ="Label2"
                            Caption ="گروه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2495
                    Top =2267
                    ColumnWidth =1020
                    TabIndex =3
                    Name ="ENABL"
                    ControlSource ="ENABL"
                    RowSourceType ="Value List"
                    RowSource ="0;\"فعال\";1;\"غير فعال\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =793
                            Top =2267
                            Width =930
                            Height =240
                            Name ="Label13"
                            Caption ="وضعيت"
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
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Width =4221
                    Height =315
                    ColumnWidth =4260
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME, hes AS Expr1 FROM CUST_HESAB ORDER BY NAME"
                    ColumnWidths ="0;2835;567"
                    Tag ="0"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4309
                            Top =4
                            Width =990
                            Height =315
                            ForeColor =8388608
                            Name ="Label19"
                            Caption =" معين معادل"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OverlapFlags =85
                    Left =1247
                    Top =4988
                    Width =2530
                    Height =2645
                    ColumnWidth =1417
                    TabIndex =5
                    Name ="EMZA"
                    ControlSource ="EMZA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3867
                            Top =5838
                            Width =990
                            Height =315
                            ForeColor =8388608
                            Name ="Label12"
                            Caption =" امضاء"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2495
                    Top =2627
                    TabIndex =6
                    Name ="menup"
                    ControlSource ="menup"
                    RowSourceType ="Value List"
                    RowSource ="0;\"استاندارد\";1;\"تصويري\";2;\"پنلي\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =795
                            Top =2625
                            Width =1200
                            Height =240
                            Name ="Label15"
                            Caption ="منوي پيش فرض"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =56
                    Top =3973
                    ColumnWidth =2505
                    TabIndex =7
                    Name ="PORID"
                    ControlSource ="PORID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VISITORS_PORSANT.PORID, CAST(VISITORS_PORSANT.PORID AS nvarchar) + N' - '"
                        " + CAST(VISITORS_PORSANT.VDATE AS nvarchar) + N' - ' + ISNULL(CUSTKIND.CUSTKNAME"
                        ", N'بدون گروه (همه)') + N' - ' + ISNULL(VISITORS_PORSANT.COMMENT, N' ') + N' - '"
                        " + CUST_HESAB.NAME AS Expr1 FROM VISITORS_PORSANT INNER JOIN CUST_HESAB ON VISIT"
                        "ORS_PORSANT.HES = CUST_HESAB.hes LEFT OUTER JOIN CUSTKIND ON VISITORS_PORSANT.CU"
                        "ST_COD = CUSTKIND.CUST_COD"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =2845
                            Top =3968
                            Width =2400
                            Height =240
                            Name ="Label16"
                            Caption ="پيش فرض الگوي پرداخت پورسانت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =984
                    Top =3303
                    Width =2145
                    Height =285
                    ColumnWidth =2115
                    TabIndex =8
                    Name ="erjabe"
                    ControlSource ="erjabe"
                    RowSourceType ="Value List"
                    RowSource ="77;\"خانم برگريزان\""
                    ColumnWidths ="0"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =178
                            TextAlign =2
                            Left =907
                            Top =2948
                            Width =2220
                            Height =330
                            Name ="Label120"
                            Caption ="ارجاع پيش فاکتور به اين کاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Visible = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =1870
                    Top =4478
                    Width =2196
                    Height =330
                    ColumnWidth =1995
                    TabIndex =9
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="DEFAULT_NAHVA"
                    ControlSource ="DEFAULT_NAHVA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PPID, PPAME, MODAT FROM PRICE_PAYNO UNION SELECT 0, 'آزاد', 0"
                    ColumnWidths ="0;1532;289"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4135
                            Top =4478
                            Width =960
                            Height =285
                            ForeColor =8388608
                            Name ="Label150"
                            Caption ="نحوه پرداخت:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_ENDN_SUB.cls"
