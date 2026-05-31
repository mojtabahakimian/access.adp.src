Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11622
    RowHeight =345
    ItemSuffix =43
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x40a24465519de340
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.TAG, HEAD_LST.MAS, INVO_LST.CODE, INVO_LST.MEGH"
        ", INVO_LST.MEGHk, INVO_LST.MABL, INVO_LST.MABL_K, INVO_LST.MANDAH, PTAMIRAT.PIDT"
        ", PTAMIRAT.PSERIAL, PTAMIRAT.PINDATE, PTAMIRAT.PINTIME, PTAMIRAT.PKIND, PTAMIRAT"
        ".MODEL, PTAMIRAT.GRENDDATE, PTAMIRAT.POWNER, PTAMIRAT.POWNERNAM, PTAMIRAT.GRCOMP"
        "A, PTAMIRAT.IYALAT, PTAMIRAT.CITY, PTAMIRAT.PINPKIND, PTAMIRAT.USER_NAME, PTAMIR"
        "AT.PTAHVKIND, PTAMIRAT.EYBEZHAR, PTAMIRAT.LAVAZEM, PTAMIRAT.IDIATEK, PTAMIRAT.AD"
        "RESS, PTAMIRAT.TAMIRKAR1, PTAMIRAT.TAMIRKAR2, PTAMIRAT.TAHV_DATE, PTAMIRAT.TAHV_"
        "TIME, PTAMIRAT.PAYAM, HEAD_LST.TAMIR, HEAD_LST.DATE_N FROM HEAD_LST INNER JOIN P"
        "TAMIRAT ON HEAD_LST.FNUMCO = PTAMIRAT.PIDT INNER JOIN INVO_LST ON HEAD_LST.TAG ="
        " INVO_LST.TAG AND HEAD_LST.NUMBER = INVO_LST.NUMBER WHERE (HEAD_LST.TAMIR = 1010"
        ")"
    Caption ="ليست تعميرات"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin Section
            Height =12843
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3344
                    Top =680
                    ColumnWidth =1065
                    ColumnOrder =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =1643
                            Top =680
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="شماره فاكتور"
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
                    Left =3344
                    Top =2040
                    ColumnWidth =555
                    ColumnOrder =4
                    TabIndex =1
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =2040
                            Width =555
                            Height =240
                            Name ="Label4"
                            Caption ="مقدار"
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
                    Left =3344
                    Top =2381
                    ColumnWidth =795
                    ColumnOrder =5
                    TabIndex =2
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =2381
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="مقدار كل"
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
                    Left =3344
                    Top =2721
                    ColumnWidth =1020
                    ColumnOrder =6
                    TabIndex =3
                    Name ="MABL"
                    ControlSource ="MABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =2721
                            Width =525
                            Height =240
                            Name ="Label6"
                            Caption ="في"
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
                    Left =3344
                    Top =3061
                    ColumnOrder =7
                    TabIndex =4
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =3061
                            Width =705
                            Height =240
                            Name ="Label7"
                            Caption ="مبلغ"
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
                    Left =3344
                    Top =3401
                    ColumnOrder =8
                    TabIndex =5
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1643
                            Top =3401
                            Width =780
                            Height =240
                            Name ="Label8"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8737
                    Top =4091
                    Height =315
                    ColumnOrder =9
                    TabIndex =6
                    BackColor =16777164
                    Name ="PIDT"
                    ControlSource ="PIDT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10495
                            Top =4091
                            Width =1020
                            Height =340
                            Name ="Label1"
                            Caption ="شماره پذيرش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    NumeralShapes =3
                    IMESentenceMode =3
                    Left =5399
                    Top =4084
                    Width =2535
                    Height =340
                    ColumnOrder =10
                    TabIndex =7
                    BackColor =13434828
                    Name ="PSERIAL"
                    ControlSource ="PSERIAL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =8052
                            Top =4084
                            Width =510
                            Height =340
                            Name ="Label33"
                            Caption ="سريال"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =4084
                    Width =4125
                    Height =340
                    ColumnOrder =11
                    TabIndex =8
                    BackColor =13434828
                    Name ="MODEL"
                    ControlSource ="MODEL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MODEL FROM PTAMIRAT GROUP BY MODEL ORDER BY MODEL; "
                    ValidationRule ="<>\"\" And Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =4485
                            Top =4081
                            Width =825
                            Height =360
                            Name ="Label11"
                            Caption ="نام دستگاه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7994
                    Top =4526
                    Width =2616
                    Height =340
                    ColumnOrder =12
                    TabIndex =9
                    Name ="PKIND"
                    ControlSource ="PKIND"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PKINDCO, PKIND FROM PKINDCODING; "
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10716
                            Top =4526
                            Width =795
                            Height =340
                            Name ="Label9"
                            Caption ="نوع پذيرش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5532
                    Top =4538
                    Width =1350
                    Height =340
                    ColumnOrder =13
                    TabIndex =10
                    Name ="PINDATE"
                    ControlSource ="PINDATE"
                    DefaultValue ="=FARSIDATE(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =6943
                            Top =4538
                            Width =885
                            Height =340
                            Name ="Label34"
                            Caption ="تاريخ پذيرش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3410
                    Top =4538
                    Width =801
                    Height =340
                    ColumnOrder =14
                    TabIndex =11
                    Name ="PINTIME"
                    ControlSource ="PINTIME"
                    DefaultValue ="Hour(Now())*100+Minute(Now())"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =4449
                            Top =4538
                            Width =1035
                            Height =340
                            Name ="Label35"
                            Caption ="ساعت پذيرش"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =5669
                    Top =4964
                    Width =4830
                    Height =340
                    ColumnOrder =15
                    TabIndex =12
                    BackColor =13434828
                    Name ="POWNER"
                    ControlSource ="POWNER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.NUMBER, DETA_HES.CODE_E FROM DET"
                        "A_HES WHERE (((DETA_HES.N_KOL)=Forms!Baseknow!BEDEHKAR) And ((DETA_HES.CODE_E) L"
                        "ike IIf(Forms!Baseknow!PERSON=2,\"0\",\"*\") Or (DETA_HES.CODE_E) Is Null)) ORDE"
                        "R BY DETA_HES.NAME; "
                    ColumnWidths ="0;3689;1134"
                    ValidationRule ="<>\"\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10605
                            Top =4964
                            Width =945
                            Height =340
                            Name ="Label15"
                            Caption ="مالك دستگاه"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =292
                    Top =4961
                    Width =4320
                    Height =340
                    ColumnOrder =16
                    TabIndex =13
                    Name ="POWNERNAM"
                    ControlSource ="POWNERNAM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT POWNERNAM FROM PTAMIRAT GROUP BY POWNERNAM ORDER BY POWNERNAM; "
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =4706
                            Top =4961
                            Width =900
                            Height =340
                            Name ="Label17"
                            Caption ="نماينده مالك"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5669
                    Top =5392
                    Width =4380
                    Height =340
                    ColumnOrder =17
                    TabIndex =14
                    Name ="GRCOMPA"
                    ControlSource ="GRCOMPA"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT GRCOMPA FROM PTAMIRAT GROUP BY GRCOMPA ORDER BY GRCOMPA; "
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10095
                            Top =5392
                            Width =1485
                            Height =340
                            Name ="Label19"
                            Caption ="شركت گارانتي كننده"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3018
                    Top =5381
                    Width =2040
                    Height =340
                    ColumnOrder =18
                    TabIndex =15
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IYALAT FROM PTAMIRAT GROUP BY IYALAT ORDER BY IYALAT; "
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =5104
                            Top =5381
                            Width =495
                            Height =340
                            Name ="Label21"
                            Caption ="استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8904
                    Top =5801
                    Height =340
                    ColumnOrder =19
                    TabIndex =16
                    Name ="PINPKIND"
                    ControlSource ="PINPKIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"در محل تعميرگاه\";2;\"در محل مشتري\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10660
                            Top =5801
                            Width =885
                            Height =340
                            Name ="Label25"
                            Caption ="محل پذيرش"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =292
                    Top =6250
                    Width =5205
                    Height =1170
                    ColumnWidth =1665
                    ColumnOrder =20
                    TabIndex =17
                    Name ="EYBEZHAR"
                    ControlSource ="EYBEZHAR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =178
                            Left =1992
                            Top =5966
                            Width =1560
                            Height =240
                            Name ="Label29"
                            Caption ="عيب به اظهار مشتري"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5677
                    Top =6251
                    Width =4905
                    Height =1185
                    ColumnOrder =21
                    TabIndex =18
                    Name ="LAVAZEM"
                    ControlSource ="LAVAZEM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10666
                            Top =6251
                            Width =870
                            Height =240
                            Name ="Label31"
                            Caption ="لوازم همراه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =7516
                    Width =10260
                    Height =340
                    ColumnOrder =22
                    TabIndex =19
                    Name ="IDIATEK"
                    ControlSource ="IDIATEK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10607
                            Top =7516
                            Width =915
                            Height =340
                            Name ="Label36"
                            Caption ="نظر تكنسين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =7933
                    Width =10260
                    Height =340
                    ColumnOrder =23
                    TabIndex =20
                    BackColor =13434828
                    Name ="ADRESS"
                    ControlSource ="ADRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10654
                            Top =7933
                            Width =930
                            Height =340
                            Name ="Label37"
                            Caption ="آدرس و تلفن"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =296
                    Top =8341
                    Width =10230
                    Height =1005
                    ColumnOrder =24
                    TabIndex =21
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10557
                            Top =8341
                            Width =1065
                            Height =340
                            Name ="Label38"
                            Caption ="متن زير پذيرش"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5662
                    Top =5816
                    Width =2166
                    Height =340
                    ColumnOrder =26
                    TabIndex =24
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="=UCURRENTUSER()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =7882
                            Top =5811
                            Width =960
                            Height =330
                            Name ="Label27"
                            Caption ="كاربر پذيرش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9260
                    Top =9819
                    Width =1350
                    Height =340
                    ColumnOrder =27
                    TabIndex =25
                    Name ="TAHV_DATE"
                    ControlSource ="TAHV_DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10671
                            Top =9819
                            Width =885
                            Height =340
                            Name ="Label43"
                            Caption ="تاريخ  تحويل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7213
                    Top =9819
                    Width =801
                    Height =340
                    ColumnOrder =28
                    TabIndex =26
                    Name ="TAHV_TIME"
                    ControlSource ="TAHV_TIME"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =8177
                            Top =9819
                            Width =1035
                            Height =340
                            Name ="Label45"
                            Caption ="ساعت تحويل"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3912
                    Top =5955
                    ColumnWidth =1275
                    ColumnOrder =29
                    TabIndex =27
                    Name ="Text64"
                    ControlSource ="TAMIRKAR1"
                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextFontCharSet =178
                            Left =6673
                            Top =5956
                            Width =1020
                            Height =240
                            Name ="Label65"
                            Caption ="كد تعمير كار 2"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3345
                    Top =6012
                    ColumnWidth =1230
                    ColumnOrder =30
                    TabIndex =28
                    Name ="Text66"
                    ControlSource ="TAMIRKAR2"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =178
                            Left =6103
                            Top =6016
                            Width =1065
                            Height =240
                            Name ="Label67"
                            Caption ="كد تعمير كار 2 "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =1575
                    Height =340
                    ColumnWidth =1425
                    ColumnOrder =31
                    TabIndex =29
                    Name ="GRENDDATE"
                    ControlSource ="GRENDDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =178
                            Left =1738
                            Width =1320
                            Height =340
                            Name ="Label13"
                            Caption ="تاريخ اتمام گارانتي"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =866
                    Width =1695
                    Height =340
                    ColumnOrder =32
                    TabIndex =30
                    Name ="CITY"
                    ControlSource ="CITY"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITY FROM PTAMIRAT GROUP BY CITY ORDER BY CITY; "
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =178
                            Left =1868
                            Top =866
                            Width =780
                            Height =340
                            Name ="Label23"
                            Caption ="شهرستان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6191
                    Top =9406
                    Width =4515
                    Height =340
                    ColumnWidth =870
                    ColumnOrder =0
                    TabIndex =22
                    Name ="TAMIRKAR1"
                    ControlSource ="TAMIRKAR1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10767
                            Top =9406
                            Width =765
                            Height =340
                            Name ="Label39"
                            Caption ="تعمير كار1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =9406
                    Width =4950
                    Height =340
                    ColumnOrder =25
                    TabIndex =23
                    Name ="TAMIRKAR2"
                    ControlSource ="TAMIRKAR2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =5324
                            Top =9406
                            Width =810
                            Height =340
                            Name ="Label41"
                            Caption ="تعمير كار 2"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =3
                    ColumnCount =3
                    Width =1680
                    Height =255
                    ColumnWidth =3825
                    ColumnOrder =3
                    TabIndex =31
                    Name ="CODE"
                    ControlSource ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME; "
                    ColumnWidths ="0;2268;567"
                    ValidationRule ="Is Not Null"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =178
                            Left =3041
                            Width =510
                            Height =240
                            Name ="Label40"
                            Caption =" كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9260
                    Top =10279
                    Width =1350
                    Height =340
                    ColumnWidth =975
                    ColumnOrder =2
                    TabIndex =32
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =10671
                            Top =10279
                            Width =885
                            Height =340
                            Name ="Label42"
                            Caption ="تاريخ  فاكتور"
                        End
                    End
                End
            End
        End
    End
End
