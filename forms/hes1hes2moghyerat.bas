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
    TabularFamily =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6576
    RowHeight =405
    DatasheetFontHeight =10
    ItemSuffix =6
    DatasheetForeColor =16777215
    DatasheetBackColor =255
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xd4db3dd40ee0e540
    End
    RecordSource ="SELECT HEAD_LST_1.NUMBER AS nb, HEAD_LST_1.TAG AS tg2, HEAD_LST_1.DATE_N AS dt, "
        "HEAD_LST.CUST_NO, HEAD_LST_1.CUST_NO AS mf, CUST_HESAB.NAME FROM CUST_HESAB INNE"
        "R JOIN HEAD_LST HEAD_LST_1 ON CUST_HESAB.hes = HEAD_LST_1.CUST_NO LEFT OUTER JOI"
        "N HEAD_LST ON HEAD_LST_1.NUMBER = HEAD_LST.NUMBER AND HEAD_LST_1.TAG - 11 = HEAD"
        "_LST.TAG AND HEAD_LST_1.CUST_NO = HEAD_LST.CUST_NO WHERE (HEAD_LST.CUST_NO IS NU"
        "LL) AND (HEAD_LST_1.TAG = 12 OR HEAD_LST_1.TAG = 13)"
    Caption ="ليست فاکتورهايي که کد مشتري در حواله و رسيد با فاکتور متفاوت است"
    DatasheetFontName ="WeblogmaYekan"
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
            Height =2850
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3288
                    Top =907
                    ColumnWidth =975
                    Name ="nb"
                    ControlSource ="nb"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1590
                            Top =900
                            Width =930
                            Height =240
                            Name ="Label0"
                            Caption ="شماره حواله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3288
                    Top =1587
                    TabIndex =2
                    Name ="dt"
                    ControlSource ="dt"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1590
                            Top =1590
                            Width =390
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3288
                    Top =2267
                    TabIndex =3
                    Name ="mf"
                    ControlSource ="mf"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1590
                            Top =2265
                            Width =840
                            Height =240
                            Name ="Label4"
                            Caption ="کد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3288
                    Top =2607
                    ColumnWidth =7020
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1590
                            Top =2610
                            Width =1380
                            Height =240
                            Name ="Label5"
                            Caption ="نام حساب مشتري"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3288
                    Top =1247
                    TabIndex =1
                    Name ="tg2"
                    ControlSource ="tg2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, BARGAH FROM TAGCOD"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1590
                            Top =1245
                            Width =630
                            Height =240
                            Name ="Label1"
                            Caption ="نوع برگه"
                        End
                    End
                End
            End
        End
    End
End
