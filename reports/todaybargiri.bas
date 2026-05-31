Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11001
    DatasheetFontHeight =10
    ItemSuffix =220
    Left =270
    Top =600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5f5a4a7acf6be540
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, TCOD_VAHEDS.NAMES, SUM(INVO_LST.MEGH) AS sm"
        "egh, SUM(INVO_LST.MEGHk) AS smeghk, INVO_LST.ANBAR, stuf_def_nfani.col9 FROM HEA"
        "D_LST INNER JOIN head_lst_log ON HEAD_LST.NUMBER = head_lst_log.NUMBER AND HEAD_"
        "LST.TAG = head_lst_log.TAGG INNER JOIN INVO_LST ON HEAD_LST.NUMBER = INVO_LST.NU"
        "MBER AND HEAD_LST.TAG = INVO_LST.TAG INNER JOIN TCOD_VAHEDS ON INVO_LST.VAHED_K "
        "= TCOD_VAHEDS.CODE INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE LEFT OUT"
        "ER JOIN stuf_def_nfani ON STUF_DEF.CODE = stuf_def_nfani.CODE WHERE (HEAD_LST.TA"
        "G <> 20) GROUP BY STUF_DEF.CODE, STUF_DEF.NAME, TCOD_VAHEDS.NAMES, INVO_LST.ANBA"
        "R, stuf_def_nfani.col9 HAVING (MAX(head_lst_log.UDATEF) = 13981114)"
    Caption ="شمارش روزانه انبار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            SortOrder = NotDefault
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="ANBAR"
        End
        Begin BreakLevel
            ControlSource ="col9"
        End
        Begin BreakLevel
            ControlSource ="NAME"
        End
        Begin PageHeader
            Height =1665
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3855
                    Width =2550
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="شمارش روزانه"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Top =495
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =540
                    Width =1506
                    Height =450
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT2]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1603
                            Top =540
                            Width =660
                            Height =480
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =11001
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =2
                    Left =10965
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10395
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =3391
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    BorderWidth =2
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2145
                    Top =1230
                    Width =1215
                    Height =420
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6865
                    Top =1230
                    Width =2070
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10401
                    Top =1196
                    Width =540
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =1665
                    Width =11001
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Line
                    Left =2115
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line199"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1140
                    Top =1200
                    Width =915
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label201"
                    Caption ="شمارش"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1080
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line207"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =30
                    Top =1185
                    Width =1095
                    Height =405
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label209"
                    Caption ="موجودي سيستم"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8055
                    Top =566
                    Width =2874
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9645
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line212"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9630
                    Top =1200
                    Width =735
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label213"
                    Caption ="کد"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4366
                    Top =1155
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line216"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3465
                    Top =1200
                    Width =915
                    Height =420
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label217"
                    Caption ="مقدار گردش"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =450
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10411
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="RADIF1"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4370
                    Width =5274
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2160
                    Top =15
                    Width =1200
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10986
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10965
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    Left =10395
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    Left =3390
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin Line
                    Left =2115
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line200"
                End
                Begin Line
                    Left =1080
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line208"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =737
                    Width =999
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="NAME1"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9661
                    Top =16
                    Width =684
                    Height =390
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text214"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9645
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line215"
                End
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =15
                    Width =930
                    Height =405
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="smeghk"
                    ControlSource ="smeghk"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4365
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line219"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =1644
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10990
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8847
                    Top =64
                    Width =1551
                    Height =345
                    Name ="USERNAME"
                    ControlSource ="=ucurrentuser()"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =10439
                            Top =56
                            Width =480
                            Height =375
                            FontWeight =400
                            Name ="Label205"
                            Caption ="كاربر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7035
                    Top =65
                    Width =1206
                    Height =345
                    TabIndex =1
                    Name ="Text206"
                    ControlSource ="=Now()"
                    Format ="Short Time"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =8277
                            Top =60
                            Width =510
                            Height =375
                            FontWeight =400
                            Name ="Label207"
                            Caption ="ساعت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6030
                    Top =630
                    Width =1830
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="نام ونام خانوادگي انبار دار:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8400
                    Top =630
                    Width =2115
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label192"
                    Caption ="نام و نام خانوادگي تحويل شمارنده:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9247
                    Top =1013
                    Width =435
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label193"
                    Caption ="امضاء "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6325
                    Top =1050
                    Width =435
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label194"
                    Caption ="امضاء "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =345
                    Top =630
                    Width =1845
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label202"
                    Caption ="نام ونام خانوادگي حسابدار:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2715
                    Top =630
                    Width =2145
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label203"
                    Caption ="نام و نام خانوادگي مدير انبار "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3577
                    Top =1013
                    Width =435
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label204"
                    Caption ="امضاء "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =655
                    Top =1050
                    Width =435
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label206"
                    Caption ="امضاء "
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "todaybargiri.cls"
