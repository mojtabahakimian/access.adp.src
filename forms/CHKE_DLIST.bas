Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5190
    RowHeight =360
    ItemSuffix =59
    Left =900
    Top =390
    Right =15120
    Bottom =6810
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xa3a7db2c5b69e540
    End
    RecordSource ="CHKE_DLIST"
    Caption ="ليست چكهاي دريافتي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =14583
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =113
                    Height =255
                    ColumnWidth =1200
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Format ="Fixed"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3855
                            Top =120
                            Width =1095
                            Height =240
                            Name ="Label1"
                            Caption ="*شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =793
                    Height =255
                    ColumnWidth =1155
                    ColumnOrder =3
                    TabIndex =1
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3812
                            Top =793
                            Width =1050
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =1133
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =4
                    TabIndex =2
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3977
                            Top =1133
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="تاريخ دريافت"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1281
                    Top =1474
                    Width =1710
                    Height =255
                    ColumnWidth =1665
                    ColumnOrder =7
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4142
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =1814
                    Height =255
                    ColumnWidth =1485
                    ColumnOrder =2
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4487
                            Top =1814
                            Width =375
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =411
                    Top =2154
                    Width =2580
                    Height =255
                    ColumnWidth =3585
                    ColumnOrder =9
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3687
                            Top =2160
                            Width =1260
                            Height =240
                            Name ="Label13"
                            Caption ="نام  پرداخت كننده"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =2494
                    Height =255
                    ColumnWidth =1800
                    ColumnOrder =10
                    TabIndex =6
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3827
                            Top =2494
                            Width =1035
                            Height =240
                            Name ="Label15"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =2834
                    Height =255
                    ColumnWidth =990
                    ColumnOrder =19
                    TabIndex =7
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3977
                            Top =2834
                            Width =885
                            Height =240
                            Name ="Label17"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =201
                    Top =3174
                    Width =2790
                    Height =255
                    ColumnWidth =900
                    ColumnOrder =6
                    TabIndex =8
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4247
                            Top =3174
                            Width =615
                            Height =240
                            Name ="Label19"
                            Caption ="نام بانك"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1302
                    Top =453
                    Height =255
                    ColumnWidth =675
                    ColumnOrder =0
                    TabIndex =9
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4234
                            Top =450
                            Width =630
                            Height =240
                            Name ="Label20"
                            Caption ="شماره "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1332
                    Top =3465
                    ColumnWidth =1110
                    ColumnOrder =15
                    TabIndex =10
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3897
                            Top =3465
                            Width =1005
                            Height =240
                            Name ="Label22"
                            Caption =" به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1332
                    Top =3805
                    ColumnWidth =1230
                    ColumnOrder =16
                    TabIndex =11
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3732
                            Top =3810
                            Width =1125
                            Height =240
                            Name ="Label23"
                            Caption ="به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1332
                    Top =4145
                    ColumnWidth =1635
                    ColumnOrder =20
                    TabIndex =12
                    Name ="N_KOL2"
                    ControlSource ="N_KOL2"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3462
                            Top =4140
                            Width =1665
                            Height =405
                            Name ="Label24"
                            Caption ="برگشت به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1332
                    Top =4485
                    ColumnWidth =1800
                    ColumnOrder =21
                    TabIndex =13
                    Name ="N_MOIN2"
                    ControlSource ="N_MOIN2"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3597
                            Top =4485
                            Width =1545
                            Height =405
                            Name ="Label25"
                            Caption ="برگشت به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1332
                    Top =4825
                    ColumnWidth =1500
                    ColumnOrder =23
                    TabIndex =14
                    Name ="N_KOL3"
                    ControlSource ="N_KOL3"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3762
                            Top =4830
                            Width =1395
                            Height =240
                            Name ="Label26"
                            Caption ="وصول به حساب كل"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1332
                    Top =5165
                    ColumnWidth =1665
                    ColumnOrder =24
                    TabIndex =15
                    Name ="N_MOIN3"
                    ControlSource ="N_MOIN3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3597
                            Top =5160
                            Width =1560
                            Height =240
                            Name ="Label27"
                            Caption ="وصول به حساب معين"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1076
                    Top =5725
                    ColumnWidth =1995
                    ColumnOrder =22
                    TabIndex =16
                    Name ="N_TAF2"
                    ControlSource ="N_TAF2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3285
                            Top =5730
                            Width =1905
                            Height =240
                            Name ="Label35"
                            Caption ="برگشت به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1076
                    Top =6066
                    ColumnWidth =1860
                    ColumnOrder =25
                    TabIndex =17
                    Name ="N_TAF3"
                    ControlSource ="N_TAF3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3285
                            Top =6060
                            Width =1755
                            Height =240
                            Name ="Label36"
                            Caption ="وصول به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1076
                    Top =6406
                    ColumnWidth =1425
                    ColumnOrder =17
                    TabIndex =18
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3285
                            Top =6405
                            Width =1320
                            Height =240
                            Name ="Label37"
                            Caption ="به حساب تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1303
                    Top =623
                    ColumnWidth =930
                    ColumnOrder =18
                    TabIndex =19
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3289
                            Top =623
                            Width =1245
                            Height =240
                            Name ="Label30"
                            Caption ="نام حساب "
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1020
                    Top =6810
                    Width =2211
                    Height =315
                    ColumnWidth =1185
                    ColumnOrder =11
                    TabIndex =20
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\";7;\"حذف شده\""
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="1"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3362
                            Top =6810
                            Width =1005
                            Height =315
                            Name ="Label50"
                            Caption ="وضعيت چك*"
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
                    ColumnCount =2
                    Left =1020
                    Top =7245
                    Width =2211
                    Height =315
                    ColumnWidth =1575
                    ColumnOrder =12
                    TabIndex =21
                    Name ="SANDUGH"
                    ControlSource ="SANDUGH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = 113) AND (NUMBER = 1)"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    DefaultValue ="1"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3362
                            Top =7245
                            Width =1005
                            Height =315
                            Name ="Label40"
                            Caption ="موقعيت چك*"
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
                    Left =850
                    Top =7653
                    ColumnWidth =1110
                    ColumnOrder =26
                    TabIndex =22
                    Name ="LISTNO"
                    ControlSource ="LISTNO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3480
                            Top =7650
                            Width =1005
                            Height =240
                            Name ="Label41"
                            Caption ="شماره ليست"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =453
                    Top =8340
                    ColumnWidth =690
                    ColumnOrder =27
                    TabIndex =23
                    Name ="KIND"
                    ControlSource ="KIND"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3357
                            Top =8333
                            Width =585
                            Height =240
                            Name ="Label42"
                            Caption ="نوع چك"
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
                    Left =793
                    Top =8053
                    ColumnWidth =810
                    ColumnOrder =28
                    TabIndex =24
                    Name ="SHOB_COD"
                    ControlSource ="SHOB_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3423
                            Top =8050
                            Width =1005
                            Height =240
                            Name ="Label44"
                            Caption ="كد شعبه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =8895
                    ColumnWidth =915
                    ColumnOrder =29
                    TabIndex =25
                    Name ="bhes"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2595
                            Top =8895
                            Width =1005
                            Height =240
                            Name ="Label39"
                            Caption =" به حساب "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =9575
                    ColumnWidth =2445
                    ColumnOrder =30
                    TabIndex =26
                    Name ="brhes"
                    ControlSource ="HES2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2160
                            Top =9570
                            Width =1665
                            Height =405
                            Name ="Label45"
                            Caption ="برگشت به حساب "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =30
                    Top =10255
                    ColumnWidth =1305
                    ColumnOrder =31
                    TabIndex =27
                    Name ="vshes"
                    ControlSource ="HES3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2460
                            Top =10260
                            Width =1395
                            Height =240
                            Name ="Label43"
                            Caption ="وصول به حساب "
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
                    Left =680
                    Top =10544
                    ColumnWidth =495
                    ColumnOrder =5
                    TabIndex =28
                    Name ="modat"
                    ControlSource ="modat"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3497
                            Top =10544
                            Width =585
                            Height =240
                            Name ="Label46"
                            Caption ="مدت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1643
                    Top =10998
                    ColumnWidth =510
                    ColumnOrder =32
                    TabIndex =29
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4640
                            Top =10998
                            Width =405
                            Height =240
                            Name ="Label47"
                            Caption ="بانك:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =11505
                    Width =2580
                    Height =255
                    ColumnWidth =3225
                    ColumnOrder =14
                    TabIndex =30
                    Name ="ESTELAM"
                    ControlSource ="ESTELAM"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3270
                            Top =11505
                            Width =1380
                            Height =240
                            Name ="Label49"
                            Caption ="استعلام(توضيحات).براي ثبت دابل کليک کنيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =850
                    Top =12642
                    ColumnWidth =1050
                    ColumnOrder =34
                    TabIndex =31
                    Name ="ms"
                    ControlSource ="ms"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3915
                            Top =12645
                            Width =945
                            Height =240
                            Name ="Label51"
                            Caption ="ماه سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =850
                    Top =12982
                    ColumnWidth =1035
                    ColumnOrder =33
                    TabIndex =32
                    Name ="ds"
                    ControlSource ="ds"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3945
                            Top =12975
                            Width =930
                            Height =240
                            Name ="Label52"
                            Caption ="روز سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =850
                    Top =13322
                    ColumnWidth =1155
                    ColumnOrder =35
                    TabIndex =33
                    Name ="ys"
                    ControlSource ="ys"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3945
                            Top =13320
                            Width =1050
                            Height =240
                            Name ="Label53"
                            Caption ="سال سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =850
                    Top =13662
                    ColumnWidth =885
                    ColumnOrder =37
                    TabIndex =34
                    Name ="md"
                    ControlSource ="md"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3900
                            Top =13665
                            Width =780
                            Height =240
                            Name ="Label54"
                            Caption ="ماه دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =850
                    Top =14003
                    ColumnWidth =870
                    ColumnOrder =36
                    TabIndex =35
                    Name ="dd"
                    ControlSource ="dd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =14010
                            Width =765
                            Height =240
                            Name ="Label55"
                            Caption ="روز دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =850
                    Top =14343
                    ColumnWidth =990
                    ColumnOrder =38
                    TabIndex =36
                    Name ="yd"
                    ControlSource ="yd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =14340
                            Width =885
                            Height =240
                            Name ="Label56"
                            Caption ="سال دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =850
                    Top =5442
                    ColumnWidth =1260
                    ColumnOrder =8
                    TabIndex =37
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3412
                            Top =5442
                            Width =840
                            Height =240
                            Name ="Label57"
                            Caption ="کد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =737
                    Top =8560
                    Width =3111
                    Height =315
                    ColumnWidth =1785
                    ColumnOrder =13
                    TabIndex =38
                    Name ="SAYADI"
                    ControlSource ="SAYADI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3957
                            Top =8562
                            Width =1095
                            Height =240
                            Name ="Label78"
                            Caption ="شماره صيادي"
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =57
                    ColumnWidth =0
                    Name ="smabl"
                    ControlSource ="=Sum([MABL])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3254
                            Width =375
                            Height =240
                            Name ="Label21"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CHKE_DLIST.cls"
