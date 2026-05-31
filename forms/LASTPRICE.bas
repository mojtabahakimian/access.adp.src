Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5625
    DatasheetFontHeight =10
    ItemSuffix =18
    Left =4815
    Top =1035
    Right =10440
    Bottom =8265
    HelpContextId =600606
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x433f7bb53346e340
    End
    Caption ="آخرين قيمتها"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnActivate ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =7245
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
                    Left =270
                    Top =1065
                    Width =3493
                    Height =315
                    TabIndex =4
                    Name ="LFP"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3818
                            Top =1079
                            Width =1530
                            Height =300
                            Name ="Label1"
                            Caption ="آخرين قيمت فروش:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =270
                    Top =1500
                    Width =3493
                    Height =315
                    TabIndex =5
                    Name ="LKP"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3818
                            Top =1514
                            Width =1530
                            Height =300
                            Name ="Label3"
                            Caption ="آخرين قيمت خريد:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =270
                    Top =1935
                    Width =3493
                    Height =315
                    TabIndex =6
                    Name ="LP"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3818
                            Top =1949
                            Width =1530
                            Height =300
                            Name ="Label5"
                            Caption ="قيمت اوليه:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =250
                    Top =124
                    Width =3103
                    Height =315
                    TabIndex =1
                    BackColor =13434828
                    Name ="CODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME FROM STUF_DEF ORDER BY STUF_DEF.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =3408
                            Top =138
                            Width =390
                            Height =300
                            Name ="Label7"
                            Caption ="كالا:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3194
                    Top =514
                    Width =1431
                    Height =300
                    TabIndex =2
                    BackColor =13434828
                    Name ="HMOIN1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1 FROM TDETA_HES"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =0
                            Left =4719
                            Top =516
                            Width =660
                            Height =285
                            Name ="Label4"
                            Caption ="مشتري:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =250
                    Top =528
                    Width =2841
                    Height =300
                    TabIndex =3
                    BackColor =13434828
                    Name ="HHMOIN1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + "
                        "'-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes, NAME AS nam, RTRIM(CAST(N_KOL AS "
                        "nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS"
                        " nvarchar)) AS Expr1 FROM TDETA_HES"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =255
                    Top =2340
                    Width =2413
                    Height =315
                    TabIndex =7
                    Name ="LFPM"
                    Format ="#,##0.00\" ريال\";-#,##0.00\" ريال\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2730
                            Top =2355
                            Width =2595
                            Height =300
                            Name ="Label11"
                            Caption ="آخرين قيمت فروش به اين مشتري:"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =226
                    Top =2721
                    Width =1041
                    Height =531
                    TabIndex =8
                    Name ="Command12"
                    Caption ="Command12"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    ControlTipText ="Close Form"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3895
                    Top =154
                    Width =1431
                    Height =300
                    BackColor =13434828
                    Name ="Combo13"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE FROM STUF_DEF ORDER BY CAST(CODE AS NVARCHAR);"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =90
                    Top =45
                    Width =5377
                    Height =857
                    Name ="Box15"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2939
                    Top =2777
                    Width =1798
                    Height =315
                    TabIndex =9
                    Name ="DT"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4799
                            Top =2792
                            Width =540
                            Height =300
                            Name ="Label17"
                            Caption ="تاريخ:"
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    Top =3360
                    Width =5625
                    Height =3675
                    TabIndex =10
                    Name ="TabCtl21"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =3765
                            Width =5355
                            Height =3135
                            Name ="Page22"
                            Caption ="سابقه فروش كالا به مشتري"
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =135
                                    Top =3765
                                    Width =5265
                                    Height =3060
                                    Name ="INVO_LSTf12"
                                    SourceObject ="Form.INVO_LSTf12"
                                    LinkChildFields ="code;cust_no"
                                    LinkMasterFields ="Forms![LASTPRICE]![CODE];Forms![LASTPRICE]![HMOIN1]"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =135
                            Top =3765
                            Width =5355
                            Height =3140
                            Name ="Page23"
                            Caption ="سابقه خريد كالا"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =215
                                    Top =3845
                                    Width =5265
                                    Height =3060
                                    Name ="INVO_LSTf121"
                                    SourceObject ="Form.INVO_LSTf121"
                                    LinkChildFields ="code;cust_no"
                                    LinkMasterFields ="Forms![LASTPRICE]![CODE];Forms![LASTPRICE]![HMOIN1]"
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "LASTPRICE.cls"
