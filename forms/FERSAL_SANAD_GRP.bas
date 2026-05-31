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
    Width =6406
    DatasheetFontHeight =10
    ItemSuffix =27
    Left =7545
    Top =3465
    Right =13950
    Bottom =5430
    HelpContextId =132
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x80244eb33dfae140
    End
    Caption ="ارسال اطلاعات"
    HelpFile ="AMIN98.HLP>MAIN"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
            FontName ="MS Sans Serif"
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1984
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =2214
                    Top =1530
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =3
                    Name ="Command2"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3356
                    Top =1530
                    Width =1134
                    Height =340
                    FontWeight =700
                    TabIndex =2
                    Name ="Command13"
                    Caption ="ارسال"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ControlTipText ="اين گزينه را جهت انتقال اطلاعات به ديسك  تاييد نماييد"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3292
                    Top =116
                    Width =1581
                    Height =300
                    Name ="sername"
                    DefaultValue ="[CurrentProject].[Connection].[Properties].[Item](\"Data Source\").[Value]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5005
                            Top =113
                            Width =765
                            Height =330
                            Name ="Label4"
                            Caption ="نام سرور:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2660
                    Top =524
                    Width =2191
                    Height =330
                    TabIndex =1
                    Name ="catalog"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT name AS Expr1, dbid AS Expr2 FROM master.dbo.sysdatabases WHERE (dbid > 5"
                        ")"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4977
                            Top =510
                            Width =1320
                            Height =330
                            Name ="Label8"
                            Caption ="نام بانك اطلاعاتي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =124
                    Top =113
                    Width =1019
                    Height =340
                    TabIndex =4
                    Name ="pent"
                    DefaultValue ="100"
                    FontName ="Tahoma"
                    InputMask ="###\\%"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1257
                            Top =120
                            Width =930
                            Height =240
                            Name ="Label22"
                            Caption ="درصد انتقال:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =5388
                    Top =1077
                    Width =283
                    Height =226
                    TabIndex =5
                    Name ="ARZ_B"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =3
                            Left =4376
                            Top =1050
                            Width =915
                            Height =240
                            Name ="Label24"
                            Caption ="ارزش افزوده"
                            FontName ="Tahoma"
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
                    Left =56
                    Top =525
                    Width =1731
                    Height =300
                    TabIndex =6
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextFontCharSet =0
                            Left =1813
                            Top =541
                            Width =855
                            Height =285
                            Name ="Label1"
                            Caption ="*از تاريخ:"
                            FontName ="Tahoma"
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
                    Left =62
                    Top =945
                    Width =1716
                    Height =300
                    TabIndex =7
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =1814
                            Top =945
                            Width =840
                            Height =285
                            Name ="Label25"
                            Caption ="تا تاريخ:"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FERSAL_SANAD_GRP.cls"
