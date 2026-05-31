Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10681
    DatasheetFontHeight =10
    ItemSuffix =124
    Left =600
    Top =1500
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6a36de6f55b8e240
    End
    InputParameters ="@Forms___F_MENU_BIM___MMO=Forms![F_MENU_BIM]![MMO],@Forms___Baseknow___PERSONEL="
        "Forms![Baseknow]![PERSONEL] ,@Forms___Baseknow___PERVAM =Forms![Baseknow]![PERVA"
        "M]"
    RecordSource ="dbo.LIST_SALARY5"
    Caption ="فيش حقوق"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,324 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x420049004d005f005f005f004d004d004f003d0046006f0072006d0073002100 ,
        0x5b0046005f004d0045004e0055005f00420049004d005d0021005b004d004d00 ,
        0x4f005d002c00400046006f0072006d0073005f005f005f004200610073006500 ,
        0x6b006e006f0077005f005f005f0050004500520053004f004e0045004c003d00 ,
        0x46006f0072006d00730021005b0042006100730065006b006e006f0077005d00 ,
        0x21005b0050004500520053004f004e0045004c005d0020002c00400046006f00 ,
        0x72006d0073005f005f005f0042006100730065006b006e006f0077005f005f00 ,
        0x5f00500045005200560041004d0020003d0046006f0072006d00730021005b00 ,
        0x42006100730065006b006e006f0077005d0021005b0050004500520056004100 ,
        0x4d005d00
    End
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CODE"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =1124
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =9376
                    Top =786
                    Width =825
                    Height =330
                    Name ="CODE Label"
                    Caption ="كد پرسنلي:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="CODE_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =7194
                    Top =786
                    Width =1200
                    Height =330
                    Name ="PNAME Label"
                    Caption ="نام نام خانوادگي:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="PNAME_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3871
                    Top =786
                    Width =1050
                    Height =330
                    Name ="KHNOWNUM Label"
                    Caption ="شماره شناسنامه:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="KHNOWNUM_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =2341
                    Top =786
                    Width =570
                    Height =330
                    Name ="FATHER Label"
                    Caption ="نام پدر:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="FATHER_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =8194
                    Top =390
                    Width =2055
                    Height =330
                    Name ="Label53"
                    Caption ="صورتحساب حقوق و مزاياي ماه"
                    FontName ="Zar"
                End
                Begin ComboBox
                    TextAlign =3
                    TextFontFamily =10
                    DecimalPlaces =0
                    ColumnCount =2
                    Left =7448
                    Top =390
                    Width =735
                    Height =330
                    FontSize =9
                    FontWeight =700
                    Name ="MM"
                    ControlSource ="MM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON.MON_ID, MON.MON FROM MON;"
                    ColumnWidths ="0"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    NumeralShapes =2
                    Left =8442
                    Top =786
                    Width =885
                    Height =330
                    ColumnWidth =660
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    NumeralShapes =2
                    Left =4981
                    Top =786
                    Width =2196
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    Name ="PRF"
                    ControlSource ="PRF"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    NumeralShapes =2
                    Left =2941
                    Top =786
                    Width =960
                    Height =330
                    ColumnWidth =1200
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    NumeralShapes =2
                    Left =226
                    Top =786
                    Width =2085
                    Height =330
                    ColumnWidth =600
                    FontSize =9
                    FontWeight =700
                    TabIndex =4
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    NumeralShapes =2
                    Left =5775
                    Top =390
                    Width =510
                    Height =330
                    ColumnWidth =1170
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    Name ="DAYS1"
                    ControlSource ="=IIf([MM]<=6,31,IIf([MM]>6 And [MM]<12,30,29))"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =6255
                    Top =390
                    Width =1155
                    Height =330
                    Name ="DAYS Label"
                    Caption ="روزهاي كاركرد:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DAYS_Label"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =1110
                    Width =10651
                    Name ="Line120"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =5433
                    Height =540
                    FontSize =14
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =404
            Name ="Detail"
            Begin
                Begin Subform
                    OldBorderStyle =1
                    BorderWidth =1
                    Left =5250
                    Width =5401
                    Height =404
                    Name ="SALARYBEST"
                    SourceObject ="Report.SALARYBES2"
                    LinkChildFields ="HES_T"
                    LinkMasterFields ="CODE"
                End
                Begin Subform
                    OldBorderStyle =1
                    BorderWidth =1
                    Width =5221
                    Height =404
                    TabIndex =1
                    Name ="Child119"
                    SourceObject ="Report.SALARYBED2"
                    LinkChildFields ="HES_T"
                    LinkMasterFields ="CODE"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1485
            Name ="GroupFooter1"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6915
                    Top =15
                    Width =885
                    Height =330
                    Name ="Label104"
                    Caption ="جمع درآمد:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    NumeralShapes =2
                    Left =5725
                    Top =56
                    Width =1141
                    Height =255
                    FontSize =9
                    FontWeight =700
                    Name ="hogh"
                    ControlSource ="=[SALARYBEST].[Report]![SVAM]"
                    Format ="Standard"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1200
                    Width =1215
                    Height =330
                    Name ="Label106"
                    Caption ="جمع كسورات:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    NumeralShapes =2
                    Left =30
                    Top =60
                    Width =1066
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="kosur"
                    ControlSource ="=[Child119].[Report]![SVAM]"
                    Format ="Standard"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextAlign =2
                    Left =1995
                    Top =420
                    Width =1785
                    Height =330
                    Name ="Label108"
                    Caption ="مانده قابل پرداخت:"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    NumeralShapes =2
                    Left =15
                    Top =420
                    Width =1921
                    Height =315
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    Name ="Text109"
                    ControlSource ="=[hogh]-[kosur]"
                    Format ="Standard"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =2
                    Top =1485
                    Width =10651
                    Name ="Line110"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =10621
                    Name ="Line121"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =5745
                    Top =390
                    Width =2721
                    Height =1001
                    Name ="Box120"
                End
                Begin Label
                    TextAlign =2
                    Left =8475
                    Top =405
                    Width =2130
                    Height =255
                    Name ="Label121"
                    Caption ="محل امضاء و اثر انگشت:"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1643
                    Top =963
                    TabIndex =3
                    Name ="bedh"
                    ControlSource ="bedh"
                End
                Begin TextBox
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3300
                    Top =749
                    Width =1131
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="sfsf"
                    ControlSource ="=PRINT_MORA([WDATE],[CODE])"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =4482
                            Top =690
                            Width =1215
                            Height =345
                            FontSize =10
                            Name ="Label133"
                            Caption ="مانده مرخصي:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    NumeralShapes =2
                    Left =2745
                    Top =1092
                    Width =2061
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="VMAN"
                    ControlSource ="VMAN"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =4843
                            Top =1035
                            Width =855
                            Height =405
                            FontSize =10
                            Name ="Label136"
                            Caption ="مانده وام:"
                            FontName ="Traffic"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1134
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextAlign =2
                    Left =2830
                    Top =510
                    Width =2265
                    Height =345
                    Name ="Label130"
                    Caption ="كل  حقوق قابل پرداخت:"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Top =510
                    Width =2776
                    Height =315
                    FontSize =12
                    FontWeight =700
                    Name ="Text131"
                    ControlSource ="=Abs(Sum([bedh]))"
                    Format ="Standard"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "FISH2.cls"
