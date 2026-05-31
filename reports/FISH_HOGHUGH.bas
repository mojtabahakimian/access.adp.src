Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10666
    DatasheetFontHeight =10
    ItemSuffix =144
    Left =-210
    Top =330
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_BIM___MMO=Forms![F_MENU_BIM]![MMO],@Forms___Baseknow___PERSONEL="
        "getkol(Forms![Baseknow]![PERSONEL] ),@Forms___Baseknow___PERVAM =getkol(Forms![B"
        "aseknow]![PERVAM])"
    RecSrcDt = Begin
        0x10148d1a0f98e440
    End
    InputParameters ="@Forms___F_MENU_BIM___MMO=Forms![F_MENU_BIM]![MMO],@Forms___Baseknow___PERSONEL="
        "getkol(Forms![Baseknow]![PERSONEL] ),@Forms___Baseknow___PERVAM =getkol(Forms![B"
        "aseknow]![PERVAM])"
    RecordSource ="dbo.LIST_SALARY5"
    Caption ="فيش حقوق"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,356 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x420049004d005f005f005f004d004d004f003d0046006f0072006d0073002100 ,
        0x5b0046005f004d0045004e0055005f00420049004d005d0021005b004d004d00 ,
        0x4f005d002c00400046006f0072006d0073005f005f005f004200610073006500 ,
        0x6b006e006f0077005f005f005f0050004500520053004f004e0045004c003d00 ,
        0x6700650074006b006f006c00280046006f0072006d00730021005b0042006100 ,
        0x730065006b006e006f0077005d0021005b0050004500520053004f004e004500 ,
        0x4c005d00200029002c00400046006f0072006d0073005f005f005f0042006100 ,
        0x730065006b006e006f0077005f005f005f00500045005200560041004d002000 ,
        0x3d006700650074006b006f006c00280046006f0072006d00730021005b004200 ,
        0x6100730065006b006e006f0077005d0021005b00500045005200560041004d00 ,
        0x5d002900
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
            ControlSource ="PFAMILY"
        End
        Begin BreakLevel
            ControlSource ="PRF"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =7823
            Name ="Detail"
            Begin
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =1110
                    Width =10630
                    Height =4916
                    Name ="Box55"
                End
                Begin Line
                    BorderWidth =2
                    Left =5280
                    Top =1100
                    Width =0
                    Height =4909
                    Name ="Line56"
                End
                Begin Line
                    BorderWidth =1
                    Left =6510
                    Top =1425
                    Width =0
                    Height =4624
                    Name ="Line57"
                End
                Begin Line
                    BorderWidth =1
                    Left =10215
                    Top =1395
                    Width =0
                    Height =4309
                    Name ="Line65"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10247
                    Top =1517
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label66"
                    Caption ="1"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10247
                    Top =1862
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label67"
                    Caption ="2"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10247
                    Top =2207
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label68"
                    Caption ="3"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10247
                    Top =2552
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label69"
                    Caption ="4"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10247
                    Top =2897
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label70"
                    Caption ="5"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10242
                    Top =3249
                    Width =361
                    Height =330
                    FontWeight =400
                    Name ="Label71"
                    Caption ="6"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10242
                    Top =3583
                    Width =361
                    Height =330
                    FontWeight =400
                    Name ="Label72"
                    Caption ="7"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10242
                    Top =3924
                    Width =361
                    Height =330
                    FontWeight =400
                    Name ="Label73"
                    Caption ="8"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1155
                    Top =1425
                    Width =0
                    Height =4624
                    Name ="Line90"
                End
                Begin Line
                    BorderWidth =1
                    Left =4815
                    Top =1395
                    Width =0
                    Height =4309
                    Name ="Line91"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4862
                    Top =1523
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label92"
                    Caption ="1"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4862
                    Top =1868
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label93"
                    Caption ="2"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4862
                    Top =2213
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label94"
                    Caption ="3"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4862
                    Top =2558
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label95"
                    Caption ="4"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4862
                    Top =2903
                    Width =360
                    Height =330
                    FontWeight =400
                    Name ="Label96"
                    Caption ="5"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4857
                    Top =3255
                    Width =361
                    Height =330
                    FontWeight =400
                    Name ="Label97"
                    Caption ="6"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4857
                    Top =3589
                    Width =361
                    Height =330
                    FontWeight =400
                    Name ="Label98"
                    Caption ="7"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4857
                    Top =3930
                    Width =361
                    Height =330
                    FontWeight =400
                    Name ="Label99"
                    Caption ="8"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =1385
                    Width =10628
                    Name ="Line100"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7470
                    Top =1065
                    Width =945
                    Height =330
                    FontWeight =400
                    Name ="Label101"
                    Caption ="حقوق و مزايا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1710
                    Top =1065
                    Width =885
                    Height =330
                    FontWeight =400
                    Name ="Label102"
                    Caption ="كســــورات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =8935
                    Top =681
                    Width =825
                    Height =330
                    FontWeight =400
                    Name ="CODE Label"
                    Caption ="كد پرسنلي:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="CODE_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =6753
                    Top =681
                    Width =1200
                    Height =330
                    FontWeight =400
                    Name ="PNAME Label"
                    Caption ="نام نام خانوادگي:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="PNAME_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3430
                    Top =681
                    Width =1050
                    Height =330
                    FontWeight =400
                    Name ="KHNOWNUM Label"
                    Caption ="شماره شناسنامه:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="KHNOWNUM_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =1900
                    Top =681
                    Width =570
                    Height =330
                    FontWeight =400
                    Name ="FATHER Label"
                    Caption ="نام پدر:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="FATHER_Label"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =7753
                    Top =285
                    Width =2055
                    Height =330
                    FontWeight =400
                    Name ="Label53"
                    Caption ="صورتحساب حقوق و مزاياي ماه"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    ColumnCount =2
                    Left =7007
                    Top =285
                    Width =735
                    Height =330
                    FontSize =9
                    Name ="MM"
                    ControlSource ="MM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON.MON_ID, MON.MON FROM MON;"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =8001
                    Top =681
                    Width =885
                    Height =330
                    ColumnWidth =660
                    FontSize =9
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4540
                    Top =681
                    Width =2196
                    Height =330
                    FontSize =9
                    TabIndex =2
                    Name ="PRF"
                    ControlSource ="PRF"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =2500
                    Top =681
                    Width =960
                    Height =330
                    ColumnWidth =1200
                    FontSize =9
                    TabIndex =3
                    Name ="KHNOWNUM"
                    ControlSource ="KHNOWNUM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =115
                    Top =681
                    Width =1755
                    Height =330
                    ColumnWidth =600
                    FontSize =9
                    TabIndex =4
                    Name ="FATHER"
                    ControlSource ="FATHER"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =5334
                    Top =285
                    Width =510
                    Height =330
                    ColumnWidth =1170
                    FontSize =9
                    TabIndex =5
                    Name ="DAYS1"
                    ControlSource ="=IIf([MM]<=6,31,IIf([MM]>6 And [MM]<12,30,29))"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5814
                    Top =285
                    Width =1155
                    Height =330
                    FontWeight =400
                    Name ="DAYS Label"
                    Caption ="روزهاي كاركرد:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DAYS_Label"
                End
                Begin Line
                    BorderWidth =2
                    Top =5680
                    Width =10628
                    Name ="Line103"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7650
                    Top =5675
                    Width =885
                    Height =330
                    FontWeight =400
                    Name ="Label104"
                    Caption ="جمع درآمد:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =5355
                    Top =5700
                    Width =1111
                    Height =315
                    FontSize =9
                    TabIndex =6
                    Name ="hogh"
                    ControlSource ="=IIf(IsError([SALARYBEST].[Report]![SVAM]),0,[SALARYBEST].[Report]![SVAM])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2002
                    Top =5675
                    Width =1215
                    Height =330
                    FontWeight =400
                    Name ="Label106"
                    Caption ="جمع كسورات:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =30
                    Top =5700
                    Width =1066
                    Height =315
                    FontSize =9
                    TabIndex =7
                    Name ="kosur"
                    ControlSource ="=IIf(IsError([Child119].[Report]![SVAM]),0,[Child119].[Report]![SVAM])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =90
                    Top =6020
                    Width =1801
                    Height =390
                    FontSize =11
                    TabIndex =8
                    Name ="Text109"
                    ControlSource ="=[hogh]-[kosur]-Round(([hogh]-[kosur])/1000)*1000"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1920
                            Top =6090
                            Width =3285
                            Height =315
                            FontWeight =400
                            Name ="Label108"
                            Caption ="كسر 1000 ريال:"
                            FontName ="WeblogmaYekan"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =7815
                    Width =10666
                    Name ="Line110"
                End
                Begin Subform
                    Left =5325
                    Top =1515
                    Width =4891
                    Height =4153
                    TabIndex =9
                    Name ="SALARYBEST"
                    SourceObject ="Report.SALARYBEST"
                    LinkChildFields ="HES_T"
                    LinkMasterFields ="CODE"
                End
                Begin Subform
                    Left =60
                    Top =1530
                    Width =4741
                    Height =4138
                    TabIndex =10
                    Name ="Child119"
                    SourceObject ="Report.SALARYBED"
                    LinkChildFields ="HES_T"
                    LinkMasterFields ="CODE"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =5760
                    Top =6090
                    Width =2721
                    Height =1001
                    Name ="Box120"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8505
                    Top =6095
                    Width =2130
                    Height =255
                    FontWeight =400
                    Name ="Label121"
                    Caption ="محل امضاء و اثر انگشت:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10327
                    Top =4290
                    Width =180
                    Height =330
                    FontWeight =400
                    Name ="Label122"
                    Caption ="9"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10282
                    Top =4620
                    Width =270
                    Height =330
                    FontWeight =400
                    Name ="Label123"
                    Caption ="10"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10282
                    Top =4965
                    Width =270
                    Height =330
                    FontWeight =400
                    Name ="Label124"
                    Caption ="11"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4942
                    Top =4290
                    Width =180
                    Height =330
                    FontWeight =400
                    Name ="Label125"
                    Caption ="9"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4897
                    Top =4635
                    Width =270
                    Height =330
                    FontWeight =400
                    Name ="Label126"
                    Caption ="10"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4897
                    Top =4965
                    Width =270
                    Height =330
                    FontWeight =400
                    Name ="Label127"
                    Caption ="11"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =10267
                    Top =5310
                    Width =270
                    Height =330
                    FontWeight =400
                    Name ="Label128"
                    Caption ="12"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =4882
                    Top =5310
                    Width =270
                    Height =330
                    FontWeight =400
                    Name ="Label129"
                    Caption ="12"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =8560
                    Top =6576
                    TabIndex =11
                    Name ="bedh"
                    ControlSource ="bedh"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =6810
                    Width =2211
                    Height =285
                    TabIndex =12
                    Name ="sfsf"
                    ControlSource ="=PRINT_MORA([WDATE],[CODE])"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =2322
                            Top =6841
                            Width =825
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="Label133"
                            Caption ="مانده مرخصي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =3180
                    Top =6765
                    Width =1656
                    Height =285
                    FontSize =10
                    TabIndex =13
                    Name ="VMAN"
                    ControlSource ="VMAN"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =4873
                            Top =6768
                            Width =855
                            Height =345
                            FontSize =8
                            FontWeight =400
                            Name ="Label136"
                            Caption ="مانده وام:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =60
                    Top =7065
                    Width =9141
                    Height =465
                    FontSize =14
                    TabIndex =14
                    Name ="Text137"
                    ControlSource ="=ALPHANUM(Round(([hogh]-[kosur])/1000)*1000) & \" \" & \"ريال\""
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =60
                    Top =6400
                    Width =1801
                    Height =345
                    FontSize =11
                    TabIndex =15
                    Name ="Text141"
                    ControlSource ="=Round(([hogh]-[kosur])/1000)*1000"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1927
                            Top =6420
                            Width =2445
                            Height =270
                            FontWeight =400
                            Name ="Label142"
                            Caption ="مانده قابل پرداخت:"
                            FontName ="WeblogmaYekan"
                            Tag ="DetachedLabel"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =5268
                    Height =540
                    FontSize =14
                    TabIndex =16
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9240
                    Top =7110
                    Width =1320
                    Height =270
                    FontWeight =400
                    Name ="Label143"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =2
                    NumeralShapes =2
                    Top =7410
                    Width =10551
                    Height =375
                    FontSize =14
                    TabIndex =17
                    Name ="Text144"
                    ControlSource ="=[Forms]![F_MENU_BIM]![payam]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    OldBorderStyle =1
                    TextAlign =3
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =8674
                    Top =6746
                    Width =1801
                    Height =345
                    FontSize =11
                    TabIndex =18
                    Name ="Text148"
                    ControlSource ="=IIf(Round(([hogh]-[kosur])/1000)*1000>0,Round(([hogh]-[kosur])/1000)*1000,0)"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =907
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2778
                    Top =173
                    Width =2265
                    Height =345
                    FontWeight =400
                    Name ="Label130"
                    Caption ="كل  حقوق قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =113
                    Width =2776
                    Height =375
                    FontSize =12
                    Name ="Text131"
                    ControlSource ="=[Text148]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "FISH_HOGHUGH.cls"
