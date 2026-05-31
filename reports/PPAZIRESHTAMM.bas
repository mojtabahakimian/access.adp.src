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
    Width =10829
    DatasheetFontHeight =10
    ItemSuffix =388
    Left =270
    Top =1215
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x314768391b83e340
    End
    RecordSource ="SELECT PTAMIRAT.PIDT, PTAMIRAT.PSERIAL, PTAMIRAT.PINDATE, PTAMIRAT.PINTIME, PKIN"
        "DCODING.PKIND, PTAMIRAT.MODEL, PTAMIRAT.GRENDDATE, PTAMIRAT.POWNERNAM, PTAMIRAT."
        "GRCOMPA, PTAMIRAT.IYALAT, PTAMIRAT.CITY, PTAMIRAT.PINPKIND, PTAMIRAT.USER_NAME, "
        "PTAMIRAT.PTAHVKIND, PTAMIRAT.EYBEZHAR, PTAMIRAT.LAVAZEM, PTAMIRAT.IDIATEK, PTAMI"
        "RAT.ADRESS, PTAMIRAT.PAYAM, CUST_HESAB.NAME FROM PTAMIRAT INNER JOIN PKINDCODING"
        " ON PTAMIRAT.PKIND = PKINDCODING.PKINDCO INNER JOIN CUST_HESAB ON PTAMIRAT.POWNE"
        "R = CUST_HESAB.hes"
    Caption ="قبض پذيرش"
    DatasheetFontName ="Arial (Arabic)"
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
        Begin PageBreak
            Width =283
        End
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =2
            Width =5103
            Height =3402
            FontName ="Arial (Arabic)"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="PIDT"
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =15330
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =53
                    Top =1130
                    Width =1971
                    Height =345
                    FontSize =10
                    FontWeight =700
                    Name ="PSERIAL"
                    ControlSource ="PSERIAL"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2085
                            Top =1130
                            Width =630
                            Height =345
                            FontSize =10
                            Name ="Label239"
                            Caption ="سريال:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8310
                    Top =2040
                    Width =1536
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="PKIND"
                    ControlSource ="PKIND"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9892
                            Top =2040
                            Width =870
                            Height =345
                            FontSize =10
                            Name ="Label242"
                            Caption ="نوع پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2865
                    Top =1080
                    Width =6471
                    Height =465
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="MODEL"
                    ControlSource ="MODEL"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9399
                            Top =1123
                            Width =1350
                            Height =345
                            FontSize =10
                            Name ="Label243"
                            Caption ="مشخصات دستگاه:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6195
                    Top =2040
                    Width =1071
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="GRENDDATE"
                    ControlSource ="GRENDDATE"
                    FontName ="Nazanin Mazar"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =7327
                            Top =2040
                            Width =945
                            Height =345
                            FontSize =10
                            Name ="Label244"
                            Caption =" اتمام گارانتي:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5400
                    Top =1560
                    Width =4281
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="POWNER"
                    ControlSource ="NAME"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9705
                            Top =1618
                            Width =1050
                            Height =345
                            FontSize =10
                            Name ="Label245"
                            Caption ="مالك دستگاه:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =60
                    Top =1590
                    Width =4326
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="POWNERNAM"
                    ControlSource ="POWNERNAM"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4432
                            Top =1590
                            Width =975
                            Height =345
                            FontSize =10
                            Name ="Label246"
                            Caption ="نماينده مالك:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2385
                    Top =2040
                    Width =2691
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    Name ="GRCOMPA"
                    ControlSource ="GRCOMPA"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5137
                            Top =2040
                            Width =1020
                            Height =345
                            FontSize =10
                            Name ="Label247"
                            Caption =" گارانتي كننده:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =60
                    Top =2040
                    Width =1206
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1342
                            Top =2040
                            Width =1065
                            Height =345
                            FontSize =10
                            Name ="Label251"
                            Caption ="كاربر پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7388
                    Top =2880
                    Width =3396
                    Height =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BorderColor =4210752
                    Name ="EYBEZHAR"
                    ControlSource ="EYBEZHAR"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =8269
                            Top =2539
                            Width =1620
                            Height =345
                            FontSize =10
                            Name ="Label253"
                            Caption ="عيب به اظهار مشتري:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4493
                    Top =2880
                    Width =2826
                    Height =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BorderColor =4210752
                    Name ="LAVAZEM"
                    ControlSource ="LAVAZEM"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5300
                            Top =2535
                            Width =1395
                            Height =345
                            FontSize =10
                            Name ="Label254"
                            Caption ="ليست لوازم همراه:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Top =2880
                    Width =4326
                    Height =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    BorderColor =4210752
                    Name ="IDIATEK"
                    ControlSource ="IDIATEK"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =2
                            Left =-15
                            Top =2490
                            Width =4335
                            Height =345
                            FontSize =10
                            Name ="Label255"
                            Caption ="نظر تكنسين(پيش  بيني حدود هزينه و زمان)"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =120
                    Top =4320
                    Width =10596
                    Height =555
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    FontName ="Nazanin Mazar"
                End
                Begin Rectangle
                    Top =1080
                    Width =10829
                    Height =1408
                    Name ="Box262"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =4095
                    Top =450
                    Width =1905
                    Height =465
                    FontSize =12
                    ForeColor =0
                    Name ="Label22"
                    Caption ="قبض پذيرش تعميرگاه"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1770
                    Width =6513
                    Height =540
                    FontSize =14
                    FontWeight =700
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8698
                    Top =680
                    Width =966
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    Name ="PIDT"
                    ControlSource ="PIDT"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9710
                            Top =680
                            Width =1035
                            Height =345
                            FontSize =10
                            Name ="Label238"
                            Caption ="شماره پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1425
                    Top =645
                    Width =981
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    Name ="PINDATE"
                    ControlSource ="PINDATE"
                    FontName ="Nazanin Mazar"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2452
                            Top =645
                            Width =1020
                            Height =345
                            FontSize =10
                            Name ="Label240"
                            Caption ="تاريخ پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =60
                    Top =645
                    Width =591
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    Name ="PINTIME"
                    ControlSource ="PINTIME"
                    FontName ="Nazanin Mazar"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =697
                            Top =645
                            Width =570
                            Height =345
                            FontSize =10
                            Name ="Label241"
                            Caption ="ساعت :"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3735
                    Top =5625
                    Width =2925
                    Height =465
                    FontSize =12
                    ForeColor =0
                    Name ="Label342"
                    Caption ="قبض پذيرش نسخه تعميرگاه"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1965
                    Top =5175
                    Width =6513
                    Height =540
                    FontSize =14
                    FontWeight =700
                    TabIndex =16
                    BorderColor =32768
                    Name ="Text343"
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8698
                    Top =5855
                    Width =966
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =17
                    Name ="Text344"
                    ControlSource ="PIDT"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9710
                            Top =5855
                            Width =1035
                            Height =345
                            FontSize =10
                            Name ="Label345"
                            Caption ="شماره پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1380
                    Top =5835
                    Width =1086
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =18
                    Name ="Text346"
                    ControlSource ="PINDATE"
                    FontName ="Nazanin Mazar"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2512
                            Top =5835
                            Width =1020
                            Height =345
                            FontSize =10
                            Name ="Label347"
                            Caption ="تاريخ پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =120
                    Top =5835
                    Width =591
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =19
                    Name ="Text348"
                    ControlSource ="PINTIME"
                    FontName ="Nazanin Mazar"
                    InputMask ="##:##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =757
                            Top =5835
                            Width =570
                            Height =345
                            FontSize =10
                            Name ="Label349"
                            Caption ="ساعت :"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =113
                    Top =6320
                    Width =1911
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =20
                    Name ="Text350"
                    ControlSource ="PSERIAL"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2085
                            Top =6320
                            Width =630
                            Height =345
                            FontSize =10
                            Name ="Label351"
                            Caption ="سريال:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8190
                    Top =7140
                    Width =1656
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =21
                    Name ="Text352"
                    ControlSource ="PKIND"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9892
                            Top =7140
                            Width =870
                            Height =345
                            FontSize =10
                            Name ="Label353"
                            Caption ="نوع پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2865
                    Top =6317
                    Width =6471
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =22
                    Name ="Text354"
                    ControlSource ="MODEL"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9399
                            Top =6315
                            Width =1350
                            Height =345
                            FontSize =10
                            Name ="Label355"
                            Caption ="مشخصات دستگاه:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5449
                    Top =7140
                    Width =1251
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =23
                    Name ="Text356"
                    ControlSource ="GRENDDATE"
                    FontName ="Nazanin Mazar"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6761
                            Top =7140
                            Width =1425
                            Height =345
                            FontSize =10
                            Name ="Label357"
                            Caption ="تاريخ اتمام گارانتي:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5460
                    Top =6752
                    Width =4221
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =24
                    Name ="Text358"
                    ControlSource ="NAME"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9705
                            Top =6750
                            Width =1050
                            Height =345
                            FontSize =10
                            Name ="Label359"
                            Caption ="مالك دستگاه:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =120
                    Top =6750
                    Width =4281
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =25
                    Name ="Text360"
                    ControlSource ="POWNERNAM"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4447
                            Top =6750
                            Width =975
                            Height =345
                            FontSize =10
                            Name ="Label361"
                            Caption ="نماينده مالك:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =124
                    Top =7140
                    Width =3681
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =26
                    Name ="Text362"
                    ControlSource ="GRCOMPA"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3866
                            Top =7140
                            Width =1530
                            Height =345
                            FontSize =10
                            Name ="Label363"
                            Caption ="شركت گارانتي كننده:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =5916
                    Top =7561
                    Width =1626
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =27
                    Name ="Text364"
                    ControlSource ="IYALAT"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =7543
                            Top =7561
                            Width =630
                            Height =345
                            FontSize =10
                            Name ="Label365"
                            Caption ="استان:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3664
                    Top =7560
                    Width =1401
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =28
                    Name ="Text366"
                    ControlSource ="CITY"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5111
                            Top =7560
                            Width =750
                            Height =345
                            FontSize =10
                            Name ="Label367"
                            Caption ="شهرستان:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =120
                    Top =7561
                    Width =1896
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =29
                    Name ="Text368"
                    ControlSource ="USER_NAME"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2092
                            Top =7561
                            Width =1065
                            Height =345
                            FontSize =10
                            Name ="Label369"
                            Caption ="كاربر پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =1710
                    Top =8625
                    Width =7431
                    Height =390
                    FontSize =10
                    FontWeight =700
                    TabIndex =30
                    BackColor =10092543
                    BorderColor =4210752
                    Name ="Text370"
                    ControlSource ="EYBEZHAR"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9128
                            Top =8622
                            Width =1620
                            Height =345
                            FontSize =10
                            Name ="Label371"
                            Caption ="عيب به اظهار مشتري:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =75
                    Top =8010
                    Width =9726
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =31
                    Name ="Text372"
                    ControlSource ="ADRESS"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9817
                            Top =8010
                            Width =945
                            Height =345
                            FontSize =10
                            Name ="Label373"
                            Caption ="آدرس و تلفن:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin Rectangle
                    Left =60
                    Top =6270
                    Width =10769
                    Height =2263
                    Name ="Box374"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8190
                    Top =7560
                    Width =1551
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TabIndex =32
                    Name ="Combo375"
                    ControlSource ="PINPKIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"در محل تعميرگاه\";2;\"در محل مشتري\""
                    ColumnWidths ="0"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9787
                            Top =7560
                            Width =975
                            Height =345
                            FontSize =10
                            Name ="Label376"
                            Caption ="محل پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =65
                    Top =9221
                    Width =10711
                    Height =4378
                    Name ="Box267"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1425
                    Top =9210
                    Width =0
                    Height =5695
                    BorderColor =0
                    Name ="Line271"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =60
                    Top =9615
                    Width =10738
                    BorderColor =0
                    Name ="Line272"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6915
                    Top =9240
                    Width =2055
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label273"
                    Caption ="شرح كارهاي انجام شده  "
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =390
                    Top =9277
                    Width =585
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="Label277"
                    Caption ="مبلغ"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2942
                    Top =9240
                    Width =1155
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label284"
                    Caption ="قطعات مصرفي"
                    FontName ="Titr"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =66
                    Top =13635
                    Width =10712
                    Height =1273
                    Name ="Box289"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1425
                    Top =13659
                    Width =2295
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label292"
                    Caption ="جمع اجرت :"
                    FontName ="Nazanin Mazar"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1425
                    Top =14056
                    Width =2295
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label293"
                    Caption ="جمع قطعات مصرفي :"
                    FontName ="Nazanin Mazar"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1425
                    Top =14453
                    Width =2295
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label294"
                    Caption ="جمع كل :"
                    FontName ="Nazanin Mazar"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2325
                    Top =9240
                    Width =0
                    Height =4360
                    BorderColor =0
                    Name ="Line378"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1605
                    Top =9285
                    Width =525
                    Height =345
                    FontSize =10
                    ForeColor =0
                    Name ="Label379"
                    Caption ="تعداد"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =60
                    Top =14970
                    Width =10665
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label382"
                    Caption ="فاكتور شماره                                           به مبلغ                  "
                        "                                 مورخ                                           "
                        "    صادرگرديد."
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =75
                    Top =14445
                    Width =3013
                    BorderColor =0
                    Name ="Line298"
                End
                Begin Line
                    BorderWidth =1
                    Left =75
                    Top =14055
                    Width =3013
                    BorderColor =0
                    Name ="Line299"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3075
                    Top =13620
                    Width =0
                    Height =1285
                    BorderColor =0
                    Name ="Line300"
                End
                Begin Label
                    TextFontFamily =2
                    Left =570
                    Top =8670
                    Width =1020
                    Height =345
                    FontSize =10
                    Name ="Label383"
                    Caption ="امضاء مشتري"
                    FontName ="Nazanin Mazar"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4650
                    Top =9240
                    Width =0
                    Height =4360
                    BorderColor =0
                    Name ="Line384"
                End
                Begin Label
                    TextFontFamily =2
                    Left =9240
                    Top =13770
                    Width =1500
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label385"
                    Caption ="نام و امضاء تكنيسين:"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    Left =9240
                    Top =14250
                    Width =1185
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label386"
                    Caption ="كاركرد دستگاه:"
                    FontName ="Titr"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="GroupFooter0"
        End
    End
End
CodeBehindForm
' See "PPAZIRESHTAMM.cls"
