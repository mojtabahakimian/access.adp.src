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
    Width =11010
    DatasheetFontHeight =10
    ItemSuffix =264
    Left =270
    Top =1605
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x371ef67ef32ce340
    End
    RecordSource ="SELECT PTAMIRAT.PIDT, PTAMIRAT.PSERIAL, PTAMIRAT.PINDATE, PTAMIRAT.PINTIME, PKIN"
        "DCODING.PKIND, PTAMIRAT.MODEL, PTAMIRAT.GRENDDATE, CUST_HESAB.NAME, PTAMIRAT.POW"
        "NERNAM, PTAMIRAT.GRCOMPA, PTAMIRAT.IYALAT, PTAMIRAT.CITY, PTAMIRAT.PINPKIND, PTA"
        "MIRAT.USER_NAME, PTAMIRAT.PTAHVKIND, PTAMIRAT.EYBEZHAR, PTAMIRAT.LAVAZEM, PTAMIR"
        "AT.IDIATEK, PTAMIRAT.ADRESS, PTAMIRAT.PAYAM FROM PTAMIRAT INNER JOIN PKINDCODING"
        " ON PTAMIRAT.PKIND = PKINDCODING.PKINDCO INNER JOIN CUST_HESAB ON PTAMIRAT.POWNE"
        "R = CUST_HESAB.hes"
    Caption ="قبض پذيرش"
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
        Begin PageHeader
            Height =1025
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =4260
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
                    Left =1935
                    Width =6513
                    Height =540
                    FontSize =14
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8863
                    Top =680
                    Width =966
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="PIDT"
                    ControlSource ="PIDT"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9875
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
                    Left =1320
                    Top =645
                    Width =1086
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
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
                    TabIndex =3
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =4708
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =68
                    Top =50
                    Width =2061
                    Height =345
                    FontSize =10
                    FontWeight =700
                    Name ="PSERIAL"
                    ControlSource ="PSERIAL"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2190
                            Top =50
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
                    Left =8341
                    Top =560
                    Width =1656
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="PKIND"
                    ControlSource ="PKIND"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =10043
                            Top =560
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
                    Left =3007
                    Top =7
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
                            TextFontFamily =2
                            Left =9541
                            Top =50
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
                    Left =5568
                    Top =560
                    Width =1251
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
                            Left =6880
                            Top =560
                            Width =1425
                            Height =345
                            FontSize =10
                            Name ="Label244"
                            Caption ="تاريخ اتمام گارانتي:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =75
                    Top =512
                    Width =4371
                    Height =480
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    Name ="POWNER"
                    ControlSource ="NAME"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =4470
                            Top =570
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
                    Left =4483
                    Top =990
                    Width =5406
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
                            Left =9935
                            Top =990
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
                    Left =60
                    Top =990
                    Width =2796
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
                            Left =2917
                            Top =990
                            Width =1530
                            Height =345
                            FontSize =10
                            Name ="Label247"
                            Caption ="شركت گارانتي كننده:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6017
                    Top =1411
                    Width =1626
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =7644
                            Top =1411
                            Width =630
                            Height =345
                            FontSize =10
                            Name ="Label248"
                            Caption ="استان:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3765
                    Top =1410
                    Width =1401
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    Name ="CITY"
                    ControlSource ="CITY"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5212
                            Top =1410
                            Width =750
                            Height =345
                            FontSize =10
                            Name ="Label249"
                            Caption ="شهرستان:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =11
                    Top =1411
                    Width =2106
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2193
                            Top =1411
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
                    Left =7443
                    Top =2715
                    Width =3396
                    Height =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =11
                    BorderColor =4210752
                    Name ="EYBEZHAR"
                    ControlSource ="EYBEZHAR"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =8324
                            Top =2374
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
                    Left =4548
                    Top =2715
                    Width =2826
                    Height =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =12
                    BorderColor =4210752
                    Name ="LAVAZEM"
                    ControlSource ="LAVAZEM"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5355
                            Top =2370
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
                    Left =10
                    Top =2715
                    Width =4371
                    Height =1380
                    FontSize =10
                    FontWeight =700
                    TabIndex =13
                    BorderColor =4210752
                    Name ="IDIATEK"
                    ControlSource ="IDIATEK"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =2
                            Left =15
                            Top =2325
                            Width =4365
                            Height =345
                            FontSize =10
                            Name ="Label255"
                            Caption ="نظر تكنسين(پيش پيش بيني حدود هزينه و زمان)"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =76
                    Top =1864
                    Width =9876
                    Height =345
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    Name ="ADRESS"
                    ControlSource ="ADRESS"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9968
                            Top =1864
                            Width =945
                            Height =345
                            FontSize =10
                            Name ="Label256"
                            Caption ="آدرس و تلفن:"
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
                    Left =56
                    Top =4138
                    Width =10821
                    Height =570
                    FontSize =10
                    FontWeight =700
                    TabIndex =15
                    Name ="PAYAM"
                    ControlSource ="PAYAM"
                    FontName ="Nazanin Mazar"
                End
                Begin Rectangle
                    Width =10934
                    Height =2263
                    Name ="Box262"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =8341
                    Top =1411
                    Width =1551
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    Name ="PINPKIND"
                    ControlSource ="PINPKIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"در محل تعميرگاه\";2;\"در محل مشتري\""
                    ColumnWidths ="0"
                    FontName ="Nazanin Mazar"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =9938
                            Top =1411
                            Width =975
                            Height =345
                            FontSize =10
                            Name ="Label250"
                            Caption ="محل پذيرش:"
                            FontName ="Nazanin Mazar"
                        End
                    End
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =170
            Name ="GroupFooter0"
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "PPAZIRESHTAM.cls"
