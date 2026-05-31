Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10710
    DatasheetFontHeight =10
    ItemSuffix =96
    Left =270
    Top =1380
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x38dd77022fa3e540
    End
    RecordSource ="BUGET_MAIN"
    Caption ="بودجه"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="WeblogmaYekan"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
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
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin Subform
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin PageHeader
            Height =615
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =0
                    Left =4808
                    Width =795
                    Height =615
                    FontSize =16
                    Name ="Label0"
                    Caption ="بودجه"
                    FontName ="B Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Top =120
                    Width =1311
                    Height =360
                    FontSize =10
                    Name ="BGCDATE"
                    ControlSource ="BGCDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =1418
                            Top =120
                            Width =510
                            Height =360
                            FontSize =10
                            Name ="Label1"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =12193
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5430
                    Top =225
                    Width =2841
                    Height =600
                    FontSize =18
                    Name ="BGMAH1"
                    ControlSource ="=MAHNAME([BGMAH])"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =8327
                            Top =225
                            Width =1290
                            Height =600
                            FontSize =18
                            Name ="Label2"
                            Caption ="بودجه ماه:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    IMESentenceMode =3
                    Left =1695
                    Top =225
                    Height =600
                    FontSize =18
                    TabIndex =1
                    Name ="BGSAL"
                    ControlSource ="BGSAL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =3453
                            Top =225
                            Width =1065
                            Height =600
                            FontSize =18
                            Name ="Label3"
                            Caption ="در سال:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3285
                    Top =11415
                    Width =3081
                    Height =630
                    FontSize =16
                    TabIndex =2
                    Name ="BGBUGETMON"
                    ControlSource ="BGBUGETMON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =3585
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =3
                    Name ="BGMBCHEK0"
                    ControlSource ="BGMBCHEK0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =4035
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =4
                    Name ="BGMBCHEK1"
                    ControlSource ="BGMBCHEK1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =4485
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =5
                    Name ="BGMBCHEK2"
                    ControlSource ="BGMBCHEK2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =4950
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =6
                    Name ="BGMBCHEK3"
                    ControlSource ="BGMBCHEK3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =5385
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =7
                    Name ="BGMBCHEK4"
                    ControlSource ="BGMBCHEK4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =5850
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =8
                    Name ="BGMBCHEK5"
                    ControlSource ="BGMBCHEK5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =6300
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =9
                    Name ="BGMBCHEK6"
                    ControlSource ="BGMBCHEK6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =6750
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =10
                    Name ="BGMBCHEK7"
                    ControlSource ="BGMBCHEK7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =7215
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =11
                    Name ="BGMBCHEK8"
                    ControlSource ="BGMBCHEK8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =7665
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =12
                    Name ="BGMBCHEK9"
                    ControlSource ="BGMBCHEK9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =8115
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =13
                    Name ="BGMBCHEK10"
                    ControlSource ="BGMBCHEK10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =8565
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =14
                    Name ="BGMBCHEK11"
                    ControlSource ="BGMBCHEK11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3825
                    Top =9015
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =15
                    Name ="BGMBCHEK12"
                    ControlSource ="BGMBCHEK12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =3585
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =16
                    Name ="BGFOCXDAY"
                    ControlSource ="BGFOCXDAY"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =4035
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =17
                    Name ="BGFOCN1MON"
                    ControlSource ="BGFOCN1MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =4485
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =18
                    Name ="BGFOCN2MON"
                    ControlSource ="BGFOCN2MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =4950
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =19
                    Name ="BGFOCN3MON"
                    ControlSource ="BGFOCN3MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =5385
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =20
                    Name ="BGFOCN4MON"
                    ControlSource ="BGFOCN4MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =5850
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =21
                    Name ="BGFOCN5MON"
                    ControlSource ="BGFOCN5MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =6300
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =22
                    Name ="BGFOCN6MON"
                    ControlSource ="BGFOCN6MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =6750
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =23
                    Name ="BGFOCN7MON"
                    ControlSource ="BGFOCN7MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =7215
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =24
                    Name ="BGFOCN8MON"
                    ControlSource ="BGFOCN8MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =7665
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =25
                    Name ="BGFOCN9MON"
                    ControlSource ="BGFOCN9MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =8115
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =26
                    Name ="BGFOCN10MON"
                    ControlSource ="BGFOCN10MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =8565
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =27
                    Name ="BGFOCN11MON"
                    ControlSource ="BGFOCN11MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6465
                    Top =9015
                    Width =2331
                    Height =360
                    FontSize =10
                    TabIndex =28
                    Name ="Text48"
                    ControlSource ="BGFOCN11MON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =3585
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =29
                    Name ="BGMBCHEKJM0"
                    ControlSource ="BGMBCHEKJM0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =4035
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =30
                    Name ="BGMBCHEKJM1"
                    ControlSource ="BGMBCHEKJM1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =4485
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =31
                    Name ="BGMBCHEKJM2"
                    ControlSource ="BGMBCHEKJM2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =4950
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =32
                    Name ="BGMBCHEKJM3"
                    ControlSource ="BGMBCHEKJM3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =5385
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =33
                    Name ="BGMBCHEKJM4"
                    ControlSource ="BGMBCHEKJM4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =5850
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =34
                    Name ="BGMBCHEKJM5"
                    ControlSource ="BGMBCHEKJM5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =6300
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =35
                    Name ="BGMBCHEKJM6"
                    ControlSource ="BGMBCHEKJM6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =6750
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =36
                    Name ="BGMBCHEKJM7"
                    ControlSource ="BGMBCHEKJM7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =7215
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =37
                    Name ="BGMBCHEKJM8"
                    ControlSource ="BGMBCHEKJM8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =7665
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =38
                    Name ="BGMBCHEKJM9"
                    ControlSource ="BGMBCHEKJM9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =8115
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =39
                    Name ="BGMBCHEKJM10"
                    ControlSource ="BGMBCHEKJM10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =8565
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =40
                    Name ="BGMBCHEKJM11"
                    ControlSource ="BGMBCHEKJM11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =495
                    Top =9015
                    Width =2196
                    Height =360
                    FontSize =10
                    TabIndex =41
                    Name ="BGMBCHEKJM12"
                    ControlSource ="BGMBCHEKJM12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3810
                    Top =1485
                    Width =2256
                    Height =360
                    FontSize =10
                    TabIndex =42
                    Name ="BGMBCHEKMON"
                    ControlSource ="BGMBCHEKMON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =6519
                            Top =1485
                            Width =3735
                            Height =285
                            FontSize =10
                            Name ="Label5"
                            Caption ="جمع چکهاي موجود در بانک و صندوق در تاريخ بودجه"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3817
                    Top =1934
                    Width =2256
                    Height =360
                    FontSize =10
                    TabIndex =43
                    Name ="BGMBDARAM"
                    ControlSource ="BGMBDARAM"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =6524
                            Top =1938
                            Width =4065
                            Height =285
                            FontSize =10
                            Name ="Label6"
                            Caption ="درآمد هاي غير عملياتي که در ماه بودجه تزريق ميشود"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3817
                    Top =2384
                    Width =2256
                    Height =360
                    FontSize =10
                    TabIndex =44
                    Name ="BGMBCASH"
                    ControlSource ="BGMBCASH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =6526
                            Top =2384
                            Width =3030
                            Height =285
                            FontSize =10
                            Name ="Label7"
                            Caption ="مبلغ فروش نقدي در ماه بودجه"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =6435
                    Top =3120
                    Width =2340
                    Height =330
                    FontSize =10
                    Name ="Label62"
                    Caption ="پيش بيني فروش"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =3795
                    Top =3120
                    Width =2265
                    Height =330
                    FontSize =10
                    Name ="Label63"
                    Caption ="مبلغ چکها با سررسيد  ماه بودجه"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =450
                    Top =3120
                    Width =2265
                    Height =330
                    FontSize =10
                    Name ="Label64"
                    Caption ="جمع چکهاي دريافتي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =3570
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M1"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =4020
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M2"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =4470
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M3"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =4920
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M4"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =5370
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M5"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =5820
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M6"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =6270
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M7"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =6720
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M8"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =7170
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M9"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =7620
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M10"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =8070
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M11"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =8955
                    Top =8520
                    Width =1755
                    Height =345
                    FontSize =10
                    Name ="M12"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =2
                    Left =283
                    Top =1369
                    Width =10265
                    Height =10763
                    Name ="Box86"
                End
                Begin Line
                    BorderWidth =2
                    Left =8895
                    Top =3000
                    Width =0
                    Height =9133
                    Name ="Line87"
                End
                Begin Line
                    BorderWidth =2
                    Left =6405
                    Top =1350
                    Width =0
                    Height =10783
                    Name ="Line88"
                End
                Begin Line
                    BorderWidth =2
                    Left =3225
                    Top =1350
                    Width =0
                    Height =10783
                    Name ="Line89"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =270
                    Top =3525
                    Width =10288
                    Name ="Line92"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =8955
                    Top =3120
                    Width =1545
                    Height =330
                    FontSize =10
                    Name ="Label93"
                    Caption ="ماه"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =285
                    Top =3015
                    Width =10288
                    Name ="Line94"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =285
                    Top =11340
                    Width =10288
                    Name ="Line95"
                End
            End
        End
        Begin PageFooter
            Height =963
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "BUGETRP.cls"
