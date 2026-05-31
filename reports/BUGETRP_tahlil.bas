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
    Width =15703
    DatasheetFontHeight =10
    ItemSuffix =249
    Left =1170
    Top =1500
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
            Height =480
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =0
                    Left =7035
                    Width =1935
                    Height =435
                    FontSize =16
                    Name ="Label0"
                    Caption =" انحراف بودجه"
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
            Height =9223
            Name ="Detail"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11295
                    Width =2841
                    Height =600
                    FontSize =18
                    Name ="BGMAH1"
                    ControlSource ="=MAHNAME([BGMAH])"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =14192
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
                    Left =353
                    Top =56
                    Height =600
                    FontSize =18
                    TabIndex =1
                    Name ="BGSAL"
                    ControlSource ="BGSAL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =2111
                            Top =56
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
                    Left =2550
                    Top =795
                    Width =2826
                    Height =630
                    FontSize =16
                    TabIndex =2
                    Name ="BGBUGETMON"
                    ControlSource ="BGBUGETMON"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =5355
                            Top =915
                            Width =1740
                            Height =420
                            FontSize =12
                            Name ="Label243"
                            Caption ="پيش بيني  وجوه نقد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7275
                    Top =2855
                    Width =1776
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
                    Left =7275
                    Top =3305
                    Width =1776
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
                    Left =7275
                    Top =3755
                    Width =1776
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
                    Left =7275
                    Top =4220
                    Width =1776
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
                    Left =7275
                    Top =4655
                    Width =1776
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
                    Left =7275
                    Top =5120
                    Width =1776
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
                    Left =7275
                    Top =5570
                    Width =1776
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
                    Left =7275
                    Top =6020
                    Width =1776
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
                    Left =7275
                    Top =6485
                    Width =1776
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
                    Left =7275
                    Top =6935
                    Width =1776
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
                    Left =7275
                    Top =7385
                    Width =1776
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
                    Left =7275
                    Top =7835
                    Width =1776
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
                    Left =7275
                    Top =8285
                    Width =1776
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
                    Left =11820
                    Top =2855
                    Width =1911
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
                    Left =11820
                    Top =3305
                    Width =1911
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
                    Left =11820
                    Top =3755
                    Width =1911
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
                    Left =11820
                    Top =4220
                    Width =1911
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
                    Left =11820
                    Top =4655
                    Width =1911
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
                    Left =11820
                    Top =5120
                    Width =1911
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
                    Left =11820
                    Top =5570
                    Width =1911
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
                    Left =11820
                    Top =6020
                    Width =1911
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
                    Left =11820
                    Top =6485
                    Width =1911
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
                    Left =11820
                    Top =6935
                    Width =1911
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
                    Left =11820
                    Top =7385
                    Width =1911
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
                    Left =11820
                    Top =7835
                    Width =1911
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
                    Left =11850
                    Top =8285
                    Width =1911
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
                    Left =2430
                    Top =2840
                    Width =2061
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
                    Left =2430
                    Top =3290
                    Width =2061
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
                    Left =2430
                    Top =3740
                    Width =2061
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
                    Left =2430
                    Top =4205
                    Width =2061
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
                    Left =2430
                    Top =4640
                    Width =2061
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
                    Left =2430
                    Top =5105
                    Width =2061
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
                    Left =2430
                    Top =5555
                    Width =2061
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
                    Left =2430
                    Top =6005
                    Width =2061
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
                    Left =2430
                    Top =6470
                    Width =2061
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
                    Left =2430
                    Top =6920
                    Width =2061
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
                    Left =2430
                    Top =7370
                    Width =2061
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
                    Left =2430
                    Top =7820
                    Width =2061
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
                    Left =2430
                    Top =8270
                    Width =2061
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
                    Left =9585
                    Top =755
                    Width =2106
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
                            Left =11949
                            Top =755
                            Width =3480
                            Height =345
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
                    Left =9592
                    Top =1204
                    Width =2106
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
                            Left =11954
                            Top =1208
                            Width =3600
                            Height =345
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
                    Left =9592
                    Top =1654
                    Width =1686
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
                            Left =11956
                            Top =1654
                            Width =2115
                            Height =345
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
                    Left =11790
                    Top =2390
                    Width =1875
                    Height =330
                    FontSize =10
                    Name ="Label62"
                    Caption ="پيش بيني فروش"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    Left =7215
                    Top =2390
                    Width =1860
                    Height =255
                    FontSize =7
                    Name ="Label63"
                    Caption ="پيش بيني مبلغ چکها با سررسيد بودحه"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =2415
                    Top =2385
                    Width =2085
                    Height =330
                    FontSize =10
                    Name ="Label64"
                    Caption ="پيش بيني جمع چکهاي دريافتي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontCharSet =178
                    TextAlign =1
                    Left =14310
                    Top =2840
                    Width =1230
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
                    Left =14310
                    Top =3290
                    Width =1230
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
                    Left =14310
                    Top =3740
                    Width =1230
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
                    Left =14310
                    Top =4190
                    Width =1230
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
                    Left =14310
                    Top =4640
                    Width =1230
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
                    Left =14310
                    Top =5090
                    Width =1230
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
                    Left =14310
                    Top =5540
                    Width =1230
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
                    Left =14310
                    Top =5990
                    Width =1230
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
                    Left =14310
                    Top =6440
                    Width =1230
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
                    Left =14310
                    Top =6890
                    Width =1230
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
                    Left =14310
                    Top =7340
                    Width =1230
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
                    Left =14310
                    Top =7790
                    Width =1230
                    Height =345
                    FontSize =10
                    Name ="M12"
                    Caption ="فروردين"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =2
                    Left =73
                    Top =639
                    Width =15530
                    Height =8558
                    Name ="Box86"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =14250
                    Top =2270
                    Width =0
                    Height =6928
                    Name ="Line87"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =7200
                    Top =635
                    Width =0
                    Height =8578
                    Name ="Line88"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =2385
                    Top =2265
                    Width =0
                    Height =6958
                    Name ="Line89"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =2780
                    Width =15553
                    Name ="Line92"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =14295
                    Top =2390
                    Width =1185
                    Height =330
                    FontSize =10
                    Name ="Label93"
                    Caption ="ماه"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =2270
                    Width =15553
                    Name ="Line94"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =11790
                    Top =635
                    Width =0
                    Height =8578
                    Name ="Line96"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =9525
                    Top =635
                    Width =0
                    Height =8578
                    Name ="Line98"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =9585
                    Top =2390
                    Width =2160
                    Height =330
                    FontSize =10
                    Name ="Label99"
                    Caption ="فروش ايجاد شده"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    Left =9555
                    Top =2812
                    Width =2211
                    Height =5900
                    BackColor =10092543
                    BorderColor =10092543
                    Name ="Box100"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =2840
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =45
                    Name ="BGFOCN0MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =3290
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =46
                    Name ="BGFOCN1MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =3740
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =47
                    Name ="BGFOCN2MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =4205
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =48
                    Name ="BGFOCN3MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =4640
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =49
                    Name ="BGFOCN4MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =5105
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =50
                    Name ="BGFOCN5MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =5555
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =51
                    Name ="BGFOCN6MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =6005
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =52
                    Name ="BGFOCN7MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =6470
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =53
                    Name ="BGFOCN8MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =6920
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =54
                    Name ="BGFOCN9MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =7370
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =55
                    Name ="BGFOCN10MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9600
                    Top =7820
                    Width =2121
                    Height =360
                    FontSize =10
                    TabIndex =56
                    Name ="BGFOCN11MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =9645
                    Top =8270
                    Width =2106
                    Height =360
                    FontSize =10
                    TabIndex =57
                    Name ="BGFOCN12MONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =2855
                    Width =441
                    Height =360
                    TabIndex =58
                    Name ="BGFOCN0MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =3305
                    Width =441
                    Height =360
                    TabIndex =59
                    Name ="BGFOCN1MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =3755
                    Width =441
                    Height =360
                    TabIndex =60
                    Name ="BGFOCN2MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =4220
                    Width =441
                    Height =360
                    TabIndex =61
                    Name ="BGFOCN3MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =4655
                    Width =441
                    Height =360
                    TabIndex =62
                    Name ="BGFOCN4MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =5120
                    Width =441
                    Height =360
                    TabIndex =63
                    Name ="BGFOCN5MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =5570
                    Width =441
                    Height =360
                    TabIndex =64
                    Name ="BGFOCN6MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =6020
                    Width =441
                    Height =360
                    TabIndex =65
                    Name ="BGFOCN7MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =6485
                    Width =441
                    Height =360
                    TabIndex =66
                    Name ="BGFOCN8MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =6935
                    Width =441
                    Height =360
                    TabIndex =67
                    Name ="BGFOCN9MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =7385
                    Width =441
                    Height =360
                    TabIndex =68
                    Name ="BGFOCN10MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13770
                    Top =7835
                    Width =441
                    Height =360
                    TabIndex =69
                    Name ="BGFOCN11MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =13800
                    Top =8285
                    Width =441
                    Height =360
                    TabIndex =70
                    Name ="BGFOCN12MOND"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =13770
                    Top =2270
                    Width =0
                    Height =6928
                    Name ="Line127"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =13800
                    Top =2390
                    Width =405
                    Height =330
                    FontSize =10
                    Name ="Label128"
                    Caption ="%"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =2853
                    Width =441
                    Height =360
                    TabIndex =71
                    Name ="BGPCH0"
                    ControlSource ="BGPCH0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =3303
                    Width =441
                    Height =360
                    TabIndex =72
                    Name ="BGPCH1"
                    ControlSource ="BGPCH1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =3753
                    Width =441
                    Height =360
                    TabIndex =73
                    Name ="BGPCH2"
                    ControlSource ="BGPCH2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =4218
                    Width =441
                    Height =360
                    TabIndex =74
                    Name ="BGPCH3"
                    ControlSource ="BGPCH3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =4653
                    Width =441
                    Height =360
                    TabIndex =75
                    Name ="BGPCH4"
                    ControlSource ="BGPCH4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =5118
                    Width =441
                    Height =360
                    TabIndex =76
                    Name ="BGPCH5"
                    ControlSource ="BGPCH5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =5568
                    Width =441
                    Height =360
                    TabIndex =77
                    Name ="BGPCH6"
                    ControlSource ="BGPCH6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =6018
                    Width =441
                    Height =360
                    TabIndex =78
                    Name ="BGPCH7"
                    ControlSource ="BGPCH7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =6483
                    Width =441
                    Height =360
                    TabIndex =79
                    Name ="BGPCH8"
                    ControlSource ="BGPCH8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =6933
                    Width =441
                    Height =360
                    TabIndex =80
                    Name ="BGPCH9"
                    ControlSource ="BGPCH9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =7383
                    Width =441
                    Height =360
                    TabIndex =81
                    Name ="BGPCH10"
                    ControlSource ="BGPCH10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9070
                    Top =7833
                    Width =441
                    Height =360
                    TabIndex =82
                    Name ="BGPCH11"
                    ControlSource ="BGPCH11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9100
                    Top =8283
                    Width =441
                    Height =360
                    TabIndex =83
                    Name ="BGPCH12"
                    ControlSource ="BGPCH12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =9100
                    Top =2388
                    Width =405
                    Height =330
                    FontSize =10
                    Name ="Label172"
                    Caption ="%"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =9075
                    Top =2270
                    Width =0
                    Height =6928
                    Name ="Line173"
                End
                Begin Rectangle
                    Left =4980
                    Top =2812
                    Width =2211
                    Height =5900
                    BackColor =10092543
                    BorderColor =10092543
                    Name ="Box174"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =4995
                    Top =2385
                    Width =1815
                    Height =330
                    Name ="Label175"
                    Caption ="مبلغ چکها با سررسيد بودحه"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =2838
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =84
                    Name ="BGMBCHEKR0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =3288
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =85
                    Name ="BGMBCHEKR1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =3738
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =86
                    Name ="BGMBCHEKR2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =4203
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =87
                    Name ="BGMBCHEKR3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =4638
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =88
                    Name ="BGMBCHEKR4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =5103
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =89
                    Name ="BGMBCHEKR5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =5553
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =90
                    Name ="BGMBCHEKR6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =6003
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =91
                    Name ="BGMBCHEKR7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =6468
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =92
                    Name ="BGMBCHEKR8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =6918
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =93
                    Name ="BGMBCHEKR9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =7368
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =94
                    Name ="BGMBCHEKR10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5060
                    Top =7818
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =95
                    Name ="BGMBCHEKR11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5045
                    Top =8268
                    Width =1671
                    Height =360
                    FontSize =10
                    TabIndex =96
                    Name ="BGMBCHEKR12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =4995
                    Top =2295
                    Width =0
                    Height =6928
                    Name ="Line189"
                End
                Begin Rectangle
                    Left =120
                    Top =2812
                    Width =2241
                    Height =5900
                    BackColor =10092543
                    BorderColor =10092543
                    Name ="Box190"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =2855
                    Width =441
                    Height =360
                    TabIndex =97
                    Name ="BGPCHM0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =3305
                    Width =441
                    Height =360
                    TabIndex =98
                    Name ="BGPCHM1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =3755
                    Width =441
                    Height =360
                    TabIndex =99
                    Name ="BGPCHM2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =4220
                    Width =441
                    Height =360
                    TabIndex =100
                    Name ="BGPCHM3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =4655
                    Width =441
                    Height =360
                    TabIndex =101
                    Name ="BGPCHM4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =5120
                    Width =441
                    Height =360
                    TabIndex =102
                    Name ="BGPCHM5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =5570
                    Width =441
                    Height =360
                    TabIndex =103
                    Name ="BGPCHM6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =6020
                    Width =441
                    Height =360
                    TabIndex =104
                    Name ="BGPCHM7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =6485
                    Width =441
                    Height =360
                    TabIndex =105
                    Name ="BGPCHM8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =6935
                    Width =441
                    Height =360
                    TabIndex =106
                    Name ="BGPCHM9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =7385
                    Width =441
                    Height =360
                    TabIndex =107
                    Name ="BGPCHM10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4515
                    Top =7835
                    Width =441
                    Height =360
                    TabIndex =108
                    Name ="BGPCHM11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4530
                    Top =8280
                    Width =441
                    Height =360
                    TabIndex =109
                    Name ="BGPCHM12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =4545
                    Top =2390
                    Width =405
                    Height =330
                    FontSize =10
                    Name ="Label204"
                    Caption ="%"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =4485
                    Top =2270
                    Width =0
                    Height =6928
                    Name ="Line205"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =2835
                    Height =360
                    FontSize =10
                    TabIndex =110
                    Name ="BGMBCHEKJMR0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =3285
                    Height =360
                    FontSize =10
                    TabIndex =111
                    Name ="BGMBCHEKJMR1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =3735
                    Height =360
                    FontSize =10
                    TabIndex =112
                    Name ="BGMBCHEKJMR2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =4200
                    Height =360
                    FontSize =10
                    TabIndex =113
                    Name ="BGMBCHEKJMR3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =4635
                    Height =360
                    FontSize =10
                    TabIndex =114
                    Name ="BGMBCHEKJMR4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =5100
                    Height =360
                    FontSize =10
                    TabIndex =115
                    Name ="BGMBCHEKJMR5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =5550
                    Height =360
                    FontSize =10
                    TabIndex =116
                    Name ="BGMBCHEKJMR6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =6000
                    Height =360
                    FontSize =10
                    TabIndex =117
                    Name ="BGMBCHEKJMR7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =6465
                    Height =360
                    FontSize =10
                    TabIndex =118
                    Name ="BGMBCHEKJMR8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =6915
                    Height =360
                    FontSize =10
                    TabIndex =119
                    Name ="BGMBCHEKJMR9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =7365
                    Height =360
                    FontSize =10
                    TabIndex =120
                    Name ="BGMBCHEKJMR10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =7815
                    Height =360
                    FontSize =10
                    TabIndex =121
                    Name ="BGMBCHEKJMR11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =165
                    Top =8265
                    Height =360
                    FontSize =10
                    TabIndex =122
                    Name ="BGMBCHEKJMR12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =135
                    Top =2385
                    Width =1785
                    Height =330
                    Name ="Label207"
                    Caption =" جمع چکهاي دريافتي در اين ماه"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =2835
                    Width =441
                    Height =360
                    TabIndex =123
                    BackColor =10092543
                    Name ="BGPCHD0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =3285
                    Width =441
                    Height =360
                    TabIndex =124
                    BackColor =10092543
                    Name ="BGPCHD1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =3735
                    Width =441
                    Height =360
                    TabIndex =125
                    BackColor =10092543
                    Name ="BGPCHD2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =4200
                    Width =441
                    Height =360
                    TabIndex =126
                    BackColor =10092543
                    Name ="BGPCHD3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =4635
                    Width =441
                    Height =360
                    TabIndex =127
                    BackColor =10092543
                    Name ="BGPCHD4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =5100
                    Width =441
                    Height =360
                    TabIndex =128
                    BackColor =10092543
                    Name ="BGPCHD5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =5550
                    Width =441
                    Height =360
                    TabIndex =129
                    BackColor =10092543
                    Name ="BGPCHD6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =6000
                    Width =441
                    Height =360
                    TabIndex =130
                    BackColor =10092543
                    Name ="BGPCHD7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =6465
                    Width =441
                    Height =360
                    TabIndex =131
                    BackColor =10092543
                    Name ="BGPCHD8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =6915
                    Width =441
                    Height =360
                    TabIndex =132
                    BackColor =10092543
                    Name ="BGPCHD9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =7365
                    Width =441
                    Height =360
                    TabIndex =133
                    BackColor =10092543
                    Name ="BGPCHD10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =7815
                    Width =441
                    Height =360
                    TabIndex =134
                    BackColor =10092543
                    Name ="BGPCHD11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6750
                    Top =8265
                    Width =441
                    Height =360
                    TabIndex =135
                    BackColor =10092543
                    Name ="BGPCHD12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =6780
                    Top =2370
                    Width =405
                    Height =330
                    FontSize =10
                    Name ="Label221"
                    Caption ="%"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =6750
                    Top =2265
                    Width =0
                    Height =6928
                    Name ="Line222"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =2835
                    Width =441
                    Height =360
                    TabIndex =136
                    BackColor =10092543
                    Name ="BGPCHR0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =3285
                    Width =441
                    Height =360
                    TabIndex =137
                    BackColor =10092543
                    Name ="BGPCHR1"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =3735
                    Width =441
                    Height =360
                    TabIndex =138
                    BackColor =10092543
                    Name ="BGPCHR2"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =4200
                    Width =441
                    Height =360
                    TabIndex =139
                    BackColor =10092543
                    Name ="BGPCHR3"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =4635
                    Width =441
                    Height =360
                    TabIndex =140
                    BackColor =10092543
                    Name ="BGPCHR4"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =5100
                    Width =441
                    Height =360
                    TabIndex =141
                    BackColor =10092543
                    Name ="BGPCHR5"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =5550
                    Width =441
                    Height =360
                    TabIndex =142
                    BackColor =10092543
                    Name ="BGPCHR6"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =6000
                    Width =441
                    Height =360
                    TabIndex =143
                    BackColor =10092543
                    Name ="BGPCHR7"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =6465
                    Width =441
                    Height =360
                    TabIndex =144
                    BackColor =10092543
                    Name ="BGPCHR8"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =6915
                    Width =441
                    Height =360
                    TabIndex =145
                    BackColor =10092543
                    Name ="BGPCHR9"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =7365
                    Width =441
                    Height =360
                    TabIndex =146
                    BackColor =10092543
                    Name ="BGPCHR10"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =7815
                    Width =441
                    Height =360
                    TabIndex =147
                    BackColor =10092543
                    Name ="BGPCHR11"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1920
                    Top =8265
                    Width =441
                    Height =360
                    TabIndex =148
                    BackColor =10092543
                    Name ="BGPCHR12"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =1957
                    Top =2381
                    Width =405
                    Height =330
                    FontSize =10
                    Name ="Label236"
                    Caption ="%"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =1920
                    Top =2265
                    Width =0
                    Height =6928
                    Name ="Line237"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7245
                    Top =1650
                    Width =1806
                    Height =360
                    FontSize =10
                    TabIndex =149
                    BackColor =10092543
                    Name ="BGMBCASHD"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =11295
                    Top =1650
                    Width =471
                    Height =360
                    FontSize =10
                    TabIndex =150
                    BorderColor =4210752
                    Name ="BGPCASH"
                    ControlSource ="BGPCASH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =1
                    OldBorderStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9030
                    Top =1650
                    Width =486
                    Height =360
                    FontSize =10
                    TabIndex =151
                    BackColor =10092543
                    BorderColor =4210752
                    Name ="BGPCASH0"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2535
                    Top =1545
                    Width =2826
                    Height =630
                    FontSize =16
                    TabIndex =152
                    BackColor =10092543
                    Name ="BGBUGETMONR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =5385
                            Top =1665
                            Width =1695
                            Height =420
                            FontSize =12
                            Name ="Label245"
                            Caption ="وجوه نقد ايجاد شده:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =510
                    Top =630
                    Width =1410
                    Height =540
                    FontSize =16
                    Name ="Label246"
                    Caption ="درصد تحقق "
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =396
                    Top =1200
                    Width =1656
                    Height =960
                    FontSize =40
                    TabIndex =153
                    BackColor =255
                    Name ="DDD"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "BUGETRP_tahlil.cls"
