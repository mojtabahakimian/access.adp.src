Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7662
    DatasheetFontHeight =10
    ItemSuffix =334
    Left =2895
    Top =1080
    Right =10650
    Bottom =7185
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x173e70f2aa14e340
    End
    RecordSource ="SIGN"
    Caption ="SIGN"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            Height =24512
            BackColor =16763904
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2834
                    Width =336
                    Name ="USERCO"
                    ControlSource ="USERCO"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5963
                    Top =570
                    TabIndex =1
                    Name ="FFR_FROOSH"
                    ControlSource ="FFR_FROOSH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6417
                            Top =510
                            Width =495
                            Height =240
                            Name ="Label3"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5966
                    Top =971
                    TabIndex =2
                    Name ="FFR_HESAB"
                    ControlSource ="FFR_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6400
                            Top =915
                            Width =795
                            Height =240
                            Name ="Label5"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5966
                    Top =1350
                    TabIndex =3
                    Name ="FFR_MODIR"
                    ControlSource ="FFR_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6400
                            Top =1350
                            Width =780
                            Height =240
                            Name ="Label7"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =570
                    TabIndex =4
                    Name ="KFR_BAZAR"
                    ControlSource ="KFR_BAZAR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2485
                            Top =570
                            Width =615
                            Height =240
                            Name ="Label9"
                            Caption ="بازرگاني"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =975
                    TabIndex =5
                    Name ="KFR_HESAB"
                    ControlSource ="KFR_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2482
                            Top =975
                            Width =795
                            Height =240
                            Name ="Label11"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =1365
                    TabIndex =6
                    Name ="KFR_MODIR"
                    ControlSource ="KFR_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2482
                            Top =1365
                            Width =780
                            Height =240
                            Name ="Label13"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5966
                    Top =2265
                    TabIndex =7
                    Name ="ANB_RASID_ANB"
                    ControlSource ="ANB_RASID_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6366
                            Top =2205
                            Width =555
                            Height =300
                            Name ="Label15"
                            Caption ="انبار دار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =2265
                    TabIndex =9
                    Name ="ANB_HAVL_FROSH"
                    ControlSource ="ANB_HAVL_FROSH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2478
                            Top =2265
                            Width =495
                            Height =240
                            Name ="Label17"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =2655
                    TabIndex =10
                    Name ="ANB_HAVL_ANB"
                    ControlSource ="ANB_HAVL_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2483
                            Top =2655
                            Width =510
                            Height =240
                            Name ="Label19"
                            Caption ="انباردار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =3060
                    TabIndex =11
                    Name ="ANB_HAVL_MODIR"
                    ControlSource ="ANB_HAVL_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2483
                            Top =3060
                            Width =780
                            Height =240
                            Name ="Label21"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =3975
                    TabIndex =14
                    Name ="ANB_TOLID_MODIRT"
                    ControlSource ="ANB_TOLID_MODIRT"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2464
                            Top =3915
                            Width =735
                            Height =240
                            Name ="Label23"
                            Caption ="مدير توليد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =4380
                    TabIndex =15
                    Name ="ANB_TOLID_QC"
                    ControlSource ="ANB_TOLID_QC"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2464
                            Top =4320
                            Width =930
                            Height =240
                            Name ="Label25"
                            Caption ="كنترل كيفيت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =4815
                    TabIndex =16
                    Name ="ANB_TOLID_ANB"
                    ControlSource ="ANB_TOLID_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2464
                            Top =4755
                            Width =330
                            Height =240
                            Name ="Label27"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6015
                    Top =4365
                    TabIndex =12
                    Name ="ANB_KHOROG_ANB"
                    ControlSource ="ANB_KHOROG_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6359
                            Top =4365
                            Width =510
                            Height =240
                            Name ="Label29"
                            Caption ="انباردار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6015
                    Top =3975
                    TabIndex =13
                    Name ="ANB_KHOROG_MODIRT"
                    ControlSource ="ANB_KHOROG_MODIRT"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6355
                            Top =3975
                            Width =735
                            Height =240
                            Name ="Label31"
                            Caption ="مدير توليد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6014
                    Top =6070
                    TabIndex =17
                    Name ="ANB_KHOROGS_ANB"
                    ControlSource ="ANB_KHOROGS_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6354
                            Top =6070
                            Width =555
                            Height =240
                            Name ="Label33"
                            Caption ="انبار دار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6014
                    Top =5666
                    TabIndex =18
                    Name ="ANB_KHOROGS_MODIRT"
                    ControlSource ="ANB_KHOROGS_MODIRT"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6354
                            Top =5666
                            Width =780
                            Height =240
                            Name ="Label35"
                            Caption ="مدير واحد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =5610
                    TabIndex =19
                    Name ="SND_TAHI"
                    ControlSource ="SND_TAHI"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2469
                            Top =5611
                            Width =885
                            Height =240
                            Name ="Label37"
                            Caption ="تنظيم كننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =6015
                    TabIndex =20
                    Name ="SND_MALI"
                    ControlSource ="SND_MALI"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2463
                            Top =6015
                            Width =825
                            Height =240
                            Name ="Label39"
                            Caption ="مدير مالي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2145
                    Top =6435
                    TabIndex =21
                    Name ="SND_MODIR"
                    ControlSource ="SND_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2463
                            Top =6435
                            Width =975
                            Height =240
                            Name ="Label41"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =339
                    Width =3560
                    Height =1399
                    Name ="Box42"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =165
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label43"
                    Caption ="فاكتور فروش"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =339
                    Width =3560
                    Height =1399
                    Name ="Box44"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =165
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label45"
                    Caption ="فاكتور خريد"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =2002
                    Width =3575
                    Height =1399
                    Name ="Box46"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =1828
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label47"
                    Caption ="رسيد انبار"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =1989
                    Width =3560
                    Height =1399
                    Name ="Box48"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =1815
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label49"
                    Caption ="حواله انبار فروش"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =3697
                    Width =3560
                    Height =1399
                    Name ="Box50"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =3525
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label51"
                    Caption ="رسيد توليد"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =3691
                    Width =3560
                    Height =1399
                    Name ="Box52"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =3525
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label53"
                    Caption ="حواله خروج مواد"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =5391
                    Width =3560
                    Height =1399
                    Name ="Box54"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =5220
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label55"
                    Caption ="حواله خروج ساير"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =5385
                    Width =3560
                    Height =1399
                    Name ="Box56"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =5220
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label57"
                    Caption ="سند حسابداري"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =5966
                    Top =2670
                    TabIndex =8
                    Name ="ANB_RASID_QC"
                    ControlSource ="ANB_RASID_QC"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6366
                            Top =2610
                            Width =930
                            Height =240
                            Name ="Label59"
                            Caption ="كنترل كيفيت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6026
                    Top =7377
                    TabIndex =22
                    Name ="FFRB_FROOSH"
                    ControlSource ="FFRB_FROOSH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6467
                            Top =7320
                            Width =495
                            Height =240
                            Name ="Label61"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6026
                    Top =8156
                    TabIndex =24
                    Name ="FFRB_HESAB"
                    ControlSource ="FFRB_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6460
                            Top =8100
                            Width =795
                            Height =240
                            Name ="Label63"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    Visible = NotDefault
                    OverlapFlags =93
                    Left =6026
                    Top =8565
                    TabIndex =25
                    Name ="FFRB_MODIR"
                    ControlSource ="FFRB_MODIR"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =93
                            Left =6460
                            Top =8565
                            Width =780
                            Height =240
                            Name ="Label65"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2092
                    Top =7320
                    TabIndex =26
                    Name ="KFRB_BAZAR"
                    ControlSource ="KFRB_BAZAR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2432
                            Top =7320
                            Width =615
                            Height =240
                            Name ="Label67"
                            Caption ="بازرگاني"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2085
                    Top =8100
                    TabIndex =28
                    Name ="KFRB_HESAB"
                    ControlSource ="KFRB_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2422
                            Top =8100
                            Width =795
                            Height =240
                            Name ="Label69"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2085
                    Top =8505
                    TabIndex =29
                    Name ="KFRB_MODIR"
                    ControlSource ="KFRB_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2422
                            Top =8505
                            Width =780
                            Height =240
                            Name ="Label71"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =7044
                    Width =3575
                    Height =1804
                    Name ="Box72"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =6870
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label73"
                    Caption =" برگشت فروش"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =7044
                    Width =3560
                    Height =1804
                    Name ="Box74"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =6870
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label75"
                    Caption =" برگشت خريد"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =6026
                    Top =7767
                    TabIndex =23
                    Name ="FFRB_ANB"
                    ControlSource ="FFRB_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6467
                            Top =7710
                            Width =495
                            Height =240
                            Name ="Label77"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2085
                    Top =7707
                    TabIndex =27
                    Name ="KFRB_ANB"
                    ControlSource ="KFRB_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2526
                            Top =7650
                            Width =495
                            Height =240
                            Name ="Label81"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6026
                    Top =9417
                    TabIndex =30
                    Name ="FFRP_FROOSH"
                    ControlSource ="FFRP_FROOSH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6467
                            Top =9360
                            Width =495
                            Height =240
                            Name ="Label83"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6026
                    Top =10151
                    TabIndex =32
                    Name ="FFRP_HESAB"
                    ControlSource ="FFRP_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6460
                            Top =10095
                            Width =795
                            Height =240
                            Name ="Label85"
                            Caption ="مديرعامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2092
                    Top =9465
                    TabIndex =33
                    Name ="RASM_HESAB"
                    ControlSource ="RASM_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2430
                            Top =9472
                            Width =795
                            Height =240
                            Name ="Label89"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =9127
                    Width =3575
                    Height =1399
                    Name ="Box94"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =8955
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label95"
                    Caption ="پيش فاكتور"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =9127
                    Width =3560
                    Height =1399
                    Name ="Box96"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =8955
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label97"
                    Caption ="رسيد مستقيم"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =6026
                    Top =9760
                    TabIndex =31
                    Name ="FFRP_ANB"
                    ControlSource ="FFRP_ANB"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6461
                            Top =9705
                            Width =795
                            Height =240
                            Name ="Label99"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =2085
                    Top =9865
                    TabIndex =34
                    Name ="RASM_MODIR"
                    ControlSource ="RASM_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2430
                            Top =9810
                            Width =735
                            Height =240
                            Name ="Label101"
                            Caption ="مديرعامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6041
                    Top =11109
                    TabIndex =35
                    Name ="PAY_MVAHED"
                    ControlSource ="PAY_MVAHED"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6476
                            Top =11055
                            Width =735
                            Height =240
                            Name ="Label103"
                            Caption ="مدير واحد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6026
                    Top =11486
                    TabIndex =36
                    Name ="PAY_HESAB"
                    ControlSource ="PAY_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6460
                            Top =11430
                            Width =795
                            Height =240
                            Name ="Label105"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =6026
                    Top =11850
                    TabIndex =37
                    Name ="PAY_MAMEL"
                    ControlSource ="PAY_MAMEL"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6460
                            Top =11790
                            Width =780
                            Height =300
                            Name ="Label107"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =10818
                    Width =3560
                    Height =1399
                    Name ="Box108"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =10650
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label109"
                    Caption ="درخواست پرداخت"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2035
                    Top =11110
                    TabIndex =38
                    Name ="ENTGH_AZANB"
                    ControlSource ="ENTGH_AZANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2380
                            Top =11056
                            Width =1125
                            Height =240
                            Name ="Label113"
                            Caption ="انبار فرستنده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2035
                    Top =11502
                    TabIndex =39
                    Name ="ENTGH_BEANB"
                    ControlSource ="ENTGH_BEANB"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2379
                            Top =11446
                            Width =1185
                            Height =240
                            Name ="Label115"
                            Caption ="انبار گيرنده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2035
                    Top =11851
                    TabIndex =40
                    Name ="ENTGH_MODIR"
                    ControlSource ="ENTGH_MODIR"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2379
                            Top =11851
                            Width =1170
                            Height =240
                            Name ="Label117"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =10818
                    Width =3560
                    Height =1399
                    Name ="Box118"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =10650
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label119"
                    Caption ="انتقال کالا از انبار به انبار"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2055
                    Top =12744
                    TabIndex =41
                    Name ="PAZ_PAZ"
                    ControlSource ="PAZ_PAZ"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2490
                            Top =12690
                            Width =735
                            Height =240
                            Name ="Label121"
                            Caption ="پذيرش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =13151
                    TabIndex =42
                    Name ="PAZ_HES"
                    ControlSource ="PAZ_HES"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2474
                            Top =13095
                            Width =795
                            Height =240
                            Name ="Label123"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =13485
                    TabIndex =43
                    Name ="PAZ_CEO"
                    ControlSource ="PAZ_CEO"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2474
                            Top =13485
                            Width =780
                            Height =240
                            Name ="Label125"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =12462
                    Width =3560
                    Height =1399
                    Name ="Box126"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =12300
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label127"
                    Caption ="پذيرش تعميرگاه"
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3930
                    Top =12462
                    Width =3560
                    Height =1399
                    Name ="Box134"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =5656
                    Top =12300
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label135"
                    Caption ="ساير رسيد انبارها"
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6026
                    Top =12687
                    TabIndex =44
                    Name ="SGN0124"
                    ControlSource ="SGN0124"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6287
                            Top =12630
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN1"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6026
                    Top =13151
                    TabIndex =45
                    Name ="SGN0224"
                    ControlSource ="SGN0224"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6280
                            Top =13095
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN2"
                            Caption ="كنترل كيفيت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6026
                    Top =13548
                    TabIndex =46
                    Name ="SGN0324"
                    ControlSource ="SGN0324"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6280
                            Top =13485
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =93
                    Left =3930
                    Top =14115
                    Width =3560
                    Height =1399
                    Name ="Box136"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =5656
                    Top =13950
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label137"
                    Caption ="ساير حواله انبارها"
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6026
                    Top =14397
                    TabIndex =47
                    Name ="SGN0126"
                    ControlSource ="SGN0126"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6287
                            Top =14340
                            Width =675
                            Height =240
                            ForeColor =16711680
                            Name ="Label139"
                            Caption ="انبار"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6026
                    Top =14801
                    TabIndex =48
                    Name ="SGN0226"
                    ControlSource ="SGN0226"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6280
                            Top =14745
                            Width =975
                            Height =240
                            ForeColor =16711680
                            Name ="Label141"
                            Caption ="كنترل كيفيت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =215
                    Left =6026
                    Top =15153
                    TabIndex =49
                    Name ="SGN0326"
                    ControlSource ="SGN0326"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6280
                            Top =15090
                            Width =960
                            Height =240
                            ForeColor =16711680
                            Name ="Label143"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5828
                    Top =16055
                    TabIndex =50
                    Name ="SGN0133"
                    ControlSource ="SGN0133"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6269
                            Top =15998
                            Width =495
                            Height =240
                            Name ="Label145"
                            Caption ="فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5846
                    Top =16436
                    TabIndex =51
                    Name ="SGN0233"
                    ControlSource ="SGN0233"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6280
                            Top =16380
                            Width =795
                            Height =240
                            Name ="Label147"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5846
                    Top =16785
                    TabIndex =52
                    Name ="SGN0333"
                    ControlSource ="SGN0333"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6280
                            Top =16785
                            Width =780
                            Height =240
                            Name ="Label149"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =15990
                    TabIndex =53
                    Name ="SGN0132"
                    ControlSource ="SGN0132"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2380
                            Top =15990
                            Width =615
                            Height =240
                            Name ="Label151"
                            Caption ="بازرگاني"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =16440
                    TabIndex =54
                    Name ="SGN0232"
                    ControlSource ="SGN0232"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =16440
                            Width =795
                            Height =240
                            Name ="Label153"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =16830
                    TabIndex =55
                    Name ="SGN0332"
                    ControlSource ="SGN0332"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =16830
                            Width =780
                            Height =240
                            Name ="Label155"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =15759
                    Width =3560
                    Height =1399
                    Name ="Box156"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =15585
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label157"
                    Caption ="فاكتور فروش صادرات"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =15759
                    Width =3560
                    Height =1399
                    Name ="Box158"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =15585
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label159"
                    Caption ="فاكتور خريد صادرات"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =18195
                    TabIndex =56
                    Name ="SGN0134"
                    ControlSource ="SGN0134"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2380
                            Top =18195
                            Width =885
                            Height =240
                            Name ="Label161"
                            Caption ="تنظيم كننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =18555
                    TabIndex =57
                    Name ="SGN0234"
                    ControlSource ="SGN0234"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =18555
                            Width =765
                            Height =240
                            Name ="Label163"
                            Caption ="مدير مالي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =18930
                    TabIndex =58
                    Name ="SGN0334"
                    ControlSource ="SGN0334"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =18930
                            Width =780
                            Height =240
                            Name ="Label165"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =17887
                    Width =3560
                    Height =1399
                    Name ="Box166"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =17715
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label167"
                    Caption ="خزانه داري"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =14460
                    TabIndex =59
                    Name ="SGN0135"
                    ControlSource ="SGN0135"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2380
                            Top =14460
                            Width =885
                            Height =240
                            Name ="Label169"
                            Caption ="تنظيم كننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =14805
                    TabIndex =60
                    Name ="SGN0235"
                    ControlSource ="SGN0235"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =14805
                            Width =765
                            Height =240
                            Name ="Label171"
                            Caption ="مدير مالي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =15195
                    TabIndex =61
                    Name ="SGN0335"
                    ControlSource ="SGN0335"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =15195
                            Width =780
                            Height =240
                            Name ="Label173"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =14120
                    Width =3560
                    Height =1399
                    Name ="Box174"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1860
                    Top =13950
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label175"
                    Caption ="سفارش"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5846
                    Top =18135
                    TabIndex =62
                    Name ="SGN0136"
                    ControlSource ="SGN0136"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6184
                            Top =18142
                            Width =1155
                            Height =240
                            Name ="Label177"
                            Caption ="درخواست كننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5846
                    Top =18585
                    TabIndex =63
                    Name ="SGN0236"
                    ControlSource ="SGN0236"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6183
                            Top =18585
                            Width =765
                            Height =240
                            Name ="Label179"
                            Caption ="تاييد کننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5846
                    Top =18930
                    TabIndex =64
                    Name ="SGN0336"
                    ControlSource ="SGN0336"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6183
                            Top =18930
                            Width =780
                            Height =240
                            Name ="Label181"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3930
                    Top =17917
                    Width =3560
                    Height =1399
                    Name ="Box182"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5656
                    Top =17745
                    Width =1755
                    Height =240
                    BackColor =13434828
                    Name ="Label183"
                    Caption ="درخواست خريد"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =510
                    Width =1648
                    Height =284
                    TabIndex =65
                    Name ="KFR_BAZARTX"
                    ControlSource ="KFR_BAZARTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =170
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label187"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =914
                    Width =1648
                    Height =284
                    TabIndex =66
                    Name ="KFR_HESABTX"
                    ControlSource ="KFR_HESABTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =1318
                    Width =1648
                    Height =284
                    TabIndex =67
                    Name ="KFR_MODIRTX"
                    ControlSource ="KFR_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =510
                    Width =1648
                    Height =284
                    TabIndex =68
                    Name ="FFR_FROOSHTX"
                    ControlSource ="FFR_FROOSHTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4274
                    Top =225
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label197"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =914
                    Width =1648
                    Height =284
                    TabIndex =69
                    Name ="FFR_HESABTX"
                    ControlSource ="FFR_HESABTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =1318
                    Width =1648
                    Height =284
                    TabIndex =70
                    Name ="FFR_MODIRTX"
                    ControlSource ="FFR_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =2200
                    Width =1648
                    Height =284
                    TabIndex =71
                    Name ="ANB_HAVL_FROSHTX"
                    ControlSource ="ANB_HAVL_FROSHTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =1860
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label201"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =2604
                    Width =1648
                    Height =284
                    TabIndex =72
                    Name ="ANB_HAVL_ANBTX"
                    ControlSource ="ANB_HAVL_ANBTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =3008
                    Width =1648
                    Height =284
                    TabIndex =73
                    Name ="ANB_HAVL_MODIRTX"
                    ControlSource ="ANB_HAVL_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =2200
                    Width =1648
                    Height =284
                    TabIndex =74
                    Name ="ANB_RASID_ANBTX"
                    ControlSource ="ANB_RASID_ANBTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4274
                    Top =1860
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label205"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =2604
                    Width =1648
                    Height =284
                    TabIndex =75
                    Name ="ANB_RASID_QCTX"
                    ControlSource ="ANB_RASID_QCTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =3911
                    Width =1648
                    Height =284
                    TabIndex =76
                    Name ="ANB_TOLID_MODIRTTX"
                    ControlSource ="ANB_TOLID_MODIRTTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =454
                    Top =3571
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label209"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =4315
                    Width =1648
                    Height =284
                    TabIndex =77
                    Name ="ANB_TOLID_QCTX"
                    ControlSource ="ANB_TOLID_QCTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =4719
                    Width =1648
                    Height =284
                    TabIndex =78
                    Name ="ANB_TOLID_ANBTX"
                    ControlSource ="ANB_TOLID_ANBTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4025
                    Top =4308
                    Width =1648
                    Height =284
                    TabIndex =79
                    Name ="ANB_KHOROG_ANBTX"
                    ControlSource ="ANB_KHOROG_ANBTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4264
                    Top =3571
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label216"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4035
                    Top =3915
                    Width =1648
                    Height =284
                    TabIndex =80
                    Name ="ANB_KHOROG_MODIRTTX"
                    ControlSource ="ANB_KHOROG_MODIRTTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =5560
                    Width =1648
                    Height =284
                    TabIndex =81
                    Name ="SND_TAHITX"
                    ControlSource ="SND_TAHITX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =5220
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label220"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =5964
                    Width =1648
                    Height =284
                    TabIndex =82
                    Name ="SND_MALITX"
                    ControlSource ="SND_MALITX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =6368
                    Width =1648
                    Height =284
                    TabIndex =83
                    Name ="SND_MODIRTX"
                    ControlSource ="SND_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4035
                    Top =6015
                    Width =1648
                    Height =284
                    TabIndex =84
                    Name ="ANB_KHOROGS_ANBTX"
                    ControlSource ="ANB_KHOROGS_ANBTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4264
                    Top =5215
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label224"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4035
                    Top =5610
                    Width =1648
                    Height =284
                    TabIndex =85
                    Name ="ANB_KHOROGS_MODIRTTX"
                    ControlSource ="ANB_KHOROGS_MODIRTTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =7256
                    Width =1648
                    Height =284
                    TabIndex =86
                    Name ="KFRB_BAZARTX"
                    ControlSource ="KFRB_BAZARTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =454
                    Top =6916
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label228"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =7660
                    Width =1648
                    Height =284
                    TabIndex =87
                    Name ="KFRB_ANBTX"
                    ControlSource ="KFRB_ANBTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =8064
                    Width =1648
                    Height =284
                    TabIndex =88
                    Name ="KFRB_HESABTX"
                    ControlSource ="KFRB_HESABTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =8468
                    Width =1648
                    Height =284
                    TabIndex =89
                    Name ="KFRB_MODIRTX"
                    ControlSource ="KFRB_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =7270
                    Width =1648
                    Height =284
                    TabIndex =90
                    Name ="FFRB_FROOSHTX"
                    ControlSource ="FFRB_FROOSHTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4259
                    Top =6930
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label233"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =7674
                    Width =1648
                    Height =284
                    TabIndex =91
                    Name ="FFRB_ANBTX"
                    ControlSource ="FFRB_ANBTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =8078
                    Width =1648
                    Height =284
                    TabIndex =92
                    Name ="FFRB_HESABTX"
                    ControlSource ="FFRB_HESABTX"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =8482
                    Width =1648
                    Height =284
                    TabIndex =93
                    Name ="FFRB_MODIRTX"
                    ControlSource ="FFRB_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =9355
                    Width =1648
                    Height =284
                    TabIndex =94
                    Name ="RASM_HESABTX"
                    ControlSource ="RASM_HESABTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =9015
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label238"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =9759
                    Width =1648
                    Height =284
                    TabIndex =95
                    Name ="RASM_MODIRTX"
                    ControlSource ="RASM_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =9295
                    Width =1648
                    Height =284
                    TabIndex =96
                    Name ="FFRP_FROOSHTX"
                    ControlSource ="FFRP_FROOSHTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4274
                    Top =8955
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label242"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =9699
                    Width =1648
                    Height =284
                    TabIndex =97
                    Name ="FFRP_ANBTX"
                    ControlSource ="FFRP_ANBTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =10103
                    Width =1648
                    Height =284
                    TabIndex =98
                    Name ="FFRP_HESABTX"
                    ControlSource ="FFRP_HESABTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =10995
                    Width =1648
                    Height =284
                    TabIndex =99
                    Name ="ENTGH_AZANBTX"
                    ControlSource ="ENTGH_AZANBTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =398
                    Top =10714
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label246"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =11399
                    Width =1648
                    Height =284
                    TabIndex =100
                    Name ="ENTGH_BEANBTX"
                    ControlSource ="ENTGH_BEANBTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =11803
                    Width =1648
                    Height =284
                    TabIndex =101
                    Name ="ENTGH_MODIRTX"
                    ControlSource ="ENTGH_MODIRTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =10990
                    Width =1648
                    Height =284
                    TabIndex =102
                    Name ="PAY_MVAHEDTX"
                    ControlSource ="PAY_MVAHEDTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4274
                    Top =10650
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label251"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =11394
                    Width =1648
                    Height =284
                    TabIndex =103
                    Name ="PAY_HESABTX"
                    ControlSource ="PAY_HESABTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =11798
                    Width =1648
                    Height =284
                    TabIndex =104
                    Name ="PAY_MAMELTX"
                    ControlSource ="PAY_MAMELTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =12640
                    Width =1648
                    Height =284
                    TabIndex =105
                    Name ="PAZ_PAZTX"
                    ControlSource ="PAZ_PAZTX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =12300
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label255"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =13044
                    Width =1648
                    Height =284
                    TabIndex =106
                    Name ="PAZ_HESTX"
                    ControlSource ="PAZ_HESTX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =13448
                    Width =1648
                    Height =284
                    TabIndex =107
                    Name ="PAZ_CEOTX"
                    ControlSource ="PAZ_CEOTX"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =12640
                    Width =1648
                    Height =284
                    TabIndex =108
                    Name ="SGN0124TX"
                    ControlSource ="SGN0124TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4274
                    Top =12300
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label259"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =13044
                    Width =1648
                    Height =284
                    TabIndex =109
                    Name ="SGN0224TX"
                    ControlSource ="SGN0224TX"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =13448
                    Width =1648
                    Height =284
                    TabIndex =110
                    Name ="SGN0324TX"
                    ControlSource ="SGN0324TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =14335
                    Width =1648
                    Height =284
                    TabIndex =111
                    Name ="SGN0135TX"
                    ControlSource ="SGN0135TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =13995
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label263"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =14739
                    Width =1648
                    Height =284
                    TabIndex =112
                    Name ="SGN0235TX"
                    ControlSource ="SGN0235TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =15143
                    Width =1648
                    Height =284
                    TabIndex =113
                    Name ="SGN0335TX"
                    ControlSource ="SGN0335TX"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =14290
                    Width =1648
                    Height =284
                    TabIndex =114
                    Name ="SGN0126TX"
                    ControlSource ="SGN0126TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =215
                    TextAlign =2
                    Left =4279
                    Top =13950
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label267"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =14694
                    Width =1648
                    Height =284
                    TabIndex =115
                    Name ="SGN0226TX"
                    ControlSource ="SGN0226TX"
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =15098
                    Width =1648
                    Height =284
                    TabIndex =116
                    Name ="SGN0326TX"
                    ControlSource ="SGN0326TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =15985
                    Width =1648
                    Height =284
                    TabIndex =117
                    Name ="SGN0132TX"
                    ControlSource ="SGN0132TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =15645
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label271"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =16389
                    Width =1648
                    Height =284
                    TabIndex =118
                    Name ="SGN0232TX"
                    ControlSource ="SGN0232TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =16793
                    Width =1648
                    Height =284
                    TabIndex =119
                    Name ="SGN0332TX"
                    ControlSource ="SGN0332TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =15925
                    Width =1648
                    Height =284
                    TabIndex =120
                    Name ="SGN0133TX"
                    ControlSource ="SGN0133TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4274
                    Top =15585
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label275"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =16329
                    Width =1648
                    Height =284
                    TabIndex =121
                    Name ="SGN0233TX"
                    ControlSource ="SGN0233TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =16733
                    Width =1648
                    Height =284
                    TabIndex =122
                    Name ="SGN0333TX"
                    ControlSource ="SGN0333TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =18085
                    Width =1648
                    Height =284
                    TabIndex =123
                    Name ="SGN0134TX"
                    ControlSource ="SGN0134TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =453
                    Top =17745
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label279"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =18489
                    Width =1648
                    Height =284
                    TabIndex =124
                    Name ="SGN0234TX"
                    ControlSource ="SGN0234TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =225
                    Top =18893
                    Width =1648
                    Height =284
                    TabIndex =125
                    Name ="SGN0334TX"
                    ControlSource ="SGN0334TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =18085
                    Width =1648
                    Height =284
                    TabIndex =126
                    Name ="SGN0136TX"
                    ControlSource ="SGN0136TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4274
                    Top =17745
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label283"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =18489
                    Width =1648
                    Height =284
                    TabIndex =127
                    Name ="SGN0236TX"
                    ControlSource ="SGN0236TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4047
                    Top =18893
                    Width =1648
                    Height =284
                    TabIndex =128
                    Name ="SGN0336TX"
                    ControlSource ="SGN0336TX"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2036
                    Top =19950
                    TabIndex =129
                    Name ="SGN0137"
                    ControlSource ="SGN0137"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2374
                            Top =19957
                            Width =1155
                            Height =240
                            Name ="Label287"
                            Caption ="درخواست كننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2040
                    Top =20355
                    TabIndex =130
                    Name ="SGN0237"
                    ControlSource ="SGN0237"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2377
                            Top =20355
                            Width =765
                            Height =240
                            Name ="Label289"
                            Caption ="مدير واحد"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2036
                    Top =20745
                    TabIndex =131
                    Name ="SGN0337"
                    ControlSource ="SGN0337"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2373
                            Top =20745
                            Width =780
                            Height =240
                            Name ="Label291"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =19732
                    Width =3560
                    Height =1399
                    Name ="Box292"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1830
                    Top =19560
                    Width =1785
                    Height =240
                    BackColor =13434828
                    Name ="Label293"
                    Caption ="درخواست مرخصي"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =237
                    Top =19900
                    Width =1648
                    Height =284
                    TabIndex =132
                    Name ="SGN0137TX"
                    ControlSource ="SGN0137TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =464
                    Top =19560
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label295"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =237
                    Top =20304
                    Width =1648
                    Height =284
                    TabIndex =133
                    Name ="SGN0237TX"
                    ControlSource ="SGN0237TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =237
                    Top =20708
                    Width =1648
                    Height =284
                    TabIndex =134
                    Name ="SGN0337TX"
                    ControlSource ="SGN0337TX"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5827
                    Top =19949
                    TabIndex =135
                    Name ="SGN0140"
                    ControlSource ="SGN0140"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6165
                            Top =19956
                            Width =1155
                            Height =240
                            Name ="Label299"
                            Caption ="تعميرات"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5831
                    Top =20354
                    TabIndex =136
                    Name ="SGN0240"
                    ControlSource ="SGN0240"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6165
                            Top =20355
                            Width =795
                            Height =240
                            Name ="Label301"
                            Caption ="حسابداري"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5827
                    Top =20744
                    TabIndex =137
                    Name ="SGN0340"
                    ControlSource ="SGN0340"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6164
                            Top =20744
                            Width =780
                            Height =240
                            Name ="Label303"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3911
                    Top =19731
                    Width =3560
                    Height =1399
                    Name ="Box304"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5610
                    Top =19560
                    Width =1815
                    Height =240
                    BackColor =13434828
                    Name ="Label305"
                    Caption ="تعريف تجهيزات PM"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4028
                    Top =19899
                    Width =1648
                    Height =284
                    TabIndex =138
                    Name ="SGN0140TX"
                    ControlSource ="SGN0140TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4255
                    Top =19559
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label307"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4028
                    Top =20303
                    Width =1648
                    Height =284
                    TabIndex =139
                    Name ="SGN0240TX"
                    ControlSource ="SGN0240TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4028
                    Top =20707
                    Width =1648
                    Height =284
                    TabIndex =140
                    Name ="SGN0340TX"
                    ControlSource ="SGN0340TX"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5827
                    Top =21640
                    TabIndex =141
                    Name ="SGN0141"
                    ControlSource ="SGN0141"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6165
                            Top =21647
                            Width =1155
                            Height =240
                            Name ="Label311"
                            Caption ="درخواست کننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5831
                    Top =22045
                    TabIndex =142
                    Name ="SGN0241"
                    ControlSource ="SGN0241"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6165
                            Top =22046
                            Width =795
                            Height =240
                            Name ="Label313"
                            Caption ="تاييد کننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =5827
                    Top =22435
                    TabIndex =143
                    Name ="SGN0341"
                    ControlSource ="SGN0341"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6164
                            Top =22435
                            Width =780
                            Height =240
                            Name ="Label315"
                            Caption ="انبار"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3911
                    Top =21422
                    Width =3560
                    Height =1399
                    Name ="Box316"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =5595
                    Top =21255
                    Width =1845
                    Height =240
                    BackColor =13434828
                    Name ="Label317"
                    Caption ="درخواست کالا از انبار"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4028
                    Top =21590
                    Width =1648
                    Height =284
                    TabIndex =144
                    Name ="SGN0141TX"
                    ControlSource ="SGN0141TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =4255
                    Top =21250
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label319"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4028
                    Top =21994
                    Width =1648
                    Height =284
                    TabIndex =145
                    Name ="SGN0241TX"
                    ControlSource ="SGN0241TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4028
                    Top =22398
                    Width =1648
                    Height =284
                    TabIndex =146
                    Name ="SGN0341TX"
                    ControlSource ="SGN0341TX"
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2051
                    Top =21645
                    TabIndex =147
                    Name ="SGN0142"
                    ControlSource ="SGN0142"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2389
                            Top =21652
                            Width =1155
                            Height =240
                            Name ="Label323"
                            Caption ="اعلام کننده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2055
                    Top =22050
                    TabIndex =148
                    Name ="SGN0242"
                    ControlSource ="SGN0242"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2389
                            Top =22051
                            Width =795
                            Height =240
                            Name ="Label325"
                            Caption ="تعميرات"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =2051
                    Top =22440
                    TabIndex =149
                    Name ="SGN0342"
                    ControlSource ="SGN0342"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2388
                            Top =22440
                            Width =780
                            Height =240
                            Name ="Label327"
                            Caption ="مدير عامل"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =135
                    Top =21427
                    Width =3560
                    Height =1399
                    Name ="Box328"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =1800
                    Top =21255
                    Width =1875
                    Height =240
                    BackColor =13434828
                    Name ="Label329"
                    Caption ="اعلام خرابي"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =252
                    Top =21595
                    Width =1648
                    Height =284
                    TabIndex =150
                    Name ="SGN0142TX"
                    ControlSource ="SGN0142TX"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =479
                    Top =21255
                    Width =975
                    Height =240
                    BackColor =16763904
                    Name ="Label331"
                    Caption ="عنوان در چاپ"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =252
                    Top =21999
                    Width =1648
                    Height =284
                    TabIndex =151
                    Name ="SGN0242TX"
                    ControlSource ="SGN0242TX"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =252
                    Top =22403
                    Width =1648
                    Height =284
                    TabIndex =152
                    Name ="SGN0342TX"
                    ControlSource ="SGN0342TX"
                End
            End
        End
    End
End
CodeBehindForm
' See "SIGN_SUB.cls"
