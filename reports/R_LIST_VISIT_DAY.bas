Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =6
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15363
    DatasheetFontHeight =10
    ItemSuffix =65
    Left =600
    Top =1710
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xed53051434a0e440
    End
    RecordSource ="SELECT VISITORS_DAY.HES, VISITORS_DAY.CDATE, VISITORS_DAY.USERNAME, VISITORS_DAY"
        ".OKF, CUST_HESAB.NAME, CUST_HESAB.ADDRESS, CUST_HESAB.TEL, CUST_HESAB.MOBILE, CU"
        "ST_HESAB_1.NAME AS VISITOR, VISITORS_DAY.VDATE, VISITORS_DAY_DTL.COUST_NO, VISIT"
        "ORS_DAY_DTL.CLASS, VISITORS_DAY_DTL.RACTIVE FROM VISITORS_DAY_DTL INNER JOIN VIS"
        "ITORS_DAY ON VISITORS_DAY_DTL.HES = VISITORS_DAY.HES AND VISITORS_DAY_DTL.VDATE "
        "= VISITORS_DAY.VDATE INNER JOIN CUST_HESAB CUST_HESAB_1 ON CUST_HESAB_1.hes = VI"
        "SITORS_DAY.HES INNER JOIN CUST_HESAB ON VISITORS_DAY_DTL.COUST_NO = CUST_HESAB.h"
        "es"
    Caption ="ليست مشتريان ويزيتور در اين تاريخ"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial"
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
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Times New Roman"
            AsianLineBreak =255
        End
        Begin ListBox
            TextFontFamily =18
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Times New Roman"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontFamily =18
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Times New Roman"
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
            GroupFooter = NotDefault
            ControlSource ="VDATE"
        End
        Begin PageHeader
            Height =922
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =5895
                    Top =120
                    Width =5676
                    Height =285
                    FontSize =10
                    Name ="VISITOR"
                    ControlSource ="VISITOR"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =11640
                            Top =120
                            Width =660
                            Height =285
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label32"
                            Caption ="ويزيتور:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Top =135
                    Width =1365
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text46"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    Left =1814
                    Width =3225
                    Height =435
                    FontSize =16
                    FontWeight =400
                    BackColor =10092543
                    ForeColor =0
                    Name ="Label47"
                    Caption ="فهرست مشتريان"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    BackStyle =1
                    Top =510
                    Width =15363
                    Height =412
                    BackColor =10092543
                    Name ="Box48"
                End
                Begin Label
                    TextAlign =2
                    Left =10016
                    Top =583
                    Width =3411
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="COUST_NO_Label"
                    Caption ="مشتري"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =13477
                    Top =583
                    Width =1311
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label30"
                    Caption ="كد مشتري"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =4290
                    Top =573
                    Width =5616
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label34"
                    Caption ="آدرس"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2895
                    Top =588
                    Width =1365
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label37"
                    Caption ="تلفن ثابت"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1200
                    Top =588
                    Width =1395
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label38"
                    Caption ="تلفن همراه"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =14865
                    Top =573
                    Width =495
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label50"
                    Caption ="رديف"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =90
                    Top =588
                    Width =885
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label60"
                    Caption ="كلاس"
                    FontName ="Tahoma"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =12358
                    Top =120
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="VDATE"
                    ControlSource ="VDATE"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =0
                            Left =14115
                            Top =120
                            Width =1215
                            Height =285
                            FontSize =10
                            Name ="Label63"
                            Caption ="تاريخ ويزيت:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5272
                    TabIndex =3
                    Name ="HES"
                    ControlSource ="HES"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =396
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =10020
                    Width =3426
                    Height =345
                    FontSize =10
                    Name ="NAME1"
                    ControlSource ="NAME"
                    FontName ="Tahoma"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BorderWidth =1
                    Left =9237
                    Width =275
                    Height =300
                    TabIndex =1
                    Name ="RACTIVE"
                    ControlSource ="RACTIVE"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =13472
                    Width =1371
                    Height =345
                    FontSize =10
                    TabIndex =2
                    Name ="COUST_NO1"
                    ControlSource ="COUST_NO"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4320
                    Width =5616
                    Height =345
                    FontSize =10
                    TabIndex =3
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    BorderWidth =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =2805
                    Width =1461
                    Height =345
                    FontSize =10
                    TabIndex =4
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =1140
                    Width =1626
                    Height =345
                    FontSize =10
                    TabIndex =5
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =14880
                    Width =471
                    Height =345
                    FontSize =10
                    TabIndex =6
                    Name ="Text49"
                    ControlSource ="=1"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =14865
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line52"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =13455
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line53"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =10005
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line54"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =4308
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line55"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2790
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line56"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line58"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =15345
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line59"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =34
                    IMESentenceMode =3
                    Width =1086
                    Height =345
                    FontSize =10
                    TabIndex =7
                    Name ="CLASS"
                    ControlSource ="CLASS"
                    FontName ="Tahoma"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1110
                    Width =0
                    Height =396
                    BorderColor =0
                    Name ="Line62"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =56
            Name ="GroupFooter0"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =15363
                    Name ="Line57"
                End
            End
        End
        Begin PageFooter
            Height =240
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =225
                    Width =1260
                    FontSize =8
                    Name ="Text25"
                    ControlSource ="=Now()"
                    Format ="Short Time"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =8850
                    Width =5040
                    FontSize =8
                    TabIndex =1
                    Name ="Text26"
                    ControlSource ="=\"صفحه  \" & [Page] & \" از  \" & [Pages]"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_LIST_VISIT_DAY.cls"
