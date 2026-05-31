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
    ItemSuffix =66
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x356545b0af05e540
    End
    RecordSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.HES, Visit_route.IYALAT, Visit_route."
        "CITY, Visit_route.District, Visit_route.CDATE, Visit_route.USERNAME, Visit_route"
        ".RACTIVE AS Visit_route_RACTIVE, Visit_route.OKF, Visit_route_dtl.COUST_NO, Visi"
        "t_route_dtl.RACTIVE AS Visit_route_dtl_RACTIVE, Visit_route_dtl.IDR, CUST_HESAB."
        "NAME, CUST_HESAB.ADDRESS, CUST_HESAB.TEL, CUST_HESAB.MOBILE, CUST_HESAB_1.NAME A"
        "S VISITOR, Visit_route_dtl.CLASS FROM Visit_route INNER JOIN Visit_route_dtl ON "
        "Visit_route.ROUTE_NAME = Visit_route_dtl.ROUTE_NAME INNER JOIN CUST_HESAB ON Vis"
        "it_route_dtl.COUST_NO = CUST_HESAB.hes INNER JOIN CUST_HESAB CUST_HESAB_1 ON Vis"
        "it_route.HES = CUST_HESAB_1.hes WHERE (Visit_route_dtl.RACTIVE = 1)"
    Caption ="فهرست مشتريان مسير فعال"
    OnOpen ="[Event Procedure]"
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
            ControlSource ="ROUTE_NAME"
        End
        Begin PageHeader
            Height =1489
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    HideDuplicates = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =13255
                    Top =566
                    Width =1311
                    Height =285
                    FontSize =10
                    Name ="IYALAT"
                    ControlSource ="IYALAT"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =2
                            Left =14625
                            Top =570
                            Width =660
                            Height =285
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label43"
                            Caption ="استان:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2835
                    Top =120
                    Width =5676
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="VISITOR"
                    ControlSource ="VISITOR"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =2
                            Left =8580
                            Top =120
                            Width =660
                            Height =285
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label32"
                            Caption ="ويزيتور:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =9450
                    Top =135
                    Width =5151
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8663
                    Top =566
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="District"
                    ControlSource ="District"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =0
                            TextFontFamily =2
                            Left =10420
                            Top =566
                            Width =690
                            Height =285
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label41"
                            Caption ="منطقه:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =14640
                    Top =120
                    Width =630
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label42"
                    Caption ="مسير:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =11214
                    Top =566
                    Width =1416
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="CITY"
                    ControlSource ="CITY"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =2
                            Left =12704
                            Top =570
                            Width =540
                            Height =285
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label45"
                            Caption ="شهر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Top =135
                    Width =1365
                    Height =300
                    FontSize =10
                    TabIndex =5
                    Name ="Text46"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =285
                    Top =570
                    Width =3915
                    Height =435
                    FontSize =16
                    FontWeight =400
                    BackColor =10092543
                    ForeColor =0
                    Name ="Label47"
                    Caption ="فهرست مشتريان مسير فعال"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BackStyle =1
                    Top =1077
                    Width =15363
                    Height =412
                    BackColor =16777164
                    Name ="Box48"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =10016
                    Top =1150
                    Width =3411
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="COUST_NO_Label"
                    Caption ="مشتري"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =13477
                    Top =1150
                    Width =1311
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label30"
                    Caption ="كد مشتري"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4290
                    Top =1140
                    Width =5616
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label34"
                    Caption ="آدرس"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =2895
                    Top =1155
                    Width =1365
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label37"
                    Caption ="تلفن ثابت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1200
                    Top =1155
                    Width =1395
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label38"
                    Caption ="تلفن همراه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =14865
                    Top =1140
                    Width =495
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label50"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =90
                    Top =1155
                    Width =885
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label60"
                    Caption ="كلاس"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
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
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =10020
                    Width =3426
                    Height =345
                    FontSize =10
                    Name ="NAME1"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin CheckBox
                    Visible = NotDefault
                    OldBorderStyle =0
                    BorderWidth =1
                    Left =9237
                    Width =275
                    Height =300
                    TabIndex =1
                    Name ="Visit_route_dtl_RACTIVE"
                    ControlSource ="Visit_route_dtl_RACTIVE"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =13472
                    Width =1371
                    Height =345
                    FontSize =10
                    TabIndex =2
                    Name ="COUST_NO1"
                    ControlSource ="COUST_NO"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    BorderWidth =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4320
                    Width =5676
                    Height =345
                    FontSize =10
                    TabIndex =3
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    BorderWidth =1
                    TextFontCharSet =178
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2805
                    Width =1461
                    Height =345
                    FontSize =10
                    TabIndex =4
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    BorderWidth =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1140
                    Width =1626
                    Height =345
                    FontSize =10
                    TabIndex =5
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =2
                    BorderWidth =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =14880
                    Width =471
                    Height =345
                    FontSize =10
                    TabIndex =6
                    Name ="Text49"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
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
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Width =1086
                    Height =345
                    FontSize =10
                    TabIndex =7
                    Name ="CLASS"
                    ControlSource ="CLASS"
                    FontName ="WeblogmaYekan"
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
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =225
                    Width =1260
                    FontSize =8
                    Name ="Text25"
                    ControlSource ="=Now()"
                    Format ="Short Time"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =8850
                    Width =5040
                    FontSize =8
                    TabIndex =1
                    Name ="Text26"
                    ControlSource ="=\"صفحه  \" & [Page] & \" از  \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_LIST_MASIR.cls"
