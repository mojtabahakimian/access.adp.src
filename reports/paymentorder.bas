Version =20
VersionRequired =20
Begin Report
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7880
    RowHeight =465
    DatasheetFontHeight =10
    ItemSuffix =78
    Left =3255
    Top =60
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xc90f82bb318fe540
    End
    ServerFilter ="IDD= 5"
    RecordSource ="SELECT payorder.IDD, payorder.MABLS, payorder.BABAT, payorder.CUST_NO, payorder."
        "CUST_NO_TXT, payorder.ORDERER, payorder.PAYDATE, payorder.DTT, payorder.TOZIH, p"
        "ayorder.CACHTIC, payorder.CHACHMABL, payorder.CHEKTIC, payorder.CHEKMABL, payord"
        "er.CHEKFAGH, payorder.CHEKDIST, payorder.CHEKMTIC, payorder.CHEKMMABL, payorder."
        "CHEKMFAGH, payorder.CHEKMDIST, payorder.SGN1, payorder.SGN2, payorder.SGN3, CUST"
        "_HESAB.NAME AS nam, CUST_HESAB_1.NAME AS orderername FROM payorder INNER JOIN CU"
        "ST_HESAB ON payorder.CUST_NO = CUST_HESAB.hes INNER JOIN CUST_HESAB CUST_HESAB_1"
        " ON payorder.ORDERER = CUST_HESAB_1.hes"
    Caption ="فرم درخواست پرداخت\015\012"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin PageHeader
            Height =643
            BackColor =12632256
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =1
                    TextFontFamily =2
                    Left =3240
                    Width =1575
                    Height =495
                    FontSize =14
                    Name ="Label77"
                    Caption ="دستور  پرداخت"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =120
                    Top =225
                    Width =951
                    Height =360
                    ColumnWidth =555
                    ColumnOrder =0
                    FontSize =10
                    BackColor =0
                    Name ="IDD"
                    ControlSource ="IDD"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =1
                            TextFontFamily =2
                            Left =1167
                            Top =225
                            Width =525
                            Height =360
                            FontSize =10
                            Name ="Label1"
                            Caption ="شماره:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =5835
                    Top =225
                    Width =1290
                    Height =360
                    ColumnOrder =1
                    FontSize =10
                    TabIndex =1
                    Name ="PAYDATE"
                    ControlSource ="PAYDATE"
                    DefaultValue ="=farsidate(Date())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =7215
                            Top =225
                            Width =540
                            Height =360
                            FontSize =10
                            Name ="Label11"
                            Caption ="تاريخ:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =7655
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =4358
                    Top =226
                    Height =360
                    ColumnOrder =5
                    FontSize =10
                    Name ="MABLS"
                    ControlSource ="MABLS"
                    Format ="#,###"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6225
                            Top =225
                            Width =390
                            Height =360
                            FontSize =10
                            Name ="Label3"
                            Caption ="مبلغ:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =690
                    Width =6045
                    Height =675
                    ColumnWidth =4110
                    ColumnOrder =6
                    FontSize =10
                    TabIndex =1
                    Name ="BABAT"
                    ControlSource ="BABAT"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6240
                            Top =697
                            Width =435
                            Height =360
                            FontSize =10
                            Name ="Label5"
                            Caption ="بايت:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =2047
                    Width =6045
                    Height =360
                    ColumnWidth =3930
                    ColumnOrder =4
                    FontSize =10
                    TabIndex =3
                    Name ="CUST_NO_TXT"
                    ControlSource ="CUST_NO_TXT"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6210
                            Top =2040
                            Width =1605
                            Height =360
                            FontSize =10
                            Name ="Label9"
                            Caption ="دريافت کننده(نماينده):"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =15
                    Top =2550
                    Width =6045
                    Height =2835
                    ColumnOrder =7
                    FontSize =10
                    TabIndex =4
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6195
                            Top =2599
                            Width =825
                            Height =360
                            FontSize =10
                            Name ="Label15"
                            Caption ="توضيحات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3753
                    Top =5955
                    Height =360
                    ColumnOrder =9
                    FontSize =10
                    TabIndex =7
                    BackColor =10079487
                    Name ="CHACHMABL"
                    ControlSource ="CHACHMABL"
                    Format ="#,###"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5624
                            Top =6011
                            Width =345
                            Height =360
                            FontSize =10
                            Name ="Label19"
                            Caption ="مبلغ"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =705
                    Top =6510
                    Width =636
                    Height =360
                    ColumnWidth =705
                    ColumnOrder =13
                    FontSize =10
                    TabIndex =11
                    BackColor =10092543
                    Name ="CHEKDIST"
                    ControlSource ="CHEKDIST"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1501
                            Top =6556
                            Width =615
                            Height =360
                            FontSize =10
                            BackColor =10092543
                            Name ="Label27"
                            Caption ="با فاصله"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    Left =6948
                    Top =6045
                    ColumnOrder =8
                    TabIndex =6
                    Name ="CACHTIC"
                    ControlSource ="CACHTIC"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =6529
                            Top =6006
                            Width =315
                            Height =360
                            FontSize =10
                            BackColor =10079487
                            Name ="Label44"
                            Caption ="نقد"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3738
                    Top =6525
                    Height =360
                    ColumnOrder =11
                    FontSize =10
                    TabIndex =9
                    BackColor =10092543
                    Name ="CHEKMABL"
                    ControlSource ="CHEKMABL"
                    Format ="#,###"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5609
                            Top =6581
                            Width =345
                            Height =360
                            FontSize =10
                            BackColor =10092543
                            Name ="Label56"
                            Caption ="مبلغ"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    Left =6933
                    Top =6615
                    ColumnOrder =10
                    TabIndex =8
                    BorderColor =10092543
                    Name ="CHEKTIC"
                    ControlSource ="CHEKTIC"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =6048
                            Top =6570
                            Width =780
                            Height =360
                            FontSize =10
                            BackColor =10092543
                            Name ="Label58"
                            Caption ="چک شرکت"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =3738
                    Top =7200
                    Height =360
                    ColumnOrder =15
                    FontSize =10
                    TabIndex =13
                    BackColor =13434828
                    Name ="CHEKMMABL"
                    ControlSource ="CHEKMMABL"
                    Format ="#,###"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =5609
                            Top =7256
                            Width =345
                            Height =360
                            FontSize =10
                            BackColor =13434828
                            Name ="Label60"
                            Caption ="مبلغ"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    SpecialEffect =0
                    OldBorderStyle =0
                    Left =6933
                    Top =7290
                    ColumnOrder =14
                    TabIndex =12
                    Name ="CHEKMTIC"
                    ControlSource ="CHEKMTIC"
                    DefaultValue ="False"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =6033
                            Top =7245
                            Width =795
                            Height =360
                            FontSize =10
                            BackColor =13434828
                            Name ="Label62"
                            Caption ="چک مشتري"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2243
                    Top =6510
                    Width =606
                    Height =360
                    ColumnWidth =765
                    ColumnOrder =12
                    FontSize =10
                    TabIndex =10
                    BackColor =10092543
                    Name ="CHEKFAGH"
                    ControlSource ="CHEKFAGH"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2905
                            Top =6556
                            Width =795
                            Height =360
                            FontSize =10
                            BackColor =10092543
                            Name ="Label64"
                            Caption ="تعداد فقره"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =708
                    Top =7207
                    Width =636
                    Height =360
                    ColumnWidth =630
                    ColumnOrder =17
                    FontSize =10
                    TabIndex =15
                    BackColor =13434828
                    Name ="CHEKMDIST"
                    ControlSource ="CHEKMDIST"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1504
                            Top =7253
                            Width =540
                            Height =360
                            FontSize =10
                            BackColor =13434828
                            Name ="Label66"
                            Caption ="با راس"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =2246
                    Top =7207
                    Width =606
                    Height =360
                    ColumnOrder =16
                    FontSize =10
                    TabIndex =14
                    BackColor =13434828
                    Name ="CHEKMFAGH"
                    ControlSource ="CHEKMFAGH"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =2908
                            Top =7253
                            Width =795
                            Height =360
                            FontSize =10
                            BackColor =13434828
                            Name ="Label68"
                            Caption ="تعداد فقره"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextFontFamily =2
                    Left =285
                    Top =7254
                    Width =315
                    Height =360
                    FontSize =10
                    BackColor =13434828
                    Name ="Label75"
                    Caption ="روز"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =285
                    Top =6556
                    Width =315
                    Height =360
                    FontSize =10
                    BackColor =10092543
                    Name ="Label76"
                    Caption ="روز"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =15
                    Top =1492
                    Width =6045
                    Height =360
                    ColumnWidth =4845
                    ColumnOrder =3
                    FontSize =10
                    TabIndex =2
                    Name ="nam"
                    ControlSource ="nam"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6240
                            Top =1492
                            Width =1110
                            Height =360
                            FontSize =10
                            Name ="Label42"
                            Caption =" دريافت کننده: "
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =15
                    Top =5490
                    Width =6045
                    Height =360
                    FontSize =10
                    TabIndex =5
                    Name ="ORDERER"
                    ControlSource ="ORDERER"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6196
                            Top =5497
                            Width =1125
                            Height =360
                            FontSize =10
                            Name ="Label71"
                            Caption ="درخواست کننده"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Height =2097
            BackColor =10092543
            Name ="PageFooterSection"
            Begin
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =6434
                    Top =225
                    Width =284
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6707
                            Top =165
                            Width =720
                            Height =420
                            FontSize =10
                            Name ="LSGN1"
                            Caption ="مدير واحد"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =3420
                    Top =225
                    Width =284
                    TabIndex =1
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3678
                            Top =165
                            Width =690
                            Height =420
                            FontSize =10
                            Name ="LSGN2"
                            Caption ="حسابداري"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =765
                    Top =225
                    Width =284
                    TabIndex =2
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =1038
                            Top =165
                            Width =735
                            Height =420
                            FontSize =10
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "paymentorder.cls"
