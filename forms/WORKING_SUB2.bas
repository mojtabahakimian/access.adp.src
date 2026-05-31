Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4598
    RowHeight =270
    ItemSuffix =31
    Left =315
    Top =855
    Right =15135
    Bottom =8775
    DatasheetGridlinesColor =12632256
    OrderBy ="PCODE"
    RecSrcDt = Begin
        0xf7815bf0ba81e240
    End
    RecordSource ="WORKING"
    Caption ="WORKING"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =7005
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =712
                    Top =453
                    ColumnWidth =990
                    ColumnOrder =1
                    TabIndex =1
                    Name ="WDATE"
                    ControlSource ="WDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3894
                            Top =453
                            Width =390
                            Height =240
                            Name ="Label3"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =712
                    Top =793
                    Height =255
                    ColumnWidth =615
                    ColumnOrder =7
                    TabIndex =2
                    Name ="DAYS"
                    ControlSource ="DAYS"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3249
                            Top =793
                            Width =1035
                            Height =240
                            Name ="Label5"
                            Caption ="كاركرد.ا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =712
                    Top =1133
                    ColumnWidth =1020
                    ColumnOrder =11
                    TabIndex =3
                    Name ="EZAFAH"
                    ControlSource ="EZAFAH"
                    Format ="#,###"
                    AfterUpdate ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3594
                            Top =1133
                            Width =690
                            Height =240
                            Name ="Label7"
                            Caption ="اضافه كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =680
                    Top =1421
                    ColumnWidth =975
                    ColumnOrder =14
                    TabIndex =4
                    Name ="PADASH"
                    ControlSource ="PADASH"
                    Format ="#,###"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3778
                            Top =1417
                            Width =555
                            Height =240
                            Name ="Label9"
                            Caption ="راندمان"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =712
                    Top =1814
                    ColumnWidth =975
                    ColumnOrder =15
                    TabIndex =5
                    Name ="KASR_VAM"
                    ControlSource ="KASR_VAM"
                    Format ="#,###"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3609
                            Top =1814
                            Width =675
                            Height =240
                            Name ="Label11"
                            Caption ="كسر وام"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =3
                    Left =712
                    Top =113
                    Width =1776
                    Height =255
                    ColumnWidth =3075
                    ColumnOrder =0
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT PERSONEL.CODE, PERSONEL.PNAME + N' ' + PERSONEL.PFAMILY AS PER, PERSONEL."
                        "CODE AS Expr1 FROM PERSONEL INNER JOIN PHOKM ON PERSONEL.CODE = PHOKM.CODE ORDER"
                        " BY PERSONEL.PNAME + N' ' + PERSONEL.PFAMILY"
                    ColumnWidths ="0;1985;567"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3534
                            Top =113
                            Width =750
                            Height =240
                            Name ="Label1"
                            Caption ="كد پرسنل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =719
                    Top =2100
                    ColumnWidth =945
                    ColumnOrder =9
                    TabIndex =7
                    Name ="EZAD"
                    ControlSource ="EZAD"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="000:00"
                    ControlTipText ="ساعت اضافه كار عادي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3206
                            Top =2100
                            Width =915
                            Height =240
                            Name ="Label13"
                            Caption ="س.ا.ع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =707
                    Top =2378
                    ColumnWidth =645
                    ColumnOrder =10
                    TabIndex =8
                    Name ="EZTAT"
                    ControlSource ="EZTAT"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="000:00"
                    ControlTipText ="ساعت اضافه كار تعطيلي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3164
                            Top =2378
                            Width =945
                            Height =240
                            Name ="Label14"
                            Caption ="س.ا.ت"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =707
                    Top =2661
                    ColumnWidth =525
                    ColumnOrder =4
                    TabIndex =9
                    Name ="KHADAMAT"
                    ControlSource ="KHADAMAT"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3164
                            Top =2661
                            Width =945
                            Height =240
                            Name ="Label15"
                            Caption ="خدمات"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =707
                    Top =2938
                    ColumnWidth =495
                    ColumnOrder =5
                    TabIndex =10
                    Name ="FROSH"
                    ControlSource ="FROSH"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3164
                            Top =2938
                            Width =945
                            Height =240
                            Name ="Label16"
                            Caption ="فروش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =705
                    Top =3220
                    ColumnWidth =435
                    ColumnOrder =2
                    TabIndex =11
                    Name ="TOLID"
                    ControlSource ="TOLID"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3162
                            Top =3220
                            Width =945
                            Height =240
                            Name ="Label17"
                            Caption ="توليد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =705
                    Top =3511
                    Height =255
                    ColumnWidth =495
                    ColumnOrder =3
                    TabIndex =6
                    Name ="HAZCENTER"
                    ControlSource ="HAZCENTER"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3222
                            Top =3526
                            Width =885
                            Height =240
                            Name ="Label12"
                            Caption ="اداري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =705
                    Top =3795
                    ColumnWidth =735
                    ColumnOrder =12
                    TabIndex =12
                    Name ="KASRS"
                    ControlSource ="KASRS"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="000:00"
                    ControlTipText ="ساعت اضافه كار تعطيلي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3162
                            Top =3795
                            Width =945
                            Height =240
                            Name ="Label19"
                            Caption ="س.ك.ك"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =705
                    Top =4155
                    ColumnWidth =1020
                    ColumnOrder =13
                    TabIndex =13
                    Name ="KASRM"
                    ControlSource ="KASRM"
                    Format ="#,###"
                    AfterUpdate ="[Event Procedure]"
                    ControlTipText ="ساعت اضافه كار تعطيلي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3162
                            Top =4155
                            Width =945
                            Height =240
                            Name ="Label21"
                            Caption ="كسر كار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =705
                    Top =4515
                    ColumnWidth =495
                    ColumnOrder =8
                    TabIndex =14
                    Name ="TDAYS"
                    ControlSource ="TDAYS"
                    Format ="Fixed"
                    ControlTipText ="ساعت اضافه كار تعطيلي"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3162
                            Top =4515
                            Width =945
                            Height =240
                            Name ="Label23"
                            Caption ="ت.ك"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =4818
                    Height =255
                    ColumnWidth =645
                    ColumnOrder =6
                    TabIndex =15
                    Name ="DAYSB"
                    ControlSource ="DAYSB"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3103
                            Top =4818
                            Width =1035
                            Height =240
                            Name ="Label25"
                            Caption ="كاركرد.ر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1643
                    Top =5385
                    ColumnWidth =1335
                    ColumnOrder =16
                    TabIndex =16
                    Name ="ECOL1"
                    ControlSource ="ECOL1"
                    Format ="#,###"
                    ControlTipText ="مقدار درج شده در اين ستون به حقوق اضافه مي شود"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =5385
                            Width =660
                            Height =240
                            Name ="SAYER1"
                            Caption ="+ساير 1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1700
                    Top =5782
                    ColumnWidth =1335
                    ColumnOrder =17
                    TabIndex =17
                    Name ="ECOL2"
                    ControlSource ="ECOL2"
                    Format ="#,###"
                    ControlTipText ="مقدار درج شده در اين ستون به حقوق اضافه مي شود"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3983
                            Top =5782
                            Width =615
                            Height =240
                            Name ="SAYER2"
                            Caption ="+ساير2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1643
                    Top =6066
                    ColumnWidth =1290
                    ColumnOrder =18
                    TabIndex =18
                    Name ="ECOL3"
                    ControlSource ="ECOL3"
                    Format ="#,###"
                    ControlTipText ="مقدار درج شده در اين ستون به حقوق اضافه مي شود"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =6060
                            Width =660
                            Height =240
                            Name ="SAYER3"
                            Caption ="+ساير 3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1643
                    Top =6406
                    ColumnWidth =1290
                    ColumnOrder =20
                    TabIndex =19
                    Name ="KCOL1"
                    ControlSource ="KCOL1"
                    Format ="#,###"
                    ControlTipText ="مقدار درج شده در اين ستون از حقوق كسر مي شود"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3930
                            Top =6405
                            Width =660
                            Height =600
                            Name ="KOSORAT1"
                            Caption ="-ساير كسورات 1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =283
                    Top =5102
                    ColumnOrder =19
                    TabIndex =20
                    Name ="ECOL4"
                    ControlSource ="ECOL4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3075
                            Top =5100
                            Width =660
                            Height =240
                            Name ="SAYER4"
                            Caption ="+ساير 4"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =56
                    Top =3
                    ColumnWidth =2070
                    TabIndex =21
                    Name ="CRT"
                    ControlSource ="CRT"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2261
                            Width =1515
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Top =283
                    TabIndex =22
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="[Forms]![BASEKNOW]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2205
                            Top =285
                            Width =1365
                            Height =240
                            Name ="Label36"
                            Caption ="کد کاربر ايجاد کننده"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "WORKING_SUB2.cls"
