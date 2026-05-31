Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =13
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12566
    DatasheetFontHeight =10
    ItemSuffix =62
    Left =495
    Top =1680
    Right =15165
    Bottom =8175
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x2f58214ef201e640
    End
    RecordSource ="HEAD_LST_EXTENDED"
    Caption ="اطلاعات سامانه موديان"
    AfterUpdate ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =6066
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8220
                    Top =345
                    Width =945
                    Height =285
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =9346
                            Top =345
                            Width =525
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8220
                    Top =685
                    Width =945
                    Height =285
                    TabIndex =1
                    Name ="tgu"
                    ControlSource ="tgu"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =9346
                            Top =685
                            Width =300
                            Height =240
                            Name ="Label3"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =2205
                    Width =2475
                    Height =340
                    TabIndex =5
                    Name ="sbc"
                    ControlSource ="sbc"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =2205
                            Width =1350
                            Height =240
                            Name ="Label13"
                            Caption ="کد شعبه فروشنده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =2610
                    Width =2475
                    Height =340
                    TabIndex =6
                    Name ="bbc"
                    ControlSource ="bbc"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =2610
                            Width =1155
                            Height =240
                            Name ="Label15"
                            Caption ="کد شعبه خريدار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =3000
                    Width =2475
                    Height =340
                    TabIndex =7
                    Name ="ft"
                    ControlSource ="ft"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =3000
                            Width =645
                            Height =240
                            Name ="Label17"
                            Caption ="نوع پرواز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =3405
                    Width =2475
                    Height =340
                    TabIndex =8
                    Name ="bpn"
                    ControlSource ="bpn"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =3405
                            Width =1530
                            Height =240
                            Name ="Label19"
                            Caption ="شماره گذرنامه خريدار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =3795
                    Width =2475
                    Height =340
                    TabIndex =9
                    Name ="scln"
                    ControlSource ="scln"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =3795
                            Width =1485
                            Height =240
                            Name ="Label21"
                            Caption ="شماره پروانه گمرکي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =4200
                    Width =2475
                    Height =340
                    TabIndex =10
                    Name ="scc"
                    ControlSource ="scc"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =4200
                            Width =2100
                            Height =240
                            Name ="Label23"
                            Caption ="کد گمرک محل اظهار فروشنده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =4590
                    Width =2475
                    Height =340
                    TabIndex =11
                    Name ="cdcn"
                    ControlSource ="cdcn"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =4590
                            Width =2070
                            Height =240
                            Name ="Label25"
                            Caption ="شماره کوتاژ اظهارنامه گمرکي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =4980
                    Width =2475
                    Height =340
                    TabIndex =12
                    Name ="cdcd"
                    ControlSource ="cdcd"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =4980
                            Width =1935
                            Height =240
                            Name ="Label27"
                            Caption ="تاريخ کوتاژ اظهارنامه گمرکي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =6689
                    Top =5385
                    Width =2475
                    Height =340
                    TabIndex =13
                    Name ="crn"
                    ControlSource ="crn"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9345
                            Top =5385
                            Width =2490
                            Height =240
                            Name ="Label29"
                            Caption ="شناسه يکتاي ثبت قرارداد فروشنده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =1050
                    Width =2475
                    Height =340
                    TabIndex =14
                    Name ="billid"
                    ControlSource ="billid"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =1050
                            Width =2790
                            Height =240
                            Name ="Label31"
                            Caption ="شماره اشتراک /شناسه قبض بهره بردار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =1450
                    Width =2475
                    Height =340
                    TabIndex =15
                    Name ="todam"
                    ControlSource ="todam"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =1450
                            Width =2970
                            Height =240
                            Name ="Label33"
                            Caption ="مجموع ساير ماليات، عوارض و وجوه قانوني"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =1840
                    Width =2475
                    Height =340
                    TabIndex =16
                    Name ="tonw"
                    ControlSource ="tonw"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =1840
                            Width =1275
                            Height =240
                            Name ="Label35"
                            Caption ="مجموع وزن خالص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =2235
                    Width =2475
                    Height =340
                    TabIndex =17
                    Name ="torv"
                    ControlSource ="torv"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =2235
                            Width =1365
                            Height =240
                            Name ="Label37"
                            Caption ="مجموع ارزش ريالي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =2640
                    Width =2475
                    Height =340
                    TabIndex =18
                    Name ="tocv"
                    ControlSource ="tocv"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =2640
                            Width =1320
                            Height =240
                            Name ="Label39"
                            Caption ="مجموع ارزش ارزي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =3435
                    Width =2475
                    Height =340
                    TabIndex =20
                    Name ="cap"
                    ControlSource ="cap"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =3435
                            Width =1410
                            Height =240
                            Name ="Label43"
                            Caption ="مبلغ پرداختي نقدي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =3825
                    Width =2475
                    Height =340
                    TabIndex =21
                    Name ="insp"
                    ControlSource ="insp"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =3825
                            Width =795
                            Height =240
                            Name ="Label45"
                            Caption ="مبلغ نسيه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =4230
                    Width =2475
                    Height =340
                    TabIndex =22
                    Name ="tvop"
                    ControlSource ="tvop"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =4230
                            Width =3165
                            Height =240
                            Name ="Label47"
                            Caption ="مجموع سهم ماليات بر ارزش افزوده از پرداخت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =4620
                    Width =2475
                    Height =340
                    TabIndex =23
                    Name ="tax17"
                    ControlSource ="tax17"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =4620
                            Width =1575
                            Height =240
                            Name ="Label49"
                            Caption ="ماليات موضوع ماده 17"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =5010
                    Width =2475
                    Height =340
                    TabIndex =24
                    Name ="CRT"
                    ControlSource ="CRT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =5010
                            Width =390
                            Height =240
                            Name ="Label51"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =5415
                    Width =2475
                    Height =340
                    TabIndex =25
                    Name ="UID"
                    ControlSource ="UID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =5415
                            Width =375
                            Height =240
                            Name ="Label53"
                            Caption ="کاربر"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6690
                    Top =1025
                    Width =2475
                    Height =345
                    TabIndex =2
                    Name ="inty"
                    ControlSource ="inty"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نوع اول\";2;\"نوع دوم\";3;\"نوع سوم\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9346
                            Top =1025
                            Width =1125
                            Height =240
                            Name ="Label5"
                            Caption ="نوع صورتحساب"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6690
                    Top =1425
                    Width =2475
                    Height =345
                    TabIndex =3
                    Name ="inp"
                    ControlSource ="inp"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فروش\";2;\"فروش ارزي\";3;\"طلاوجواهر\";4;\"پيمانکاري\";5;\"قبوض خدماتي\";6;\""
                        "بليط هواپيما\";7;\"صادرات\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9346
                            Top =1425
                            Width =1305
                            Height =240
                            Name ="Label7"
                            Caption ="الگوي صورتحساب"
                        End
                    End
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6690
                    Top =1815
                    Width =2475
                    Height =345
                    TabIndex =4
                    Name ="ins"
                    ControlSource ="ins"
                    RowSourceType ="Value List"
                    RowSource ="1;\"اصلي\";2;\"اصلاحي\";3;\"ابطالي\";4;\"برگشت فروش\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9346
                            Top =1815
                            Width =1380
                            Height =240
                            Name ="Label9"
                            Caption ="موضوع صورتحساب"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9977
                    Top =340
                    Width =2280
                    Height =405
                    TabIndex =26
                    Name ="Command55"
                    Caption ="ارسال فاکتور به سامانه موديان"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =450
                    Top =3030
                    Width =2475
                    Height =340
                    TabIndex =19
                    Name ="setm"
                    ControlSource ="setm"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نقد\";2;\"نسيه\";3;\"نقد/نسيه\""
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =3030
                            Width =885
                            Height =240
                            Name ="Label41"
                            Caption ="روش تسويه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =453
                    Top =623
                    Width =2475
                    Height =340
                    TabIndex =27
                    Name ="CUT"
                    ControlSource ="CUT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Code, Title, CountryName FROM TCOD_ARZ ORDER BY Code"
                    ColumnWidths ="0;567"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3109
                            Top =623
                            Width =885
                            Height =240
                            Name ="Label59"
                            Caption ="نوع ارز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    Left =450
                    Top =225
                    Width =2475
                    Height =340
                    TabIndex =28
                    Name ="irtaxid"
                    ControlSource ="irtaxid"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3106
                            Top =225
                            Width =3255
                            Height =240
                            Name ="Label61"
                            Caption ="شماره منحصر به فرد مالياتي صورتحساب مرجع"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "HEAD_LST_EXTENDED_FORM.cls"
