Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =7880
    DatasheetFontHeight =10
    ItemSuffix =52
    Left =3690
    Top =1440
    Right =11565
    Bottom =7515
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb821017d9ed4e240
    End
    RecordSource ="PAY_GETP"
    Caption ="برگشت چك پرداختي"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin FormHeader
            Height =540
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =10
                    Left =2354
                    Top =15
                    Width =2790
                    Height =525
                    FontSize =16
                    ForeColor =16777215
                    Name ="Label42"
                    Caption ="برگشت چك پرداختي"
                    FontName ="Sina"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =10
                    Left =2324
                    Width =2790
                    Height =525
                    FontSize =16
                    Name ="Label43"
                    Caption ="برگشت چك پرداختي"
                    FontName ="Sina"
                End
            End
        End
        Begin Section
            Height =5554
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4604
                    Top =1915
                    Width =1506
                    Height =315
                    TabIndex =5
                    Name ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6262
                            Top =1910
                            Width =1095
                            Height =300
                            Name ="Label5"
                            Caption ="تاريخ دريافت:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4044
                    Top =2276
                    Width =2061
                    Height =315
                    TabIndex =6
                    Name ="MABL"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6257
                            Top =2270
                            Width =480
                            Height =300
                            Name ="Label11"
                            Caption ="مبلغ:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3744
                    Top =2956
                    Width =2361
                    Height =315
                    TabIndex =8
                    Name ="N_HESAB"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6257
                            Top =2960
                            Width =1230
                            Height =300
                            Name ="Label15"
                            Caption ="شماره حساب:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3591
                    Top =840
                    Width =2511
                    Height =315
                    TabIndex =2
                    Name ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6254
                            Top =845
                            Width =450
                            Height =300
                            Name ="Label3"
                            Caption ="بانك:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5038
                    Top =108
                    Width =1071
                    Height =315
                    Name ="RADIF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6222
                            Top =105
                            Width =870
                            Height =300
                            Name ="Label64"
                            Caption ="رديف دفتر:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =1875
                    Top =4980
                    Width =1569
                    Height =394
                    FontWeight =700
                    TabIndex =9
                    Name ="Command44"
                    Caption ="&تصويب و خروج"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =206
                    TextFontCharSet =0
                    Left =3465
                    Top =4980
                    Width =1569
                    Height =394
                    FontWeight =700
                    TabIndex =10
                    Name ="Command45"
                    Caption ="&خـــــروج"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    Left =3591
                    Top =465
                    Width =2511
                    Height =315
                    TabIndex =1
                    Name ="N_SERI"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT N_SERI, N_S, N_KOL2, N_KOL3 FROM PAY_GETP WHERE (N_KOL3 IS NULL) AND (N_K"
                        "OL2 IS NULL) AND (N_S IS NULL OR N_S = 0)"
                    AfterUpdate ="[Event Procedure]"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6254
                            Top =465
                            Width =1080
                            Height =300
                            Name ="Label1"
                            Caption ="شماره سريال:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    Left =4605
                    Top =1190
                    Width =1506
                    Height =315
                    TabIndex =3
                    Name ="DATE_S"
                    RowSourceType ="Table/View/StoredProc"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6263
                            Top =1190
                            Width =1110
                            Height =300
                            Name ="Label7"
                            Caption ="تاريخ سررسيد:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3150
                    Top =1550
                    Width =2955
                    Height =315
                    TabIndex =4
                    Name ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6257
                            Top =1554
                            Width =945
                            Height =300
                            Name ="Label9"
                            Caption ="نام شعبه:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2805
                    Top =2616
                    Width =3300
                    Height =315
                    TabIndex =7
                    Name ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6257
                            Top =2615
                            Width =1350
                            Height =300
                            Name ="Label13"
                            Caption ="نام پرداخت كننده:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4599
                    Top =3661
                    Width =1506
                    Height =315
                    TabIndex =11
                    Name ="MOIN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6255
                            Top =3665
                            Width =1170
                            Height =300
                            Name ="Label47"
                            Caption ="از حساب معين:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4590
                    Top =3321
                    Width =1515
                    Height =315
                    TabIndex =12
                    Name ="KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6257
                            Top =3320
                            Width =1350
                            Height =300
                            Name ="Label49"
                            Caption ="ازحساب كل:"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4605
                    Top =4055
                    Width =1506
                    Height =315
                    TabIndex =13
                    Name ="TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =6255
                            Top =4055
                            Width =1455
                            Height =300
                            Name ="Label51"
                            Caption ="از حساب تفصيلي:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4365
                    Top =4457
                    Width =1746
                    Height =317
                    TabIndex =14
                    Name ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد  شخص\";2;\"عودت شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6270
                            Top =4457
                            Width =900
                            Height =240
                            Name ="Label28"
                            Caption ="وضعيت چك"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "BAKCHEKP.cls"
