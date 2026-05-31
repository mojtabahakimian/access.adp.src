Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14286
    DatasheetFontHeight =10
    ItemSuffix =10
    Left =2280
    Right =16560
    Bottom =8805
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3eafee1c9318e440
    End
    Caption ="آمار كل"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
            CanGrow = NotDefault
            Height =9694
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =8555
                    Top =509
                    Width =5550
                    Height =3795
                    Name ="task_count_daysFORM_SUB"
                    SourceObject ="Form.task_count_daysFORM_SUB"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =8555
                    Top =4424
                    Width =5550
                    Height =5040
                    TabIndex =1
                    Name ="EVENT_COUNT_SUB"
                    SourceObject ="Form.EVENT_COUNT_SUB"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6519
                    Top =170
                    Width =1023
                    Height =228
                    TabIndex =2
                    Name ="Text3"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7608
                            Top =170
                            Width =555
                            Height =240
                            Name ="Label4"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =510
                    Top =566
                    Width =2373
                    Height =765
                    FontSize =20
                    FontWeight =700
                    TabIndex =3
                    BackColor =13434828
                    Name ="Text5"
                    ControlSource ="=[task_count_daysFORM_SUB].[Form]![sc]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =2940
                            Top =570
                            Width =3810
                            Height =675
                            FontSize =18
                            FontWeight =700
                            Name ="Label6"
                            Caption ="تعداد كل ورودي هاي ثبت شد"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =512
                    Top =1427
                    Width =2373
                    Height =765
                    FontSize =20
                    FontWeight =700
                    TabIndex =4
                    BackColor =16764057
                    Name ="Text7"
                    ControlSource ="=[EVENT_COUNT_SUB].[Form]![sc]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =2944
                            Top =1427
                            Width =3990
                            Height =675
                            FontSize =18
                            FontWeight =700
                            Name ="Label8"
                            Caption ="تعداد كل روي داد هاي ثبت شد"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =120
                    Top =2265
                    Width =8220
                    Height =7200
                    TabIndex =5
                    Name ="task_count_daysFORM_SUB_cust"
                    SourceObject ="Form.task_count_daysFORM_SUB_cust"
                End
            End
        End
    End
End
CodeBehindForm
' See "AMAR_KOL.cls"
