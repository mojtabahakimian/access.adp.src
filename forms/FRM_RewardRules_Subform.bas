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
    Width =8726
    RowHeight =390
    ItemSuffix =32
    Left =525
    Top =10170
    Right =14820
    Bottom =12330
    DatasheetBackColor =10079487
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0x8232ef855c5de640
    End
    RecordSource ="RewardRules"
    Caption ="جايزه"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            Height =7535
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =737
                    ColumnOrder =0
                    Name ="RuleID"
                    ControlSource ="RuleID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =737
                            Width =555
                            Height =240
                            Name ="شماره"
                            Caption ="RuleID"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6236
                    Top =1417
                    ColumnWidth =1110
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Quantity_Threshold"
                    ControlSource ="Quantity_Threshold"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =1417
                            Width =1500
                            Height =240
                            Name ="Label5"
                            Caption ="به ازاي هر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6236
                    Top =2438
                    ColumnWidth =1110
                    ColumnOrder =2
                    TabIndex =3
                    Name ="Reward_Quantity"
                    ControlSource ="Reward_Quantity"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =2438
                            Width =1350
                            Height =240
                            Name ="Label11"
                            Caption ="تعداد جايزه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =2778
                    ColumnWidth =1035
                    ColumnOrder =6
                    TabIndex =4
                    Name ="Reward_Discount_Percentage"
                    ControlSource ="Reward_Discount_Percentage"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4422
                            Top =2778
                            Width =2250
                            Height =240
                            Name ="Label13"
                            Caption ="درصد تخفيف"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =3458
                    ColumnOrder =7
                    TabIndex =5
                    Name ="StartDate"
                    ControlSource ="StartDate"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =3458
                            Width =780
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ شروع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =3798
                    ColumnOrder =8
                    TabIndex =6
                    Name ="EndDate"
                    ControlSource ="EndDate"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =3798
                            Width =690
                            Height =240
                            Name ="Label19"
                            Caption ="تاريخ پايان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =4138
                    Width =2490
                    ColumnOrder =9
                    TabIndex =7
                    Name ="Description"
                    ControlSource ="Description"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =4138
                            Width =870
                            Height =240
                            Name ="Label21"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =4479
                    ColumnWidth =2070
                    ColumnOrder =10
                    TabIndex =8
                    Name ="CRT"
                    ControlSource ="CRT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =4479
                            Width =375
                            Height =240
                            Name ="Label23"
                            Caption ="ايجاد"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6236
                    Top =4819
                    ColumnOrder =11
                    TabIndex =9
                    Name ="UID"
                    ControlSource ="UID"
                    DefaultValue ="=UCurrentUserIDD()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =4819
                            Width =345
                            Height =240
                            Name ="Label25"
                            Caption ="کاربر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Width =1680
                    Height =255
                    ColumnWidth =5280
                    ColumnOrder =12
                    TabIndex =10
                    Name ="ProductID_Target"
                    ControlSource ="ProductID_Target"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;5103;567"
                    ValidationRule ="Is Not Null"
                    Tag ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =3041
                            Width =510
                            Height =240
                            Name ="Label27"
                            Caption =" كالا"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6236
                    Top =1757
                    Width =2490
                    ColumnWidth =1290
                    ColumnOrder =5
                    TabIndex =2
                    Name ="Reward_Type"
                    ControlSource ="Reward_Type"
                    RowSourceType ="Value List"
                    RowSource ="\"Product\";\"محصول\";\"Discount\";\"تخفيف\""
                    ColumnWidths ="0"
                    DefaultValue ="\"Product\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4422
                            Top =1757
                            Width =1080
                            Height =240
                            Name ="Label7"
                            Caption ="نوع جايزه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Top =375
                    Width =1680
                    Height =255
                    ColumnWidth =6630
                    ColumnOrder =3
                    TabIndex =11
                    Name ="Reward_ProductID"
                    ControlSource ="Reward_ProductID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_D"
                        "EF.NAME"
                    ColumnWidths ="0;5103;567"
                    ValidationRule ="Is Not Null"
                    DefaultValue ="=[Forms]![stuf_def]![code]"
                    Tag ="1"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =3041
                            Top =375
                            Width =510
                            Height =240
                            Name ="Label29"
                            Caption ="كالاي جايزه"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2664
                    Top =2721
                    ColumnOrder =4
                    TabIndex =12
                    Name ="IsActive"
                    ControlSource ="IsActive"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2894
                            Top =2691
                            Width =690
                            Height =240
                            Name ="Label31"
                            Caption ="فعال"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FRM_RewardRules_Subform.cls"
