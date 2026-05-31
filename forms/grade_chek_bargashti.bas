Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7088
    RowHeight =405
    ItemSuffix =45
    Left =825
    Top =7065
    Right =13545
    Bottom =9555
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xb4738682fdbae540
    End
    RecordSource ="SELECT N_SERI, vaz,BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, HES1, HE"
        "S2, HES3, VAZ FROM PAY_GETD WHERE (VAZ = 5 OR VAZ = 6) AND (CUST_NO = N'112-1-92"
        "')"
    Caption ="ليست حواله هاي بارگيري نشده"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
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
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =6066
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =510
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3640
                            Top =510
                            Width =1065
                            Height =240
                            Name ="Label32"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =1191
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3595
                            Top =1191
                            Width =1110
                            Height =240
                            Name ="Label34"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =1531
                    TabIndex =3
                    Name ="DATE"
                    ControlSource ="DATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3760
                            Top =1531
                            Width =945
                            Height =240
                            Name ="Label35"
                            Caption ="تاريخ دريافت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =1871
                    TabIndex =4
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3925
                            Top =1871
                            Width =780
                            Height =210
                            Name ="Label36"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =2211
                    TabIndex =5
                    Name ="MABL"
                    ControlSource ="MABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4270
                            Top =2211
                            Width =435
                            Height =240
                            Name ="Label37"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =2551
                    ColumnWidth =3855
                    TabIndex =6
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3475
                            Top =2551
                            Width =1230
                            Height =240
                            Name ="Label38"
                            Caption ="نام تحويل گيرنده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1303
                    Top =2892
                    TabIndex =7
                    Name ="N_HESAB"
                    ControlSource ="N_HESAB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3900
                            Top =2895
                            Width =1035
                            Height =240
                            Name ="Label39"
                            Caption ="شماره حساب"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1897
                    Top =3628
                    Width =2211
                    Height =315
                    ColumnWidth =1965
                    TabIndex =8
                    Name ="VAZ"
                    ControlSource ="VAZ"
                    RowSourceType ="Value List"
                    RowSource ="1;\"نزد صندوق\";2;\"نزد بانك\";3;\"وصول شده\";4;\"واگذار شده\";5;\"برگشت شده\";6"
                        ";\"مسترد شده\""
                    ColumnWidths ="0"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4239
                            Top =3628
                            Width =1005
                            Height =315
                            Name ="Label50"
                            Caption ="وضعيت چك*"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2415
                    Top =4095
                    ColumnWidth =3060
                    TabIndex =9
                    Name ="bhes"
                    ControlSource ="HES1"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =4980
                            Top =4095
                            Width =1005
                            Height =240
                            Name ="Label44"
                            Caption =" به حساب "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2415
                    Top =4775
                    ColumnWidth =3090
                    TabIndex =10
                    Name ="brhes"
                    ControlSource ="HES2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =4545
                            Top =4770
                            Width =1665
                            Height =405
                            Name ="Label45"
                            Caption ="برگشت به حساب "
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2415
                    Top =5455
                    TabIndex =11
                    Name ="vshes"
                    ControlSource ="HES3"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, hes + N' : ' + NAME AS Expr1 FROM CUST_HESAB"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =4845
                            Top =5460
                            Width =1395
                            Height =240
                            Name ="Label43"
                            Caption ="وصول به حساب "
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1303
                    Top =851
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_BANKS"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4300
                            Top =851
                            Width =405
                            Height =240
                            Name ="Label33"
                            Caption ="بانك"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "grade_chek_bargashti.cls"
