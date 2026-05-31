Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5040
    RowHeight =330
    ItemSuffix =17
    Left =765
    Top =1710
    Right =14745
    Bottom =4125
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xbb3e4b60ee1ee440
    End
    RecordSource ="VISITOR_DTL"
    Caption ="پورسانت"
    BeforeInsert ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3288
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnWidth =1417
                    ColumnOrder =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =690
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =210
                    ColumnOrder =2
                    TabIndex =1
                    Name ="TAG"
                    ControlSource ="TAG"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    Height =255
                    ColumnWidth =825
                    ColumnOrder =3
                    TabIndex =2
                    Name ="DARSAD"
                    ControlSource ="DARSAD"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =690
                            Height =240
                            Name ="Label7"
                            Caption ="*درصد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    ColumnWidth =1200
                    ColumnOrder =4
                    TabIndex =3
                    Name ="PURSANT"
                    ControlSource ="PURSANT"
                    Format ="#,##0.00;-#,##0.00"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =1470
                            Width =1095
                            Height =240
                            Name ="Label9"
                            Caption ="*مبلغ پورسانت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1814
                    Width =2490
                    ColumnWidth =2565
                    ColumnOrder =7
                    TabIndex =4
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1814
                            Width =540
                            Height =240
                            Name ="Label11"
                            Caption ="توضيح"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =675
                    Top =737
                    Width =3546
                    Height =315
                    ColumnWidth =5325
                    ColumnOrder =0
                    TabIndex =5
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     hes, NAME AS nam, hes AS Expr1\015\012FROM         CUST_HESAB"
                    ColumnWidths ="0;3969"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4305
                            Top =737
                            Width =735
                            Height =300
                            ForeColor =8388608
                            Name ="Label19"
                            Caption =" شخص"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =1157
                    Top =2381
                    ColumnWidth =810
                    ColumnOrder =5
                    TabIndex =6
                    Name ="STAT"
                    ControlSource ="STAT"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =495
                            Top =2355
                            Width =705
                            Height =240
                            Name ="Label13"
                            Caption ="مبلغ ثابت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =226
                    Top =2721
                    ColumnWidth =6720
                    ColumnOrder =6
                    TabIndex =7
                    Name ="PORID"
                    ControlSource ="PORID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT VISITORS_PORSANT.PORID, CAST(VISITORS_PORSANT.PORID AS nvarchar) + N' - '"
                        " + CAST(VISITORS_PORSANT.VDATE AS nvarchar) + N' - ' + ISNULL(CUSTKIND.CUSTKNAME"
                        ", N'بدون گروه (همه)') + N' - ' + ISNULL(VISITORS_PORSANT.COMMENT, N' ') + N' - '"
                        " + CUST_HESAB.NAME AS Expr1 FROM VISITORS_PORSANT INNER JOIN CUST_HESAB ON VISIT"
                        "ORS_PORSANT.HES = CUST_HESAB.hes LEFT OUTER JOIN CUSTKIND ON VISITORS_PORSANT.CU"
                        "ST_COD = CUSTKIND.CUST_COD"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3015
                            Top =2715
                            Width =1650
                            Height =240
                            Name ="Label16"
                            Caption ="الگوي پرداخت پورسانت"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =240
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Name ="SCMAB"
                    ControlSource ="=Sum([PURSANT])"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3197
                            Width =375
                            Height =240
                            Name ="Label43"
                            Caption ="مبلغ"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "VISITOR_DTL_SUB.cls"
