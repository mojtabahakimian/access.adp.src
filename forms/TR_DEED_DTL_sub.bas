Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =2
    GridY =10
    Width =7237
    RowHeight =450
    ItemSuffix =34
    Left =135
    Top =1305
    Right =15000
    Bottom =8100
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    OrderBy ="DEED_DTL.BED DESC"
    RecSrcDt = Begin
        0xd972559353d7e340
    End
    ResyncCommand ="0"
    RecordSource ="TR_DEED_DTL"
    Caption ="DEED_DTL"
    DatasheetFontName ="Tahoma"
    AllowFormView =0
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
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4265
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =480
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =8
                    Name ="N_S"
                    ControlSource ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6092
                            Top =480
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =1814
                    Height =255
                    ColumnWidth =1560
                    ColumnOrder =6
                    TabIndex =4
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,##0.00;-#,##0.00"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6455
                            Top =1814
                            Width =525
                            Height =240
                            Name ="Label11"
                            Caption ="بدهكار"
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
                    Left =3408
                    Top =2154
                    Height =255
                    ColumnWidth =1575
                    ColumnOrder =7
                    TabIndex =5
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,##0.00;-#,##0.00"
                    ValidationRule ="Not Is Null"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6335
                            Top =2154
                            Width =645
                            Height =240
                            Name ="Label13"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2494
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =9
                    TabIndex =6
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6020
                            Top =2494
                            Width =960
                            Height =240
                            Name ="Label15"
                            Caption ="شماره سري"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2834
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =10
                    TabIndex =7
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6605
                            Top =2834
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3174
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =11
                    TabIndex =8
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6020
                            Top =3174
                            Width =960
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3514
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =12
                    TabIndex =9
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6350
                            Top =3514
                            Width =630
                            Height =240
                            Name ="Label21"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =180
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =13
                    TabIndex =11
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="=[RecordsetClone].[RecordCount]+1"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1303
                    Width =5790
                    Height =255
                    ColumnWidth =5610
                    ColumnOrder =5
                    TabIndex =3
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHARH.SHARH FROM SHARH GROUP BY SHARH.SHARH ORDER BY SHARH.SHARH;"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6536
                            Top =1303
                            Width =420
                            Height =240
                            Name ="Label9"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =255
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3468
                    Top =1493
                    Width =1641
                    Height =255
                    ColumnWidth =735
                    ColumnOrder =3
                    TabIndex =10
                    Name ="HES_T"
                    ControlSource ="HES_T"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6035
                            Top =1493
                            Width =945
                            Height =240
                            Name ="Label28"
                            Caption ="تفضيلي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3468
                    Top =1133
                    Width =1641
                    Height =255
                    ColumnWidth =750
                    ColumnOrder =2
                    TabIndex =2
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6035
                            Top =1133
                            Width =945
                            Height =240
                            Name ="Label7"
                            Caption ="معين"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3468
                    Top =793
                    Width =1641
                    Height =255
                    ColumnWidth =720
                    ColumnOrder =1
                    TabIndex =1
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6200
                            Top =793
                            Width =780
                            Height =240
                            Name ="Label5"
                            Caption =" كل"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =56
                    Top =113
                    Width =4956
                    Height =300
                    ColumnWidth =1290
                    ColumnOrder =0
                    TabIndex =12
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDET"
                        "A_HES.NAME FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NU"
                        "MBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER "
                        "= DETA_HES.N_KOL"
                    ColumnWidths ="1418;2835"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextFontFamily =10
                            Left =5106
                            Top =115
                            Width =495
                            Height =285
                            Name ="Label4"
                            Caption ="حساب"
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =56
                    Top =533
                    Width =4956
                    Height =300
                    ColumnWidth =4110
                    ColumnOrder =4
                    TabIndex =13
                    Name ="Combo30"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMB"
                        "ER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, ISNU"
                        "LL(TOTA_HES.NAME, N' ') + N' - ' + ISNULL(TDETA_HES.NAME, N' ') AS Expr1 FROM TO"
                        "TA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HES.N"
                        "UMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KOL"
                    ColumnWidths ="0"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontFamily =10
                            Left =5100
                            Top =540
                            Width =720
                            Height =285
                            Name ="Label31"
                            Caption ="شرح حساب"
                            FontName ="Arabic Transparent"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =623
                    Top =2777
                    TabIndex =14
                    Name ="UP_DATE"
                    ControlSource ="UP_DATE"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3215
                            Top =2777
                            Width =810
                            Height =240
                            Name ="Label32"
                            Caption ="UP_DATE:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =453
                    Top =1757
                    TabIndex =15
                    Name ="UP_TIME"
                    ControlSource ="UP_TIME"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3075
                            Top =1757
                            Width =780
                            Height =240
                            Name ="Label33"
                            Caption ="UP_TIME:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =113
                    Height =255
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5314
                            Top =113
                            Width =525
                            Height =240
                            Name ="Label24"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =473
                    Height =255
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5194
                            Top =473
                            Width =645
                            Height =240
                            Name ="Label26"
                            Caption ="بستانكار"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TR_DEED_DTL_sub.cls"
