Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =2
    GridY =10
    Width =7237
    RowHeight =450
    ItemSuffix =41
    Left =465
    Top =1815
    Right =15075
    Bottom =8730
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    RecSrcDt = Begin
        0xd64205c66caee240
    End
    ResyncCommand ="0"
    RecordSource ="DEED_DTL"
    Caption ="DEED_DTL"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
            Height =6122
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =480
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =11
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =1814
                    Height =255
                    ColumnWidth =1830
                    ColumnOrder =9
                    TabIndex =4
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,##0.00;-#,##0.00"
                    ValidationRule ="Not Is Null"
                    OnExit ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2154
                    Height =255
                    ColumnWidth =1830
                    ColumnOrder =10
                    TabIndex =5
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,##0.00;-#,##0.00"
                    ValidationRule ="Not Is Null"
                    OnExit ="[Event Procedure]"
                    OnKeyPress ="[Event Procedure]"
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2494
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =12
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =2834
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =13
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3174
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =14
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3514
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =15
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =180
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =16
                    TabIndex =11
                    Name ="RADIF"
                    ControlSource ="RADIF"
                    DefaultValue ="=[RecordsetClone].[RecordCount]+1"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1303
                    Width =5790
                    Height =255
                    ColumnWidth =10583
                    ColumnOrder =8
                    TabIndex =3
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHARH.SHARH FROM SHARH GROUP BY SHARH.SHARH ORDER BY SHARH.SHARH;"
                    AfterUpdate ="[Event Procedure]"
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
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
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
                    OnDblClick ="[Event Procedure]"
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
                    ReadingOrder =2
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
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
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
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1077
                    Top =5669
                    Width =4956
                    Height =300
                    ColumnWidth =1740
                    ColumnOrder =0
                    TabIndex =12
                    Name ="HES"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="1418;2835"
                    ValidationRule ="Not Is Null And <>\"\""
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =6127
                            Top =5671
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
                    Locked = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =283
                    Top =1870
                    Width =4956
                    Height =300
                    ColumnWidth =6132
                    ColumnOrder =7
                    TabIndex =13
                    Name ="Combo30"
                    ControlSource ="HES"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT hes, NAME FROM CUST_HESAB"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =5327
                            Top =1877
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
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3408
                    Top =3889
                    Height =255
                    ColumnWidth =0
                    ColumnOrder =17
                    TabIndex =14
                    Name ="ARZD"
                    ControlSource ="ARZD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6350
                            Top =3889
                            Width =630
                            Height =240
                            Name ="Label33"
                            Caption ="نرخ ارز"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2381
                    Top =4081
                    Width =1641
                    Height =255
                    ColumnWidth =1125
                    ColumnOrder =4
                    TabIndex =15
                    Name ="HES_T2"
                    ControlSource ="HES_T2"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4948
                            Top =4081
                            Width =945
                            Height =240
                            Name ="Label35"
                            Caption ="تفضيلي2"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2381
                    Top =4456
                    Width =1641
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =5
                    TabIndex =16
                    Name ="HES_T3"
                    ControlSource ="HES_T3"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4948
                            Top =4456
                            Width =945
                            Height =240
                            Name ="Label37"
                            Caption ="تفضيلي3"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2381
                    Top =4831
                    Width =1641
                    Height =255
                    ColumnWidth =285
                    ColumnOrder =6
                    TabIndex =17
                    Name ="HES_T4"
                    ControlSource ="HES_T4"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4948
                            Top =4831
                            Width =945
                            Height =240
                            Name ="Label39"
                            Caption ="تفضيلي4"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2400
                    Top =5220
                    Width =2481
                    ColumnWidth =3000
                    TabIndex =18
                    Name ="MHAZ_NO"
                    ControlSource ="MHAZ_NO"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MHAZ_NO, MHAZNAME FROM TCOD_MARKAZHAZ"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4932
                            Top =5220
                            Width =870
                            Height =240
                            Name ="Label40"
                            Caption ="مركز هزينه"
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
                    ReadingOrder =2
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
                    ReadingOrder =2
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
' See "DEED_DTL_SUB.cls"
