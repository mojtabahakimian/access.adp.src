Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9334
    DatasheetFontHeight =10
    ItemSuffix =58
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0xa9075581d9ede340
    End
    RecordSource ="SELECT SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHKF) AS MEGHF, SUM(FROOSH_COUNTALL_CIT"
        "Y_N_FANI.MABK) AS MABKK, SUM(FROOSH_COUNTALL_CITY_N_FANI.MEGHkK) AS MEGHk, FROOS"
        "H_COUNTALL_CITY_N_FANI.SAIZE, FROOSH_COUNTALL_CITY_N_FANI.TARH, CAST(FROOSH_COUN"
        "TALL_CITY_N_FANI.TARH AS INT) AS TTARH, FROOSH_COUNTALL_CITY_N_FANI.CUST_NO, CUS"
        "T_HESAB.NAME FROM FROOSH_COUNTALL_CITY_N_FANI(1, 999999) FROOSH_COUNTALL_CITY_N_"
        "FANI INNER JOIN CUST_HESAB ON FROOSH_COUNTALL_CITY_N_FANI.CUST_NO = CUST_HESAB.h"
        "es GROUP BY FROOSH_COUNTALL_CITY_N_FANI.SAIZE, FROOSH_COUNTALL_CITY_N_FANI.TARH,"
        " CAST(FROOSH_COUNTALL_CITY_N_FANI.TARH AS INT), FROOSH_COUNTALL_CITY_N_FANI.CUST"
        "_NO, CUST_HESAB.NAME"
    Caption ="گزارش فروش به تفكيك سايز و طرح"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextAlign =1
            FontSize =12
            FontName ="Haettenschweiler"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CUST_NO"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SAIZE"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="MEGHk"
        End
        Begin PageHeader
            Height =771
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =2580
                    Width =3150
                    Height =465
                    FontWeight =700
                    Name ="Label48"
                    Caption ="گزارش فروش  به تفكيك سايز و طرح"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7432
                    Top =396
                    Width =1251
                    Height =375
                    FontSize =10
                    Name ="Text34"
                    ControlSource ="CUST_NO"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =2
                            Left =8683
                            Top =396
                            Width =330
                            Height =375
                            FontSize =10
                            Name ="Label35"
                            Caption ="كد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =2962
                    Top =396
                    Width =4086
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text36"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            TextFontFamily =2
                            Left =7070
                            Top =396
                            Width =315
                            Height =375
                            FontSize =10
                            Name ="Label37"
                            Caption ="نام:"
                            FontName ="Traffic"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader2"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =690
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =6870
                    Top =285
                    Width =1581
                    Height =375
                    FontSize =10
                    Name ="SAIZE"
                    ControlSource ="SAIZE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 1)"
                    ColumnWidths ="0"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =0
                            TextFontFamily =2
                            Left =8532
                            Top =285
                            Width =435
                            Height =375
                            FontSize =10
                            Name ="Label38"
                            Caption ="سايز:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =4815
                    Top =315
                    Width =495
                    Height =375
                    FontSize =10
                    Name ="Label51"
                    Caption ="كارتن"
                    FontName ="Traffic"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =3795
                    Top =315
                    Width =450
                    Height =375
                    FontSize =10
                    Name ="Label52"
                    Caption ="متراژ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =2040
                    Top =315
                    Width =450
                    Height =375
                    FontSize =10
                    Name ="Label53"
                    Caption ="مبلغ"
                    FontName ="Traffic"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =375
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3345
                    Width =1083
                    Height =375
                    FontSize =10
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4480
                    Width =963
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="MEGHF"
                    ControlSource ="MEGHF"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1365
                    Width =1923
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="MABKK"
                    ControlSource ="MABKK"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    FontName ="Traffic"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5550
                    Width =2091
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Combo39"
                    ControlSource ="ttarh"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 3)"
                    ColumnWidths ="0"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =0
                            TextFontFamily =2
                            Left =7677
                            Width =465
                            Height =375
                            FontSize =10
                            Name ="Label40"
                            Caption ="طرح:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =226
                    Width =573
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="MEGHkS"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1700
                    Width =633
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="MEGHFS"
                    ControlSource ="MEGHF"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =915
                    Width =678
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="MABKKS"
                    ControlSource ="MABKK"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =226
                    Width =573
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="MEGHkss"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1700
                    Width =633
                    Height =375
                    FontSize =10
                    TabIndex =8
                    Name ="MEGHFss"
                    ControlSource ="MEGHF"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =915
                    Width =678
                    Height =375
                    FontSize =10
                    TabIndex =9
                    Name ="MABKKss"
                    ControlSource ="MABKK"
                    FontName ="Traffic"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =375
            BreakLevel =1
            BackColor =16777164
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3345
                    Width =1083
                    Height =375
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=[MEGHkS]"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4480
                    Width =963
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text42"
                    ControlSource ="=[MEGHFS]"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1365
                    Width =1923
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text43"
                    ControlSource ="=[MABKKs]"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    FontName ="Traffic"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =5580
                    Width =1581
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Combo49"
                    ControlSource ="SAIZE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MPCODE, MPNAME FROM TCOD_MAP WHERE (MPP = 1)"
                    ColumnWidths ="0"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =0
                            TextFontFamily =2
                            Left =7245
                            Width =765
                            Height =375
                            FontSize =10
                            Name ="Label50"
                            Caption ="جمع سايز:"
                            FontName ="Traffic"
                        End
                    End
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =375
            BackColor =10092543
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4478
                    Width =1011
                    Height =375
                    FontSize =10
                    Name ="Text28"
                    ControlSource ="MEGHF"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1133
                    Width =2151
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text30"
                    ControlSource ="MABKK"
                    Format ="#,##0.00\" ريال \";-#,##0.00\" ريال \""
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3341
                    Width =1086
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text32"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =0
                    TextFontFamily =2
                    Left =5610
                    Width =765
                    Height =375
                    FontSize =10
                    Name ="Label54"
                    Caption ="جمع كل::"
                    FontName ="Traffic"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "list_froosh_cust_size_tarh.cls"
