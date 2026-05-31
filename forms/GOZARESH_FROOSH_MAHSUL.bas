Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =10
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14173
    DatasheetFontWeight =700
    ItemSuffix =46
    Left =2460
    Top =90
    Right =13935
    Bottom =9405
    DatasheetGridlinesColor =12632256
    Tag ="ALTER FUNCTION dbo.FROOSH_COUNT\015\012(@F_MENU_DATE_DT1 bigint,\015\012@F_MENU_"
        "DATE_DT2 bigint)\015\012RETURNS TABLE\015\012AS\015\012RETURN ( SELECT     TOP 1"
        "00 PERCENT COUNT(dbo.INVO_LST.CODE) AS TEDAD, dbo.STUF_DEF.CODE, dbo.STUF_DEF.NA"
        "ME, SUM(dbo.INVO_LST.MEGHk) AS SMEGHk, \015\012                      dbo.HEAD_LS"
        "T.DEPATMAN, dbo.DEPART.DEPNAME\015\012FROM         dbo.HEAD_LST INNER JOIN\015\012"
        "                      dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER "
        "AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN\015\012                      "
        "dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN\015\012        "
        "              dbo.DEPART ON dbo.HEAD_LST.DEPATMAN = dbo.DEPART.DEPATMAN\015\012W"
        "HERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.DATE_N BETWEEN @F_MENU_DATE_DT"
        "1 AND @F_MENU_DATE_DT2)\015\012GROUP BY dbo.STUF_DEF.NAME, dbo.HEAD_LST.DEPATMAN"
        ", dbo.DEPART.DEPNAME, dbo.STUF_DEF.CODE\015\012ORDER BY dbo.STUF_DEF.NAME )\015\012"
        "\015\012"
    RecSrcDt = Begin
        0xbe200f42eeafe240
    End
    RecordSource ="DEPART"
    Caption ="گزار ش فروش روزانه محصولات"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arabic Transparent"
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin FormHeader
            Height =623
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7259
                    Top =170
                    Width =5721
                    Height =255
                    FontWeight =700
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPATMAN, DEPNAME FROM DEPART"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =1
                            Left =13064
                            Top =170
                            Width =810
                            Height =240
                            FontWeight =700
                            Name ="Label34"
                            Caption =" دپارتمان:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4413
                    Top =110
                    Width =1577
                    Height =389
                    TabIndex =1
                    Name ="Command46"
                    Caption ="اطلاعات تخفيفات"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2836
                    Top =110
                    Width =1577
                    Height =389
                    TabIndex =2
                    Name ="Command47"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =9165
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =67
                    Top =45
                    Width =14055
                    Height =8235
                    Name ="FROOSH_COUNT_SUB"
                    SourceObject ="Form.FROOSH_COUNT_SUB"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4317
                    Top =8325
                    Width =1746
                    Height =345
                    FontWeight =700
                    TabIndex =1
                    Name ="Text41"
                    ControlSource ="=[FROOSH_COUNT_SUB].[Form].[SSMEGHK]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =6069
                            Top =8325
                            Width =570
                            Height =345
                            FontWeight =700
                            Name ="Label42"
                            Caption ="جمع كل"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2400
                    Top =8325
                    Height =345
                    FontWeight =700
                    TabIndex =2
                    Name ="Text43"
                    ControlSource ="=[FROOSH_COUNT_SUB].[Form].[STEDAD]"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =120
                    Top =8325
                    Width =2211
                    Height =345
                    FontWeight =700
                    TabIndex =3
                    Name ="Text45"
                    ControlSource ="=[FROOSH_COUNT_SUB].[Form].[Smabl]"
                    FontName ="Titr"
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
' See "GOZARESH_FROOSH_MAHSUL.cls"
