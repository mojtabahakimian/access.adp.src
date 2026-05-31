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
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4762
    RowHeight =315
    ItemSuffix =48
    Left =345
    Top =4530
    Right =11535
    Bottom =5970
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___Baseknow___ADA =[Forms]![Baseknow]![ADA],@Forms___Baseknow___ADV =[Form"
        "s]![Baseknow]![ADV],@Forms___Baseknow___BANKHA =[Forms]![Baseknow]![BANKHA]"
    RecSrcDt = Begin
        0xa2629c4f93f3e340
    End
    InputParameters ="@Forms___Baseknow___ADA =[Forms]![Baseknow]![ADA],@Forms___Baseknow___ADV =[Form"
        "s]![Baseknow]![ADV],@Forms___Baseknow___BANKHA =[Forms]![Baseknow]![BANKHA]"
    RecordSource ="dbo.CHEK_W_PASSPW"
    Caption ="سندهاي بستانكار اسناد دريافتني كه چك آنها در دفتر چك وصولي يا واگذاري يا برگشتي "
        "نخورده است"
    DatasheetFontName ="Tahoma"
    UnknownProp = {258 ,66 ,12 ,4 ,310 } Begin
        0x400046006f0072006d0073005f005f005f0042006100730065006b006e006f00 ,
        0x77005f005f005f0041004400410020003d005b0046006f0072006d0073005d00 ,
        0x21005b0042006100730065006b006e006f0077005d0021005b00410044004100 ,
        0x5d002c00400046006f0072006d0073005f005f005f0042006100730065006b00 ,
        0x6e006f0077005f005f005f0041004400560020003d005b0046006f0072006d00 ,
        0x73005d0021005b0042006100730065006b006e006f0077005d0021005b004100 ,
        0x440056005d002c00400046006f0072006d0073005f005f005f00420061007300 ,
        0x65006b006e006f0077005f005f005f00420041004e004b004800410020003d00 ,
        0x5b0046006f0072006d0073005d0021005b0042006100730065006b006e006f00 ,
        0x77005d0021005b00420041004e004b00480041005d00
    End
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
        Begin Section
            Height =8297
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =168
                    Top =113
                    Width =2466
                    ColumnOrder =1
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3500
                            Top =113
                            Width =1005
                            Height =240
                            Name ="Label1"
                            Caption ="شماره سريال"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =168
                    Top =793
                    Width =2466
                    ColumnWidth =5460
                    ColumnOrder =3
                    TabIndex =2
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3455
                            Top =793
                            Width =1050
                            Height =240
                            Name ="Label5"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =168
                    Top =1133
                    Width =2466
                    ColumnOrder =0
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3620
                            Top =1133
                            Width =885
                            Height =240
                            Name ="Label7"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =159
                    Top =1474
                    Width =2475
                    TabIndex =4
                    Name ="BES"
                    ControlSource ="BES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3785
                            Top =1474
                            Width =720
                            Height =240
                            Name ="Label9"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =168
                    Top =453
                    Width =2466
                    Height =255
                    ColumnOrder =2
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4130
                            Top =453
                            Width =375
                            Height =240
                            Name ="Label3"
                            Caption ="بانك"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FCHEK_W_PASSP.cls"
