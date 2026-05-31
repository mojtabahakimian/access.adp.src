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
    Top =2730
    Right =11520
    Bottom =4170
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___Baseknow___ADA =[Forms]![Baseknow]![ADA],@Forms___Baseknow___BANKHA =[F"
        "orms]![Baseknow]![BANKHA]"
    RecSrcDt = Begin
        0x23dcb37090f1e340
    End
    InputParameters ="@Forms___Baseknow___ADA =[Forms]![Baseknow]![ADA],@Forms___Baseknow___ADV =[Form"
        "s]![Baseknow]![ADV],@Forms___Baseknow___BANKHA =[Forms]![Baseknow]![BANKHA]"
    RecordSource ="CHEK_W_PASSP"
    Caption ="چكهاي دريافتي وصول شده يا واگذار شده يا برگشت شده كه سند بستانكار آن وجود ندارند"
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
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2825
                            Top =113
                            Width =1680
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
                    TabIndex =2
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2780
                            Top =793
                            Width =1725
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ سررسيد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =128
                    Top =1133
                    Width =2475
                    TabIndex =3
                    Name ="SHOBEH"
                    ControlSource ="SHOBEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3079
                            Top =1133
                            Width =1395
                            Height =240
                            Name ="Label9"
                            Caption ="نام شعبه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =137
                    Top =1473
                    Width =2466
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3424
                            Top =1473
                            Width =1050
                            Height =240
                            Name ="Label11"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1813
                    Width =2490
                    ColumnWidth =1980
                    TabIndex =5
                    Name ="NAME_TAH"
                    ControlSource ="NAME_TAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2719
                            Top =1813
                            Width =1755
                            Height =240
                            Name ="Label13"
                            Caption ="نام تحويل گيرنده"
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
                    TabIndex =1
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3455
                            Top =453
                            Width =1050
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
' See "FCHEK_W_PASS.cls"
