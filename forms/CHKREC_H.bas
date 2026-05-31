Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14796
    RowHeight =360
    ItemSuffix =14
    Top =15
    Right =14490
    Bottom =7890
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    OrderBy ="CHKREC_H.DATE DESC"
    RecSrcDt = Begin
        0xaab3f7787764e240
    End
    RecordSource ="CHKREC_H"
    Caption ="ثبت وصولي چكهاي دريافتي"
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin Section
            CanGrow = NotDefault
            Height =8390
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12389
                    Top =120
                    Width =1416
                    Height =300
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13921
                            Top =120
                            Width =750
                            Height =285
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6839
                    Top =453
                    Width =6945
                    Height =300
                    ColumnWidth =5175
                    TabIndex =1
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13917
                            Top =450
                            Width =780
                            Height =285
                            Name ="Label3"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =87
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =453
                    Top =529
                    Width =1641
                    Height =286
                    TabIndex =2
                    BackColor =32768
                    ForeColor =65535
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextFontCharSet =0
                            TextAlign =3
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =2103
                            Top =529
                            Width =1425
                            Height =285
                            BackColor =32768
                            ForeColor =65535
                            Name ="lsanad"
                            Caption ="ثبت در سند شماره:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =61
                    Top =908
                    Width =14565
                    Height =6900
                    TabIndex =3
                    Name ="CHRE_LST_SUB"
                    SourceObject ="Form.CHRE_LST_SUB"
                    LinkChildFields ="DATE"
                    LinkMasterFields ="DATE"
                    OnExit ="[Event Procedure]"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2891
                    Top =7993
                    Width =1983
                    Height =300
                    TabIndex =4
                    Name ="Text10"
                    ControlSource ="=[CHRE_LST_SUB].[Form]![smab]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4967
                            Top =7993
                            Width =780
                            Height =285
                            Name ="Label11"
                            Caption ="جمع مبالغ:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =141
                    TextFontCharSet =0
                    Left =598
                    Width =1020
                    Height =450
                    FontWeight =700
                    TabIndex =5
                    Name ="PRINTC"
                    Caption ="&چاپ برگه"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2404
                    Width =411
                    Height =345
                    TabIndex =6
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7137
                    Top =113
                    ColumnWidth =1417
                    TabIndex =7
                    Name ="IDH"
                    ControlSource ="IDH"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4726
                    Top =113
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =8
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "CHKREC_H.cls"
