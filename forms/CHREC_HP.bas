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
    Width =14740
    RowHeight =345
    ItemSuffix =13
    Left =180
    Right =14925
    Bottom =7890
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    OrderBy ="CHREC_HP.DATE DESC"
    RecSrcDt = Begin
        0xd2f51cc18c64e240
    End
    RecordSource ="CHREC_HP"
    Caption ="ثبت وصولي چكهاي پرداختي"
    OnCurrent ="[Event Procedure]"
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
            Height =9297
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
                    Left =12250
                    Top =113
                    Width =1416
                    Height =300
                    Name ="DATE"
                    ControlSource ="DATE"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13782
                            Top =113
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
                    Left =6809
                    Top =513
                    Width =6855
                    Height =300
                    ColumnWidth =4380
                    TabIndex =1
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =13797
                            Top =510
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =396
                    Top =525
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
                            Left =2046
                            Top =525
                            Width =1395
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
                    Width =14520
                    Height =7755
                    TabIndex =3
                    Name ="CHRE_LSP_SUB"
                    SourceObject ="Form.CHRE_LSP_SUB"
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
                    Left =2437
                    Top =8844
                    Width =1983
                    Height =300
                    TabIndex =4
                    Name ="Text10"
                    ControlSource ="=[CHRE_LSP_SUB].[Form]![smab]"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4513
                            Top =8844
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
                    Left =613
                    Top =56
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
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2426
                    Top =170
                    TabIndex =6
                    Name ="IDH"
                    ControlSource ="IDH"
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
                    Left =7659
                    Top =56
                    Width =411
                    Height =345
                    FontWeight =700
                    TabIndex =7
                    Name ="LETSANAD"
                    DefaultValue ="False"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =4896
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
' See "CHREC_HP.cls"
