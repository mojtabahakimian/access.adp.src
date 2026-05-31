Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =13209
    RowHeight =450
    ItemSuffix =13
    Left =1590
    Top =300
    Right =14805
    Bottom =8970
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =0
    RecSrcDt = Begin
        0x36ac1feff204e540
    End
    RecordSource ="TAKHFIF_DEF"
    Caption ="تعريف تخفيفات پيشرفته تر"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
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
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
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
            Height =8674
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11741
                    Top =165
                    Width =711
                    Height =360
                    ColumnWidth =630
                    Name ="TID"
                    ControlSource ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =12528
                            Top =165
                            Width =525
                            Height =315
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5605
                    Top =165
                    Width =5565
                    Height =360
                    ColumnWidth =5565
                    TabIndex =1
                    BackColor =10079487
                    Name ="TSHARH"
                    ControlSource ="TSHARH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =11244
                            Top =165
                            Width =420
                            Height =315
                            Name ="Label3"
                            Caption ="شرح"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7755
                    Top =675
                    Width =4365
                    Height =360
                    ColumnWidth =3360
                    TabIndex =2
                    Name ="TKIND"
                    ControlSource ="TKIND"
                    RowSourceType ="Value List"
                    RowSource ="1;\"براساس مبلغ پرداخت نقدي\";2;\"براساس مبلغ خالص فروش\""
                    ColumnWidths ="0.000"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =12226
                            Top =675
                            Width =795
                            Height =345
                            Name ="Label5"
                            Caption ="نوع تخفيف"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2040
                    Top =737
                    Width =1371
                    Height =270
                    TabIndex =5
                    Name ="TCREA"
                    ControlSource ="TCREA"
                    DefaultValue ="Now()"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextAlign =3
                            Left =3675
                            Top =735
                            Width =750
                            Height =270
                            Name ="Label7"
                            Caption ="تاريخ ايجاد"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2481
                    Top =165
                    Width =2355
                    Height =345
                    ColumnWidth =2355
                    TabIndex =6
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4925
                            Top =165
                            Width =615
                            Height =240
                            Name ="Label9"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =6803
                    Top =1640
                    Width =6236
                    Height =6570
                    TabIndex =4
                    Name ="TAKHFIF_DTL_F"
                    SourceObject ="Form.TAKHFIF_DTL_F"
                    LinkChildFields ="TID"
                    LinkMasterFields ="TID"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =300
                    Top =60
                    Width =1170
                    Height =465
                    FontWeight =700
                    TabIndex =3
                    ForeColor =0
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
                End
                Begin Subform
                    OverlapFlags =85
                    Left =225
                    Top =1645
                    Width =6236
                    Height =6570
                    TabIndex =7
                    Name ="CUSTKIND_TF_FORM"
                    SourceObject ="Form.CUSTKIND_TF_FORM"
                    LinkChildFields ="TID"
                    LinkMasterFields ="TID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1824
                            Top =1305
                            Width =2370
                            Height =270
                            Name ="Label12"
                            Caption ="گروه مشتريان مشمول اين تخفيف"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "TAKHFIF_DEF_FORM.cls"
