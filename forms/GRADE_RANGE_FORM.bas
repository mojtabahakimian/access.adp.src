Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15261
    RowHeight =315
    ItemSuffix =20
    Left =1605
    Right =16860
    Bottom =7710
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x35784bb21190e540
    End
    RecordSource ="GRADE_RANGE"
    Caption ="تعريف رنج (کلاس) مشتري"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        Begin FormHeader
            SpecialEffect =1
            Height =751
            BackColor =13434828
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =6473
                    Top =76
                    Width =3540
                    Height =675
                    FontSize =20
                    ForeColor =255
                    Name ="Label4"
                    Caption ="تعريف رنج (کلاس) مشتري"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =6465
                    Top =60
                    Width =3540
                    Height =675
                    FontSize =20
                    ForeColor =255
                    Name ="Label5"
                    Caption ="تعريف رنج (کلاس) مشتري"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            SpecialEffect =1
            Height =8163
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =180
                    Top =1815
                    Width =14853
                    Height =6232
                    Name ="GRADE_SHART_FORM"
                    SourceObject ="Form.GRADE_SHART_FORM"
                    LinkChildFields ="GID"
                    LinkMasterFields ="GID"
                    OnExit ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =225
                    Top =1020
                    Width =1695
                    Height =405
                    TabIndex =1
                    Name ="Command14"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12907
                    Top =120
                    Width =756
                    Height =340
                    ColumnWidth =630
                    TabIndex =2
                    Name ="GID"
                    ControlSource ="GID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =13718
                            Top =120
                            Width =900
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7020
                    Top =517
                    Width =6675
                    Height =340
                    ColumnWidth =4800
                    TabIndex =3
                    Name ="GNAME"
                    ControlSource ="GNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =13711
                            Top =517
                            Width =1155
                            Height =240
                            Name ="Label3"
                            Caption ="نام (عنوان)"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =12697
                    Top =930
                    Width =966
                    Height =340
                    TabIndex =4
                    BackColor =10079487
                    Name ="GAZ"
                    ControlSource ="GAZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =13718
                            Top =930
                            Width =930
                            Height =240
                            Name ="Label19"
                            Caption ="از امتياز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10155
                    Top =915
                    Width =966
                    Height =340
                    TabIndex =5
                    BackColor =10079487
                    Name ="GTA"
                    ControlSource ="GTA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =11176
                            Top =915
                            Width =930
                            Height =240
                            Name ="Label7"
                            Caption ="تا امتياز"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2610
                    Top =1365
                    Width =11085
                    Height =340
                    TabIndex =6
                    Name ="GTOZIH"
                    ControlSource ="GTOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =13711
                            Top =1365
                            Width =1200
                            Height =240
                            Name ="Label9"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =240
                    Top =510
                    Width =1695
                    Height =405
                    FontWeight =700
                    TabIndex =7
                    Name ="ESLAH"
                    Caption ="اصلاح "
                    OnClick ="[Event Procedure]"
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
' See "GRADE_RANGE_FORM.cls"
