Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15307
    DatasheetFontHeight =10
    ItemSuffix =13
    Left =2160
    Top =780
    Right =16410
    Bottom =7245
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x30269a9e188ee540
    End
    RecordSource ="GRADE_FORMAT"
    Caption ="فرمت رده بندي مشتري"
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="WeblogmaYekan"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            Height =10091
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =13545
                    Top =226
                    Width =1026
                    Height =421
                    ColumnWidth =585
                    Name ="IDD"
                    ControlSource ="IDD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14651
                            Top =226
                            Width =525
                            Height =421
                            Name ="Label1"
                            Caption ="شماره"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5556
                    Top =226
                    Width =6795
                    Height =421
                    ColumnWidth =5535
                    TabIndex =1
                    BackColor =13434828
                    Name ="GFNAME"
                    ControlSource ="GFNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =12497
                            Top =226
                            Width =945
                            Height =421
                            Name ="Label3"
                            Caption ="عنوان فرمت "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3344
                    Top =226
                    Width =1590
                    Height =421
                    TabIndex =2
                    BackColor =13434828
                    Name ="GFDATE"
                    ControlSource ="GFDATE"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5025
                            Top =226
                            Width =435
                            Height =421
                            Name ="Label5"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =100
                    Top =793
                    Width =14025
                    Height =421
                    ColumnWidth =4455
                    TabIndex =3
                    BackColor =10092543
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14481
                            Top =800
                            Width =675
                            Height =421
                            Name ="Label7"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =113
                    Top =226
                    Width =2175
                    Height =421
                    ColumnWidth =2640
                    TabIndex =4
                    BackColor =10092543
                    Name ="USERNAME"
                    ControlSource ="USERNAME"
                    DefaultValue ="UCurrentUser()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2439
                            Top =226
                            Width =735
                            Height =421
                            Name ="Label9"
                            Caption ="نام کاربر"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =56
                    Top =1693
                    Width =15131
                    Height =8116
                    TabIndex =5
                    Name ="GRADE_TAB_FT_FORM"
                    SourceObject ="Form.GRADE_TAB_FT_FORM"
                    LinkChildFields ="GFID"
                    LinkMasterFields ="IDD"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7233
                    Top =1260
                    Width =1935
                    Height =345
                    TabIndex =6
                    BackColor =13408767
                    Name ="JAMZARIB"
                    ControlSource ="JAMZARIB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9207
                            Top =1264
                            Width =1215
                            Height =240
                            Name ="Label16"
                            Caption ="جمع ضرائب گروه"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3945
                    Top =1260
                    Width =1935
                    Height =345
                    TabIndex =7
                    BackColor =13408767
                    Name ="EMTIAZ"
                    ControlSource ="EMTIAZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5919
                            Top =1264
                            Width =675
                            Height =240
                            Name ="Label18"
                            Caption ="امتياز کل"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =120
                    Top =1260
                    Width =1695
                    Height =405
                    TabIndex =8
                    Name ="Command14"
                    Caption ="چاپ"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "GRADE_FORMAT_FORM.cls"
