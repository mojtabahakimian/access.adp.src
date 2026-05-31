Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Cycle =1
    GridY =10
    Width =10885
    DatasheetFontHeight =10
    ItemSuffix =28
    Left =2190
    Top =1680
    Right =13320
    Bottom =7545
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xd7b32174b11ee440
    End
    RecordSource ="OTHER_DTL"
    Caption ="ساير اطلاعات"
    BeforeInsert ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
            Height =5640
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1281
                    Top =1080
                    Width =51
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =396
                    Top =737
                    Width =696
                    TabIndex =1
                    BackColor =128
                    Name ="TAG"
                    ControlSource ="TAG"
                    DefaultValue ="2"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7851
                    Top =114
                    Width =1215
                    Height =345
                    TabIndex =2
                    BackColor =13434828
                    Name ="REQUEST_NO"
                    ControlSource ="REQUEST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9159
                            Top =113
                            Width =1575
                            Height =240
                            Name ="Label5"
                            Caption ="شماره درخواست كالا:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4562
                    Top =162
                    Width =1980
                    Height =345
                    TabIndex =3
                    BackColor =13434828
                    Name ="BARNAMEH"
                    ControlSource ="BARNAMEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6667
                            Top =159
                            Width =1050
                            Height =240
                            Name ="Label7"
                            Caption ="شماره بارنامه:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7733
                    Top =690
                    Width =2025
                    Height =345
                    TabIndex =5
                    Name ="DRIVER_MOB"
                    ControlSource ="DRIVER_MOB"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9816
                            Top =745
                            Width =975
                            Height =240
                            Name ="Label11"
                            Caption ="موبايل راننده:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4217
                    Top =1260
                    Height =345
                    TabIndex =9
                    Name ="CAM_KHALY"
                    ControlSource ="CAM_KHALY"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6031
                            Top =1260
                            Width =1245
                            Height =240
                            Name ="Label15"
                            Caption ="وزن ماشين خالي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1274
                    Top =1260
                    Height =345
                    TabIndex =10
                    Name ="CAM_POOR"
                    ControlSource ="CAM_POOR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3140
                            Top =1260
                            Width =1005
                            Height =240
                            Name ="Label17"
                            Caption ="وزن ماشين پر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =164
                    Top =1793
                    Width =9945
                    Height =345
                    TabIndex =11
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    DefaultValue ="\".(بستن بار به عهده راننده وکشيدن چادر درفصل بارندگي الزامي ميباشد)اقلام فوق صح"
                        "يح و سالم تحويل اينجانب گرديد\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10200
                            Top =1793
                            Width =675
                            Height =240
                            Name ="Label21"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7401
                    Top =1245
                    Width =2601
                    Height =315
                    TabIndex =8
                    Name ="MAGHSAD"
                    ControlSource ="MAGHSAD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10056
                            Top =1290
                            Width =825
                            Height =240
                            Name ="Label19"
                            Caption ="*مقصد بار:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1274
                    Top =159
                    Width =2370
                    Height =345
                    TabIndex =4
                    Name ="DRIVER"
                    ControlSource ="DRIVER"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT DRIVER FROM dbo.OTHER_DTL GROUP BY DRIVER ORDER BY DRIVER"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3771
                            Top =158
                            Width =750
                            Height =240
                            Name ="Label9"
                            Caption ="نام راننده:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4236
                    Top =707
                    Width =2265
                    Height =300
                    TabIndex =6
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT CAMIUN_NUM FROM OTHER_DTL GROUP BY CAMIUN_NUM ORDER BY CA"
                        "MIUN_NUM"
                    ValidationRule ="Is Not Null"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6564
                            Top =717
                            Width =1110
                            Height =285
                            Name ="Label13"
                            Caption ="شماره ماشين:"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1274
                    Top =729
                    Width =1965
                    Height =300
                    TabIndex =7
                    Name ="CAMIUN"
                    ControlSource ="CAMIUN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOP 100 PERCENT CAMIUN FROM OTHER_DTL GROUP BY CAMIUN ORDER BY CAMIUN"
                    ValidationRule ="Is Not Null"
                    BeforeUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3294
                            Top =739
                            Width =885
                            Height =240
                            Name ="Label23"
                            Caption ="نوع ماشين:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Top =2791
                    Width =10785
                    Height =2790
                    TabIndex =13
                    Name ="OTHER_DTL_SUB_SUB"
                    SourceObject ="Form.OTHER_DTL_SUB_SUB"
                    LinkChildFields ="NUMBER;TAGG"
                    LinkMasterFields ="NUMBER;TAG"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =2
                            Left =9727
                            Top =2551
                            Width =420
                            Height =240
                            Name ="OTHER_DTL_SUB_SUB Label"
                            Caption ="توزين"
                            EventProcPrefix ="OTHER_DTL_SUB_SUB_Label"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7143
                    Top =2324
                    Width =2190
                    Height =405
                    TabIndex =12
                    Name ="Command26"
                    Caption ="دريافت نام كالاها  جهت توزين"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =4478
                    Top =2324
                    Width =2190
                    Height =405
                    TabIndex =14
                    Name ="Command27"
                    Caption ="انتقال وزن كالاها به حواله"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "OTHER_DTL_SUB.cls"
