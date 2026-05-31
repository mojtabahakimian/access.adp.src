Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =220
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15320
    DatasheetFontHeight =10
    ItemSuffix =58
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xa4ec1a5a33efe540
    End
    RecordSource ="PM_JOB"
    Caption ="PM_JOB_REPORT"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
            BorderColor =8388608
        End
        Begin Line
            Width =1701
            BorderColor =8388608
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
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin ListBox
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BackStyle =0
            Width =1701
            LabelX =-1701
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
            GroupFooter = NotDefault
            ControlSource ="JO_PERSON"
        End
        Begin BreakLevel
            ControlSource ="JO_STDATE"
        End
        Begin BreakLevel
            ControlSource ="JO_LOCATION"
        End
        Begin PageHeader
            Height =1260
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =11175
                    Top =120
                    Width =4035
                    Height =615
                    FontSize =22
                    Name ="Label42"
                    Caption ="دستور کارWork Order"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Width =15320
                    BorderColor =12632256
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =3
                    Top =29
                    Width =15320
                    BorderColor =12632256
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =3
                    Top =852
                    Width =15320
                    BorderColor =12632256
                    Name ="Line47"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =13920
                    Top =915
                    Width =966
                    Height =300
                    FontSize =8
                    FontWeight =400
                    Name ="JO_STDATE_Label"
                    Caption ="تاريخ اجرا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4434
                    Top =915
                    Width =3060
                    Height =300
                    FontSize =8
                    FontWeight =400
                    Name ="JO_LOCATION_Label"
                    Caption ="محل دستگاه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =7524
                    Top =915
                    Width =2865
                    Height =300
                    FontSize =8
                    FontWeight =400
                    Name ="AS_IDD_Label"
                    Caption ="دستگاه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =10434
                    Top =915
                    Width =2790
                    Height =300
                    FontSize =8
                    FontWeight =400
                    Name ="JO_JOB_Label"
                    Caption ="کار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =13322
                    Top =915
                    Width =456
                    Height =300
                    FontSize =8
                    FontWeight =400
                    Name ="JO_STTIME_Label"
                    Caption ="ساعت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =54
                    Top =915
                    Width =840
                    Height =300
                    FontSize =8
                    FontWeight =400
                    Name ="JO_DONE_Label"
                    Caption ="وضعيت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1644
                    Top =915
                    Width =645
                    Height =285
                    FontSize =8
                    FontWeight =400
                    Name ="JO_ENDATE_Label"
                    Caption ="تاريخ انجام"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =999
                    Top =915
                    Width =405
                    Height =285
                    FontSize =8
                    FontWeight =400
                    Name ="JO_ENTIME_Label"
                    Caption ="ساعت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =2484
                    Top =915
                    Width =1290
                    Height =285
                    FontSize =8
                    FontWeight =400
                    Name ="JO_TAKHIRMO_Label"
                    Caption ="تاخير مجاز"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =3804
                    Top =915
                    Width =600
                    Height =285
                    FontSize =8
                    FontWeight =400
                    Name ="JO_TIMENEED_Label"
                    Caption ="نفر ساعت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =14925
                    Top =915
                    Width =375
                    Height =285
                    FontSize =8
                    FontWeight =400
                    Name ="Label55"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =60
                    Top =285
                    Width =3330
                    ForeColor =8388608
                    Name ="JO_PERSON"
                    ControlSource ="JO_PERSON"
                    RowSourceType ="Value List"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =178
                            TextFontFamily =2
                            Left =3469
                            Top =287
                            Width =630
                            Height =285
                            FontSize =8
                            FontWeight =400
                            Name ="JO_PERSON_Label"
                            Caption ="مجري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1260
                    Width =15320
                    BorderColor =12632256
                    Name ="Line57"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =342
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =13894
                    Top =57
                    Width =966
                    Height =271
                    Name ="JO_STDATE"
                    ControlSource ="JO_STDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =13328
                    Top =57
                    Width =456
                    Height =271
                    TabIndex =4
                    Name ="JO_STTIME"
                    ControlSource ="JO_STTIME"
                    FontName ="WeblogmaYekan"
                    InputMask ="##:##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1510
                    Top =57
                    Width =858
                    Height =271
                    TabIndex =6
                    Name ="JO_ENDATE"
                    ControlSource ="JO_ENDATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =894
                    Top =57
                    Width =570
                    Height =271
                    TabIndex =7
                    Name ="JO_ENTIME"
                    ControlSource ="JO_ENTIME"
                    FontName ="WeblogmaYekan"
                    InputMask ="##:##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3160
                    Top =57
                    Width =627
                    Height =285
                    TabIndex =8
                    Name ="JO_TAKHIRMO"
                    ControlSource ="JO_TAKHIRMO"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =3833
                    Top =57
                    Width =570
                    Height =271
                    TabIndex =10
                    Name ="JO_TIMENEED"
                    ControlSource ="JO_TIMENEED"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =14910
                    Top =57
                    Width =396
                    Height =271
                    TabIndex =11
                    Name ="rdf"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10465
                    Top =57
                    Width =2757
                    Height =271
                    TabIndex =3
                    Name ="JO_JOB"
                    ControlSource ="JO_JOB"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, JOB_TITLE FROM PM_JOBTITLE"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7521
                    Top =57
                    Width =2862
                    TabIndex =2
                    Name ="AS_IDD"
                    ControlSource ="AS_IDD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, AS_NAME FROM PM_ASSETS"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4470
                    Top =57
                    Width =3015
                    TabIndex =1
                    Name ="JO_LOCATION"
                    ControlSource ="JO_LOCATION"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT IDD, LO_NAME FROM PM_LOCATION"
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2424
                    Top =56
                    Width =684
                    TabIndex =9
                    Name ="JO_TAKHVAHED"
                    ControlSource ="JO_TAKHVAHED"
                    RowSourceType ="Value List"
                    RowSource ="1;\"ماه\";2;\"هفته\";3;\"سال\";4;\"روز\";5;\"ساعت\";6;\"کانتر\""
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    DecimalPlaces =0
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =45
                    Top =57
                    Width =789
                    TabIndex =5
                    Name ="JO_DONE"
                    ControlSource ="JO_DONE"
                    RowSourceType ="Value List"
                    RowSource ="0;\"انتظار\";1;\"انجام شد\";2;\"لغو شد\""
                    ColumnWidths ="0"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =0
            Name ="GroupFooter2"
        End
        Begin PageFooter
            Height =528
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =907
                    Top =170
                    Width =5040
                    Height =300
                    ForeColor =8388608
                    Name ="Text43"
                    ControlSource ="=Format(FARSIDATE(Now()),\"####/##/##\") & \"   \" & Format(Now(),\"HH:MM:SS\")"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =10215
                    Top =165
                    Width =5040
                    Height =300
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text44"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Width =15320
                    Height =30
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =28
                    Width =15320
                    Height =30
                    BorderColor =12632256
                    Name ="Line54"
                End
            End
        End
    End
End
CodeBehindForm
' See "PM_JOB_REPORT.cls"
