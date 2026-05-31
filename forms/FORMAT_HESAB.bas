Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7540
    DatasheetFontHeight =10
    ItemSuffix =352
    Left =4710
    Top =2235
    Right =12255
    Bottom =7995
    HelpContextId =8017
    DatasheetGridlinesColor =12632256
    Filter ="((BANK=\"21\"))"
    RecSrcDt = Begin
        0xac66b04c906de240
    End
    RecordSource ="FORMAT_HESAB"
    Caption ="فرمت چاپي به حساب گذاشتن"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
            Height =5782
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    DecimalPlaces =0
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3830
                    Top =120
                    Width =2490
                    Height =300
                    Name ="BANK"
                    ControlSource ="BANK"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM TCOD_BANKS;"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =6408
                            Top =120
                            Width =435
                            Height =300
                            Name ="Label1"
                            Caption =" بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =886
                    Width =816
                    Height =300
                    TabIndex =2
                    Name ="DATE_TOP"
                    ControlSource ="DATE_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =886
                    Width =861
                    Height =300
                    TabIndex =3
                    Name ="DATE_LEFT"
                    ControlSource ="DATE_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =886
                    Width =681
                    Height =300
                    TabIndex =5
                    Name ="DATE_SIZE"
                    ControlSource ="DATE_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =886
                    TabIndex =7
                    Name ="DATE_P"
                    ControlSource ="DATE_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =886
                    Width =1455
                    Height =300
                    TabIndex =4
                    Name ="DATE_FONT"
                    ControlSource ="DATE_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =900
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label81"
                    Caption ="تاريخ روز:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =886
                    Width =666
                    Height =300
                    TabIndex =6
                    Name ="DATE_WITH"
                    ControlSource ="DATE_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =1335
                    Width =816
                    Height =300
                    TabIndex =8
                    Name ="SHO_TOP"
                    ControlSource ="SHO_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =1335
                    Width =861
                    Height =300
                    TabIndex =9
                    Name ="SHO_LEFT"
                    ControlSource ="SHO_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =1335
                    Width =681
                    Height =300
                    TabIndex =11
                    Name ="SHO_SIZE"
                    ControlSource ="SHO_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =1335
                    TabIndex =13
                    Name ="SHO_P"
                    ControlSource ="SHO_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =1335
                    Width =1455
                    Height =300
                    TabIndex =10
                    Name ="SHO_FONT"
                    ControlSource ="SHO_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =1349
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label218"
                    Caption ="شعبه بانك:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =1335
                    Width =666
                    Height =300
                    TabIndex =12
                    Name ="SHO_WITH"
                    ControlSource ="SHO_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =1784
                    Width =816
                    Height =300
                    TabIndex =14
                    Name ="SHES_TOP"
                    ControlSource ="SHES_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =1784
                    Width =861
                    Height =300
                    TabIndex =15
                    Name ="SHES_LEFT"
                    ControlSource ="SHES_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =1784
                    Width =681
                    Height =300
                    TabIndex =17
                    Name ="SHES_SIZE"
                    ControlSource ="SHES_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =1784
                    TabIndex =19
                    Name ="SHES_P"
                    ControlSource ="SHES_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =1784
                    Width =1455
                    Height =300
                    TabIndex =16
                    Name ="SHES_FONT"
                    ControlSource ="SHES_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =1798
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label231"
                    Caption ="شماره حساب:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =1784
                    Width =666
                    Height =300
                    TabIndex =18
                    Name ="SHES_WITH"
                    ControlSource ="SHES_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =2232
                    Width =816
                    Height =300
                    TabIndex =20
                    Name ="NAME_TOP"
                    ControlSource ="NAME_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =2232
                    Width =861
                    Height =300
                    TabIndex =21
                    Name ="NAME_LEFT"
                    ControlSource ="NAME_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =2232
                    Width =681
                    Height =300
                    TabIndex =23
                    Name ="NAME_SIZE"
                    ControlSource ="NAME_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =2232
                    TabIndex =25
                    Name ="NAME_P"
                    ControlSource ="NAME_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =2232
                    Width =1455
                    Height =300
                    TabIndex =22
                    Name ="NAME_FONT"
                    ControlSource ="NAME_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =2250
                    Width =1380
                    Height =300
                    BackColor =12632256
                    Name ="Label244"
                    Caption ="نام صاحب حساب:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =2232
                    Width =666
                    Height =300
                    TabIndex =24
                    Name ="NAME_WITH"
                    ControlSource ="NAME_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =2681
                    Width =816
                    Height =300
                    TabIndex =26
                    Name ="CDATE_TOP"
                    ControlSource ="CDATE_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =2681
                    Width =861
                    Height =300
                    TabIndex =27
                    Name ="CDATE_LEFT"
                    ControlSource ="CDATE_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =2681
                    Width =681
                    Height =300
                    TabIndex =29
                    Name ="CDATE_SIZE"
                    ControlSource ="CDATE_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =2681
                    TabIndex =31
                    Name ="CDATE_P"
                    ControlSource ="CDATE_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =2681
                    Width =1455
                    Height =300
                    TabIndex =28
                    Name ="CDATE_FONT"
                    ControlSource ="CDATE_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =2696
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label257"
                    Caption ="تاريخ چك:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =2681
                    Width =666
                    Height =300
                    TabIndex =30
                    Name ="CDATE_WITH"
                    ControlSource ="CDATE_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =3130
                    Width =816
                    Height =300
                    TabIndex =32
                    Name ="SER_TOP"
                    ControlSource ="SER_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =3130
                    Width =861
                    Height =300
                    TabIndex =33
                    Name ="SER_LEFT"
                    ControlSource ="SER_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =3130
                    Width =681
                    Height =300
                    TabIndex =35
                    Name ="SER_SIZE"
                    ControlSource ="SER_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =3130
                    TabIndex =37
                    Name ="SER_P"
                    ControlSource ="SER_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =3130
                    Width =1455
                    Height =300
                    TabIndex =34
                    Name ="SER_FONT"
                    ControlSource ="SER_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =3145
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label270"
                    Caption ="شماره چك:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =3130
                    Width =666
                    Height =300
                    TabIndex =36
                    Name ="SER_WITH"
                    ControlSource ="SER_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =3579
                    Width =816
                    Height =300
                    TabIndex =38
                    Name ="CHES_TOP"
                    ControlSource ="CHES_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =3579
                    Width =861
                    Height =300
                    TabIndex =39
                    Name ="CHES_LEFT"
                    ControlSource ="CHES_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =3579
                    Width =681
                    Height =300
                    TabIndex =41
                    Name ="CHES_SIZE"
                    ControlSource ="CHES_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =3579
                    TabIndex =43
                    Name ="CHES_P"
                    ControlSource ="CHES_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =3579
                    Width =1455
                    Height =300
                    TabIndex =40
                    Name ="CHES_FONT"
                    ControlSource ="CHES_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =3600
                    Width =1395
                    Height =285
                    BackColor =12632256
                    Name ="Label283"
                    Caption ="شماره حساب چك:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =3579
                    Width =666
                    Height =300
                    TabIndex =42
                    Name ="CHES_WITH"
                    ControlSource ="CHES_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =4028
                    Width =816
                    Height =300
                    TabIndex =44
                    Name ="BSHO_TOP"
                    ControlSource ="BSHO_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =4028
                    Width =861
                    Height =300
                    TabIndex =45
                    Name ="BSHO_LEFT"
                    ControlSource ="BSHO_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =4028
                    Width =681
                    Height =300
                    TabIndex =47
                    Name ="BSHO_SIZE"
                    ControlSource ="BSHO_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =4028
                    TabIndex =49
                    Name ="BSHO_P"
                    ControlSource ="BSHO_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =4028
                    Width =1455
                    Height =300
                    TabIndex =46
                    Name ="BSHO_FONT"
                    ControlSource ="BSHO_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =4043
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label296"
                    Caption ="بانك / شعبه:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =4028
                    Width =666
                    Height =300
                    TabIndex =48
                    Name ="BSHO_WITH"
                    ControlSource ="BSHO_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =4478
                    Width =816
                    Height =300
                    TabIndex =50
                    Name ="MAB_TOP"
                    ControlSource ="MAB_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =4478
                    Width =861
                    Height =300
                    TabIndex =51
                    Name ="MAB_LEFT"
                    ControlSource ="MAB_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =4478
                    Width =681
                    Height =300
                    TabIndex =53
                    Name ="MAB_SIZE"
                    ControlSource ="MAB_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =4478
                    TabIndex =55
                    Name ="MAB_P"
                    ControlSource ="MAB_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =4478
                    Width =1455
                    Height =300
                    TabIndex =52
                    Name ="MAB_FONT"
                    ControlSource ="MAB_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =4492
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label309"
                    Caption ="مبلغ:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =4478
                    Width =666
                    Height =300
                    TabIndex =54
                    Name ="MAB_WITH"
                    ControlSource ="MAB_WITH"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4976
                    Top =4927
                    Width =816
                    Height =300
                    TabIndex =56
                    Name ="AMAB_TOP"
                    ControlSource ="AMAB_TOP"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4016
                    Top =4927
                    Width =861
                    Height =300
                    TabIndex =57
                    Name ="AMAB_LEFT"
                    ControlSource ="AMAB_LEFT"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1706
                    Top =4927
                    Width =681
                    Height =300
                    TabIndex =59
                    Name ="AMAB_SIZE"
                    ControlSource ="AMAB_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =93
                    ReadingOrder =1
                    Left =416
                    Top =4927
                    TabIndex =61
                    Name ="AMAB_P"
                    ControlSource ="AMAB_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2471
                    Top =4927
                    Width =1455
                    Height =300
                    TabIndex =58
                    Name ="AMAB_FONT"
                    ControlSource ="AMAB_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6086
                    Top =4941
                    Width =1185
                    Height =285
                    BackColor =12632256
                    Name ="Label322"
                    Caption ="مبلغ به حروف:"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =971
                    Top =4927
                    Width =666
                    Height =300
                    TabIndex =60
                    Name ="AMAB_WITH"
                    ControlSource ="AMAB_WITH"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =4961
                    Top =540
                    Width =840
                    Height =285
                    Name ="Label23"
                    Caption ="فاصله از بالا"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =4001
                    Top =540
                    Width =870
                    Height =285
                    Name ="Label25"
                    Caption ="فاصله از چپ"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =2966
                    Top =540
                    Width =660
                    Height =285
                    Name ="Label27"
                    Caption ="نوع قلم"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1691
                    Top =540
                    Width =690
                    Height =285
                    Name ="Label29"
                    Caption ="اندازه قلم"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =956
                    Top =540
                    Width =660
                    Height =285
                    Name ="Label90"
                    Caption ="طول فيلد"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =191
                    Top =540
                    Width =660
                    Height =285
                    Name ="Label39"
                    Caption ="چاپ شود"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =131
                    Top =495
                    Width =7393
                    Height =5198
                    Name ="Box325"
                End
                Begin Line
                    OverlapFlags =127
                    SpecialEffect =5
                    Left =5850
                    Top =510
                    Width =0
                    Height =5205
                    Name ="Line326"
                End
                Begin Line
                    OverlapFlags =127
                    SpecialEffect =5
                    Left =4932
                    Top =510
                    Width =0
                    Height =5205
                    Name ="Line327"
                End
                Begin Line
                    OverlapFlags =127
                    SpecialEffect =5
                    Left =3971
                    Top =510
                    Width =0
                    Height =5205
                    Name ="Line328"
                End
                Begin Line
                    OverlapFlags =127
                    SpecialEffect =5
                    Left =2426
                    Top =510
                    Width =0
                    Height =5205
                    Name ="Line329"
                End
                Begin Line
                    OverlapFlags =127
                    SpecialEffect =5
                    Left =1661
                    Top =510
                    Width =0
                    Height =5205
                    Name ="Line330"
                End
                Begin Line
                    OverlapFlags =127
                    SpecialEffect =5
                    Left =881
                    Top =510
                    Width =0
                    Height =5205
                    Name ="Line331"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =840
                    Width =7325
                    Name ="Line332"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =1260
                    Width =7325
                    Name ="Line333"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =1710
                    Width =7325
                    Name ="Line334"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =2160
                    Width =7325
                    Name ="Line335"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =2610
                    Width =7325
                    Name ="Line336"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =3060
                    Width =7325
                    Name ="Line337"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =3510
                    Width =7325
                    Name ="Line338"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =3960
                    Width =7325
                    Name ="Line339"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =4410
                    Width =7325
                    Name ="Line340"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =116
                    Top =4860
                    Width =7325
                    Name ="Line341"
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6189
                    Top =510
                    Width =840
                    Height =285
                    Name ="Label342"
                    Caption ="نام فيلد"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =170
                    Top =56
                    Width =1245
                    Height =340
                    FontWeight =700
                    TabIndex =68
                    Name ="Command88"
                    Caption ="اعمال &تغيرات"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1820
                    Top =113
                    Width =846
                    Height =315
                    TabIndex =1
                    Name ="KIND"
                    ControlSource ="KIND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2717
                            Top =113
                            Width =1050
                            Height =300
                            Name ="Label343"
                            Caption ="شماره فرمت:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4980
                    Top =5332
                    Width =816
                    Height =300
                    TabIndex =62
                    Name ="CITY_TOP"
                    ControlSource ="CITY_TOP"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4020
                    Top =5332
                    Width =861
                    Height =300
                    TabIndex =63
                    Name ="CITY_LEFT"
                    ControlSource ="CITY_LEFT"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1710
                    Top =5332
                    Width =681
                    Height =300
                    TabIndex =65
                    Name ="CITY_SIZE"
                    ControlSource ="CITY_SIZE"
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =480
                    Top =5325
                    TabIndex =67
                    Name ="CITY_P"
                    ControlSource ="CITY_P"
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2475
                    Top =5332
                    Width =1455
                    Height =300
                    TabIndex =64
                    Name ="CITY_FONT"
                    ControlSource ="CITY_FONT"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT font.قلم FROM font ORDER BY font.قلم;"
                End
                Begin Label
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =1
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6090
                    Top =5340
                    Width =1320
                    Height =285
                    BackColor =12632256
                    Name ="Label349"
                    Caption ="شهرستان‌پرداخت:"
                End
                Begin TextBox
                    OverlapFlags =247
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =975
                    Top =5332
                    Width =666
                    Height =300
                    TabIndex =66
                    Name ="CITY_WITH"
                    ControlSource ="CITY_WITH"
                End
                Begin Line
                    OverlapFlags =119
                    SpecialEffect =5
                    Left =120
                    Top =5265
                    Width =7325
                    Name ="Line351"
                End
            End
        End
    End
End
CodeBehindForm
' See "FORMAT_HESAB.cls"
