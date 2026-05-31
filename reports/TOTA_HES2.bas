Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9070
    DatasheetFontHeight =10
    ItemSuffix =37
    Left =1005
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x43ba9e53a891e240
    End
    RecordSource ="SELECT DISTINCTROW TOTA_HES.NUMBER AS kol, TOTA_HES.NAME, DETA_HES.N_KOL, DETA_H"
        "ES.NUMBER, DETA_HES.NAME, TDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, "
        "TDETA_HES.NAME FROM TOTA_HES INNER JOIN (DETA_HES INNER JOIN TDETA_HES ON (DETA_"
        "HES.NUMBER = TDETA_HES.NUMBER) AND (DETA_HES.N_KOL = TDETA_HES.N_KOL)) ON TOTA_H"
        "ES.NUMBER = DETA_HES.N_KOL;"
    Caption ="چاپ حساب"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Arial (Arabic)"
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Arial (Arabic)"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="kol"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="[DETA_HES].[N_KOL] & \", \" & [DETA_HES].[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="TDETA_HES.N_KOL"
        End
        Begin BreakLevel
            ControlSource ="TDETA_HES.NUMBER"
        End
        Begin BreakLevel
            ControlSource ="TNUMBER"
        End
        Begin PageHeader
            Height =495
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3172
                    Width =3075
                    Height =495
                    FontSize =14
                    Name ="Label31"
                    Caption ="سرفصل حسابهاي تعريف شده"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =476
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6626
                    Top =30
                    Width =1110
                    Height =420
                    FontSize =12
                    FontWeight =700
                    BackColor =12632256
                    Name ="kol"
                    ControlSource ="kol"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontFamily =10
                            Left =7856
                            Top =30
                            Width =1170
                            Height =420
                            FontSize =12
                            BackColor =12632256
                            ForeColor =0
                            Name ="TOTA_HES.NUMBER Label"
                            Caption ="حساب كل"
                            FontName ="Zar"
                            EventProcPrefix ="TOTA_HES_NUMBER_Label"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1920
                    Top =30
                    Width =3660
                    Height =420
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BackColor =12632256
                    Name ="TOTA_HES.NAME"
                    ControlSource ="TOTA_HES.NAME"
                    FontName ="Zar"
                    EventProcPrefix ="TOTA_HES_NAME"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontFamily =10
                            Left =5700
                            Top =30
                            Width =900
                            Height =420
                            FontSize =12
                            BackColor =12632256
                            ForeColor =0
                            Name ="TOTA_HES.NAME Label"
                            Caption ="نام حساب"
                            FontName ="Zar"
                            EventProcPrefix ="TOTA_HES_NAME_Label"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =828
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5415
                    Width =1335
                    Height =420
                    FontSize =11
                    ForeColor =8388608
                    Name ="DETA_HE"
                    ControlSource ="=[DETA_HES.N_KOL] & \" - \" & [DETA_HES.NUMBER]"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =6780
                            Width =990
                            Height =375
                            Name ="DETA_HES.NUMBER Label"
                            Caption ="حساب معين"
                            FontName ="Zar"
                            EventProcPrefix ="DETA_HES_NUMBER_Label"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =795
                    Width =3300
                    Height =405
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="DETA_HES.NAME"
                    ControlSource ="DETA_HES.NAME"
                    FontName ="Zar"
                    EventProcPrefix ="DETA_HES_NAME"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =4125
                            Width =1260
                            Height =420
                            Name ="DETA_HES.NAME Label"
                            Caption ="نام حساب معين"
                            FontName ="Zar"
                            EventProcPrefix ="DETA_HES_NAME_Label"
                        End
                    End
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =5215
                    Top =420
                    Width =1065
                    Height =375
                    Name ="TNUMBER Label"
                    Caption ="معين تفضيلي"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="TNUMBER_Label"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =2040
                    Top =420
                    Width =1260
                    Height =375
                    Name ="TDETA_HES.NAME Label"
                    Caption ="نام حساب معين"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="TDETA_HES_NAME_Label"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =448
                    Width =7611
                    Name ="Line25"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =420
                    Width =7611
                    Name ="Line26"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =795
                    Width =7611
                    Name ="Line27"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =150
                    Top =823
                    Width =7611
                    Name ="Line28"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =6859
                    Top =477
                    TabIndex =2
                    Name ="DETA_HES.N_KOL"
                    ControlSource ="DETA_HES.N_KOL"
                    EventProcPrefix ="DETA_HES_N_KOL"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =6973
                    Top =477
                    TabIndex =3
                    Name ="Text34"
                    ControlSource ="DETA_HES.NUMBER"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4932
                    Top =56
                    Width =1710
                    Height =375
                    FontSize =10
                    Name ="TDETA_H"
                    ControlSource ="=[TDETA_HES.N_KOL] & \" - \" & [TDETA_HES.NUMBER] & \" - \" & [TNUMBER]"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =148
                    Top =57
                    Width =4725
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="TDETA_HES.NAME"
                    ControlSource ="TDETA_HES.NAME"
                    FontName ="Zar"
                    EventProcPrefix ="TDETA_HES_NAME"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6675
                    Top =60
                    Width =1080
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text36"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =340
            Name ="GroupFooter2"
            Begin
                Begin Line
                    BorderWidth =3
                    Width =9070
                    BorderColor =12632256
                    Name ="Line32"
                End
            End
        End
        Begin PageFooter
            Height =558
            Name ="PageFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    AllowedText =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4595
                    Top =105
                    Width =4421
                    Height =330
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text19"
                    ControlSource ="=FARSIDATE(ow())"
                    FontName ="Zar"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    AllowedText =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =105
                    Width =4421
                    Height =330
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text20"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =3
                    Width =9070
                    BorderColor =12632256
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =3
                    Top =28
                    Width =9070
                    BorderColor =12632256
                    Name ="Line30"
                End
            End
        End
    End
End
CodeBehindForm
' See "TOTA_HES2.cls"
