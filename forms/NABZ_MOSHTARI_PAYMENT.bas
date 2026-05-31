Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15193
    DatasheetFontHeight =10
    ItemSuffix =66
    Left =30
    Top =1560
    Right =15015
    Bottom =11760
    TimerInterval =10000
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0x30370547f3bde540
    End
    RecordSource ="SSM_CUST_ANALYS"
    Caption ="پيغام"
    DatasheetFontName ="Arial (Arabic)"
    OnTimer ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin Image
            BackStyle =0
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
            FontName ="MS Sans Serif"
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =10215
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7484
                    Top =170
                    Width =6816
                    Height =330
                    TabIndex =1
                    BackColor =16764057
                    Name ="MOSHTARI"
                    ControlSource ="=[Forms]![NABZ_MOSHTARI]![MOSHTARI]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14366
                            Top =170
                            Width =615
                            Height =330
                            Name ="Label35"
                            Caption ="مشتري"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =0
                    Left =7656
                    Top =793
                    Width =7140
                    Height =4260
                    Name ="NABZ-khareed_MOSHTARI"
                    SourceObject ="Form.NABZ-khareed_MOSHTARI"
                    LinkChildFields ="MMO"
                    LinkMasterFields ="MONTH"
                    EventProcPrefix ="NABZ_khareed_MOSHTARI"
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =0
                    Left =7650
                    Top =5160
                    Width =7140
                    Height =4785
                    TabIndex =2
                    Name ="NABZ-PARDAKHT_MOSHTARI"
                    SourceObject ="Form.NABZ-PARDAKHT_MOSHTARI"
                    LinkChildFields ="MONTH;CUST_NO"
                    LinkMasterFields ="MONTH;CUST_NO"
                    EventProcPrefix ="NABZ_PARDAKHT_MOSHTARI"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =3514
                    Top =226
                    Width =1716
                    Height =300
                    TabIndex =3
                    Name ="MONTH"
                    ControlSource ="MONTH"
                    RowSourceType ="Value List"
                    RowSource ="1;\"فروردين\";2;\"ارديبهشت\";3;\"خرداد\";4;\"تير\";5;\"مرداد\";6;\"شهريور\";7;\""
                        "مهر\";8;\"آبان\";9;\"آذر\";10;\"دي\";11;\"بهمن\";12;\"اسفند\""
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =5371
                            Top =226
                            Width =2070
                            Height =240
                            Name ="Label4"
                            Caption ="گزارش وضعيت پرداختهاي ماه"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =630
                    Top =283
                    Width =1461
                    TabIndex =4
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2217
                            Top =283
                            Width =1185
                            Height =240
                            Name ="Label14"
                            Caption ="شماره مشتري"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2721
                    Top =1870
                    Width =336
                    TabIndex =5
                    Name ="SCAID"
                    ControlSource ="SCAID"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =907
                            Top =1870
                            Width =540
                            Height =240
                            Name ="Label1"
                            Caption ="شماره"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2721
                    Top =2210
                    Width =1110
                    TabIndex =6
                    Name ="Text60"
                    ControlSource ="CUST_NO"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =914
                            Top =2207
                            Width =840
                            Height =240
                            Name ="Label3"
                            Caption ="کد مشتري"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2721
                    Top =2550
                    Width =336
                    TabIndex =7
                    Name ="SCADATE"
                    ControlSource ="SCADATE"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =907
                            Top =2550
                            Width =765
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =5220
                    Width =2280
                    TabIndex =8
                    Name ="SCAJAMF"
                    ControlSource ="SCAJAMF"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =5220
                            Width =810
                            Height =240
                            Name ="Label7"
                            Caption ="جمع فاکتور"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =5569
                    Width =2280
                    TabIndex =9
                    Name ="SCARASFR"
                    ControlSource ="SCARASFR"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =5565
                            Width =975
                            Height =240
                            Name ="Label9"
                            Caption ="راس فاکتورها"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =5940
                    Width =2280
                    TabIndex =10
                    Name ="SCARASTAV"
                    ControlSource ="SCARASTAV"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =5940
                            Width =810
                            Height =240
                            Name ="Label11"
                            Caption ="راس توافق"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =6756
                    Width =2280
                    TabIndex =11
                    Name ="SCARASPAY"
                    ControlSource ="SCARASPAY"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =6813
                            Width =915
                            Height =240
                            Name ="Label13"
                            Caption ="راس پرداخت"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =7153
                    Width =2280
                    TabIndex =12
                    Name ="SCATAFAVOT"
                    ControlSource ="SCATAFAVOT"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ff660000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =7153
                            Width =405
                            Height =240
                            Name ="Label15"
                            Caption ="تاخير"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =6360
                    Width =2280
                    TabIndex =13
                    Name ="SCAMABNADATE"
                    ControlSource ="SCAMABNADATE"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =6416
                            Width =690
                            Height =240
                            Name ="Label17"
                            Caption ="تاريخ مبنا"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1440
                    Top =7530
                    Width =2280
                    TabIndex =14
                    Name ="SCAMABLA"
                    ControlSource ="SCAMABLA"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Tahoma"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000003000000000000000200000001000000 ,
                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3836
                            Top =7530
                            Width =750
                            Height =240
                            Name ="Label19"
                            Caption ="مبلغ تاخير"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4932
                    Top =741
                    TabIndex =15
                    Name ="SCAMABP"
                    ControlSource ="SCAMABP"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3125
                            Top =737
                            Width =1365
                            Height =240
                            Name ="Label21"
                            Caption ="درصد کسر و اضافه"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4932
                    Top =1082
                    TabIndex =16
                    Name ="CTIM"
                    ControlSource ="CTIM"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3118
                            Top =1082
                            Width =450
                            Height =240
                            Name ="Label23"
                            Caption ="تاريخ"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4932
                    Top =1422
                    TabIndex =17
                    Name ="USERCO"
                    ControlSource ="USERCO"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3118
                            Top =1422
                            Width =690
                            Height =240
                            Name ="Label25"
                            Caption ="کد کاربر"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4932
                    Top =1762
                    TabIndex =18
                    Name ="Text61"
                    ControlSource ="MONTH"
                    FontName ="Tahoma"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3118
                            Top =1762
                            Width =615
                            Height =240
                            Name ="Label27"
                            Caption ="ماه"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    Left =1077
                    Top =5555
                    Width =270
                    Height =240
                    Name ="Label62"
                    Caption ="روز"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    Left =1077
                    Top =5915
                    Width =270
                    Height =240
                    Name ="Label63"
                    Caption ="روز"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    Left =1095
                    Top =6750
                    Width =270
                    Height =240
                    Name ="Label64"
                    Caption ="روز"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    Left =1095
                    Top =7155
                    Width =270
                    Height =240
                    Name ="Label65"
                    Caption ="روز"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "NABZ_MOSHTARI_PAYMENT.cls"
