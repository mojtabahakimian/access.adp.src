Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =11
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10560
    DatasheetFontHeight =10
    ItemSuffix =72
    Left =270
    Top =600
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xee07edf369e5e240
    End
    Caption ="ترازنامه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin PageHeader
            Height =1155
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =4748
                    Width =1140
                    Height =975
                    FontSize =10
                    FontWeight =400
                    Name ="Label24"
                    Caption ="ترازنامه"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2760
                    Top =690
                    Width =4998
                    Height =465
                    FontSize =10
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                    AsianLineBreak =0
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =2
                    Left =56
                    Top =737
                    Width =1020
                    Height =360
                    FontSize =10
                    FontWeight =400
                    Name ="dt"
                    Caption ="ترازنامه"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    FontItalic = NotDefault
                    BackStyle =1
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =2
                    Left =1080
                    Top =735
                    Width =1485
                    Height =360
                    FontSize =10
                    FontWeight =400
                    Name ="Label67"
                    Caption ="كليه اسناد تا سند :"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =11700
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Subform
                    Left =5265
                    Top =637
                    Width =5266
                    Height =9960
                    Name ="TARAZNAMAH1"
                    SourceObject ="Report.TARAZNAMAH1"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =178
                            TextAlign =2
                            TextFontFamily =2
                            Left =7483
                            Width =825
                            Height =540
                            FontSize =10
                            FontWeight =400
                            Name ="TARAZNAMAH1 Label"
                            Caption ="دارائيها"
                            FontName ="WeblogmaYekan"
                            EventProcPrefix ="TARAZNAMAH1_Label"
                        End
                    End
                End
                Begin Subform
                    Top =637
                    Width =5220
                    Height =9960
                    TabIndex =1
                    Name ="TARAZNAMAH11"
                    SourceObject ="Report.TARAZNAMAH11"
                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            TextFontCharSet =178
                            TextAlign =2
                            TextFontFamily =2
                            Left =1757
                            Width =1875
                            Height =540
                            FontSize =10
                            FontWeight =400
                            ForeColor =128
                            Name ="Label56"
                            Caption ="بدهي ها و سرمايه"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =120
                    Top =10590
                    Width =3061
                    Height =540
                    FontSize =10
                    TabIndex =2
                    Name ="hogh"
                    ControlSource ="=IIf(IsError([TARAZNAMAH11].[Report]![Sbedehe]),0,[TARAZNAMAH11].[Report]![Sbede"
                        "he])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    AsianLineBreak =0
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =5340
                    Top =10605
                    Width =3061
                    Height =540
                    FontSize =10
                    TabIndex =3
                    Name ="Text58"
                    ControlSource ="=IIf(IsError([TARAZNAMAH1].[Report]![SDARAEE]),0,[TARAZNAMAH1].[Report]![SDARAEE"
                        "])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    AsianLineBreak =0
                End
                Begin Line
                    BorderWidth =2
                    Top =10620
                    Width =10548
                    BorderColor =0
                    Name ="Line59"
                End
                Begin Label
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =4081
                    Top =10658
                    Width =450
                    Height =435
                    FontSize =10
                    FontWeight =400
                    Name ="Label61"
                    Caption ="جمع"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =9524
                    Top =10658
                    Width =450
                    Height =435
                    FontSize =10
                    FontWeight =400
                    Name ="Label62"
                    Caption ="جمع"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =5272
                    Width =0
                    Height =11168
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Rectangle
                    BorderWidth =2
                    Width =10544
                    Height =11160
                    BorderColor =0
                    Name ="Box65"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    TextFontCharSet =178
                    TextAlign =1
                    TextFontFamily =2
                    NumeralShapes =2
                    Left =120
                    Top =11160
                    Width =3061
                    Height =540
                    FontSize =10
                    TabIndex =4
                    ForeColor =255
                    Name ="Text68"
                    ControlSource ="=[TEXT58]-[HOGH]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                    AsianLineBreak =0
                End
                Begin Line
                    BorderWidth =2
                    Top =11685
                    Width =10551
                    Name ="Line70"
                End
                Begin Label
                    TextFontCharSet =178
                    TextFontFamily =2
                    Left =5442
                    Top =11225
                    Width =2805
                    Height =405
                    FontSize =10
                    FontWeight =400
                    Name ="Label71"
                    Caption =" دارائيها - بديها و سرمايه:"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "TARAZNAMAH2.cls"
