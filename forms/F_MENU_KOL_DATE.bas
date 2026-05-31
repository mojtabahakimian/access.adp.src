Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
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
    Width =5952
    DatasheetFontHeight =10
    ItemSuffix =31
    Left =6570
    Top =2550
    Right =12525
    Bottom =4665
    HelpContextId =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="دفتر كل"
    OnOpen ="[Event Procedure]"
    HelpFile ="amin98.hlp>main"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =2131
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3043
                    Top =1694
                    Width =870
                    Height =390
                    FontWeight =700
                    TabIndex =7
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2175
                    Top =1694
                    Width =870
                    Height =392
                    FontWeight =700
                    TabIndex =8
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5665
                    Top =510
                    Width =234
                    Height =255
                    TabIndex =4
                    Name ="DTT"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3113
                    Top =1140
                    Width =1431
                    Height =300
                    TabIndex =5
                    Name ="DT1"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4638
                            Top =1142
                            Width =840
                            Height =285
                            Name ="Label21"
                            Caption ="از تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =885
                    Top =1148
                    Width =1416
                    Height =300
                    TabIndex =6
                    Name ="DT2"
                    Format ="0000\\/00\\/00"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=Farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2395
                            Top =1150
                            Width =630
                            Height =285
                            Name ="Label23"
                            Caption ="تا تاريخ:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3117
                    Top =225
                    Width =1431
                    Height =300
                    Name ="HKOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4635
                            Top =225
                            Width =990
                            Height =285
                            Name ="Label17"
                            Caption ="از حساب كل:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =173
                    Top =239
                    Width =2841
                    Height =300
                    TabIndex =1
                    Name ="HHKOL"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3117
                    Top =659
                    Width =1431
                    Height =300
                    TabIndex =2
                    Name ="HKOL2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER FROM TOTA_HES ORDER BY TOTA_HES.NUMBER;"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =4635
                            Top =660
                            Width =975
                            Height =285
                            Name ="Label29"
                            Caption ="تا حساب كل:"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =173
                    Top =673
                    Width =2841
                    Height =300
                    TabIndex =3
                    Name ="HHKOL2"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT TOTA_HES.NUMBER, TOTA_HES.NAME FROM TOTA_HES ORDER BY TOTA_HES.NAME;"
                    ColumnWidths ="0"
                    AfterUpdate ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_KOL_DATE.cls"
