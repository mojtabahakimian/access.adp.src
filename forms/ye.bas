Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7993
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =4995
    Top =45
    Right =12795
    Bottom =300
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x53032f4f2a9ae240
    End
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Section
            Height =285
            BackColor =16737843
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =3
                    BackStyle =0
                    Left =6806
                    Width =555
                    Height =285
                    FontWeight =700
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16777215
                    Name ="YEE"
                    ControlSource ="=Forms!Baseknow!YEA"
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =0
                    Left =7361
                    Width =600
                    Height =285
                    FontWeight =700
                    BackColor =16711680
                    ForeColor =16777215
                    Name ="Label0"
                    Caption ="سال :"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =0
                    TextAlign =2
                    BackStyle =0
                    Left =-3
                    Width =6750
                    Height =285
                    FontWeight =700
                    TabIndex =1
                    BackColor =16711680
                    BorderColor =3355443
                    ForeColor =16777215
                    Name ="Text1"
                    ControlSource ="=[Forms]![Baseknow]![namee] & \" : \" & [Forms]![Baseknow]![YEA]"
                End
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    Width =3270
                    Height =210
                    BackColor =12632256
                    ForeColor =255
                    Name ="version"
                    Caption ="Version :7.1.1: date : 1405/02/28  21:59"
                    FontName ="MS Sans Serif"
                End
            End
        End
    End
End
CodeBehindForm
' See "ye.cls"
