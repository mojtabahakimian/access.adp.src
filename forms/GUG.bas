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
    AllowDeletions = NotDefault
    CloseButton = NotDefault
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
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =3060
    Top =2535
    Right =8730
    Bottom =4170
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x803b8a68b86ce240
    End
    Caption ="صدور اسناد گروهي"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Section
            Height =1644
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextFontFamily =10
                    BackStyle =0
                    Left =352
                    Top =825
                    Width =4743
                    Height =180
                    FontSize =6
                    ForeColor =8388608
                    Name ="Text0"
                    DefaultValue ="\"n\""
                    FontName ="Wingdings"
                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =345
                    Top =826
                    Width =4743
                    Height =185
                    Name ="Box7"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =2
                    TextFontFamily =2
                    BackStyle =0
                    Left =352
                    Top =1131
                    Width =4743
                    Height =315
                    FontSize =20
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text2"
                    DefaultValue ="\"n\""
                    FontName ="Wingdings"
                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =345
                    Top =1132
                    Width =4743
                    Height =320
                    Name ="Box3"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3057
                    Top =166
                    Width =850
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="NUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3959
                            Top =165
                            Width =825
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label5"
                            Caption ="شماره: "
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =960
                    Top =166
                    Width =850
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="SNUM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            TextAlign =2
                            Left =1862
                            Top =165
                            Width =1080
                            Height =285
                            FontSize =10
                            FontWeight =700
                            Name ="Label7"
                            Caption ="از: "
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "GUG.cls"
