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
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
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
    Width =8617
    DatasheetFontHeight =10
    ItemSuffix =5
    Left =3330
    Top =3030
    Right =11940
    Bottom =4320
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x358ea6b70e6fe240
    End
    Caption ="جمع ستون انتخابي"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
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
            Height =1306
            BackColor =10079487
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =60
                    Width =7461
                    Height =765
                    FontSize =20
                    Name ="Text0"
                    Format ="#,##0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =10
                            Left =7575
                            Top =64
                            Width =990
                            Height =750
                            FontSize =20
                            Name ="Label1"
                            Caption ="جمع  :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3632
                    Top =907
                    Width =3852
                    Height =340
                    TabIndex =1
                    Name ="Text2"
                    Format ="Fixed"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =7935
                    Top =825
                    Width =441
                    Height =456
                    TabIndex =2
                    Name ="Command4"
                    Caption ="Command4"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000000d0fffffffffffff0a ,
                        0x0f7777777fffff0d0f7f7f7f7fffff0a0f77777777777f0d0f7f7f7f7f7f7f0a ,
                        0x0f77777777777f0d0f7f7f7f7f7f7f0a0f77777777777f0d0f7f7f7f7f7f7f0a ,
                        0x0f77777777777f0d0fffffffffffff0a0f777777fff77f0d0fffffffffffff0a ,
                        0x000000000000000d000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                    Tag ="ماشين حساب"
                End
            End
        End
    End
End
CodeBehindForm
' See "JAMECOL.cls"
