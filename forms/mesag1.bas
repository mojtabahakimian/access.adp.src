Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
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
    Width =6916
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =2490
    Top =2220
    Right =9405
    Bottom =4020
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x41a653f29e1ae240
    End
    Caption =" "
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin Line
            Width =1701
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =1814
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3187
                    Top =1305
                    Width =1281
                    Height =418
                    FontWeight =700
                    Name ="Command1"
                    Caption ="تائيد"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =1342
                    Top =285
                    Width =5433
                    Height =914
                    TabIndex =1
                    BackColor =12632256
                    BorderColor =13056
                    Name ="msg"
                    ControlSource ="=[OpenArgs]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =18
                    Top =165
                    Width =1305
                    Height =1020
                    FontSize =48
                    Name ="Label13"
                    Caption ="V"
                    FontName ="Webdings"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =215
                    Left =848
                    Top =225
                    Width =170
                    Height =227
                    Name ="Line14"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =215
                    Left =893
                    Top =510
                    Width =226
                    Height =113
                    Name ="Line15"
                End
                Begin Line
                    OverlapFlags =215
                    Left =893
                    Top =795
                    Width =283
                    Height =57
                    Name ="Line16"
                End
            End
        End
    End
End
CodeBehindForm
' See "mesag1.cls"
