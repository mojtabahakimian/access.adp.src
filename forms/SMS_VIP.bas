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
    Width =6803
    DatasheetFontHeight =10
    ItemSuffix =22
    Left =7350
    Top =3000
    Right =14160
    Bottom =6330
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x41a653f29e1ae240
    End
    Caption ="ارسال اس ام اس"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
            Height =3344
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3345
                    Top =2721
                    Width =1281
                    Height =418
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Command1"
                    Caption ="تائيد"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =26
                    Top =170
                    Width =915
                    Height =1020
                    FontSize =48
                    Name ="Label13"
                    Caption ="V"
                    FontName ="Webdings"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =223
                    Left =850
                    Top =226
                    Width =170
                    Height =227
                    Name ="Line14"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =223
                    Left =907
                    Top =453
                    Width =226
                    Height =113
                    Name ="Line15"
                End
                Begin Line
                    OverlapFlags =95
                    Left =915
                    Top =675
                    Width =283
                    Height =57
                    Name ="Line16"
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =1
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =287
                    Top =1147
                    Width =6171
                    Height =1403
                    Name ="MATN"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2070
                    Top =2721
                    Width =1281
                    Height =418
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Command19"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    Left =1077
                    Top =226
                    Width =5385
                    Height =510
                    Name ="Label21"
                    Caption ="متن مورد نظر براي ارسال را در ذيل وارد كنيد اسم مشتري به صورت خودكار به پيام اضا"
                        "فه ميشود"
                    FontName ="Tahoma"
                End
            End
        End
    End
End
CodeBehindForm
' See "SMS_VIP.cls"
