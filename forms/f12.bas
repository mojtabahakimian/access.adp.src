Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =2608
    DatasheetFontHeight =10
    ItemSuffix =28
    Left =360
    Top =5910
    Right =3015
    Bottom =7350
    RecSrcDt = Begin
        0x97a17a1adbc1e540
    End
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    FilterOnLoad =0
    SplitFormSplitterBar =0
    SplitFormSplitterBar =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
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
        Begin FormHeader
            Height =453
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Width =2608
                    Height =453
                    Name ="Command2"
                    Caption ="تدارکات و انبار"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5159
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =2495
                    Height =426
                    Name ="P1E3"
                    Caption ="صدور برگه حواله انبار فروش"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =567
                    Width =2495
                    Height =426
                    Name ="P2E3"
                    Caption ="صدور برگه رسيد انبار خريد"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =1134
                    Width =2495
                    Height =426
                    Name ="P3E3"
                    Caption ="انتقال کالا از انبار به انبار"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =1701
                    Width =2495
                    Height =426
                    Name ="P4E3"
                    Caption ="ساير رسيد انبارها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =2268
                    Width =2495
                    Height =426
                    Name ="P5E3"
                    Caption ="ساير حواله انبار ها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =2835
                    Width =2495
                    Height =426
                    Name ="P6E3"
                    Caption ="درخواست خريد"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =3402
                    Width =2495
                    Height =426
                    Name ="P7E3"
                    Caption ="انبار گرداني"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =64
                    Top =4536
                    Width =2495
                    Height =426
                    Name ="P9E3"
                    Caption ="جستجوگر موجودي انبار"
                    OnClick ="[Event Procedure]"
                End
                Begin Subform
                    OverlapFlags =85
                    OldBorderStyle =0
                    SpecialEffect =0
                    Left =60
                    Top =3975
                    Width =2495
                    Height =442
                    TabIndex =1
                    Name ="f12-1"
                    SourceObject ="Form.f12-1"
                    EventProcPrefix ="f12_1"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =50
                    Height =5
                    Name ="Text58"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "f12.cls"
