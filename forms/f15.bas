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
    Width =2615
    DatasheetFontHeight =10
    ItemSuffix =84
    Left =360
    Top =7305
    Right =3000
    Bottom =8745
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
            Height =468
            BackColor =16766905
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    Top =15
                    Width =2608
                    Height =453
                    Name ="Command2"
                    Caption ="تعمير و نگهداري(CMMS)"
                    OnClick ="[Event Procedure]"
                    BorderLineStyle =0
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5159
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Width =50
                    Height =5
                    Name ="Text58"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =2495
                    Height =426
                    Name ="P1E3"
                    Caption ="مديريت دستور کارها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =606
                    Width =2495
                    Height =426
                    Name ="P2E3"
                    Caption ="تعريف دستگاه ها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =1152
                    Width =2495
                    Height =426
                    Name ="P3E3"
                    Caption ="تعريف محل استفاده دستگاه ها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =1698
                    Width =2495
                    Height =426
                    Name ="P4E3"
                    Caption ="اعلام خرابي"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =2244
                    Width =2495
                    Height =426
                    Name ="P5E3"
                    Caption ="تعريف نوع تعمير و نگهداري (Work)"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =2790
                    Width =2495
                    Height =426
                    Name ="P6E3"
                    Caption ="تقويم دستور کارها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =3336
                    Width =2495
                    Height =426
                    Name ="P7E3"
                    Caption ="درخواست کالا از انبار"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =3882
                    Width =2495
                    Height =426
                    Name ="P8E3"
                    Caption ="ليست دستور کارها"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =4428
                    Width =2495
                    Height =426
                    Name ="P9E3"
                    Caption ="تعريف نوع خرابي"
                    OnClick ="[Event Procedure]"
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
' See "f15.cls"
