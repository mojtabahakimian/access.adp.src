Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    RowHeight =330
    ItemSuffix =10
    Left =1515
    Top =1290
    Right =15390
    Bottom =6600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7bf8a85bd252e540
    End
    InputParameters ="@F_MENU_DATE_DT1=FORMS![F_MENU_DATE]![DT1],@F_MENU_DATE_DT2=FORMS![F_MENU_DATE]!"
        "[DT2]"
    RecordSource ="FROOSH_COUNTALL"
    Caption ="ليست فروش روزانه به تفكيك نوع كالا"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,170 } Begin
        0x400046005f004d0045004e0055005f0044004100540045005f00440054003100 ,
        0x3d0046004f0052004d00530021005b0046005f004d0045004e0055005f004400 ,
        0x4100540045005d0021005b004400540031005d002c00400046005f004d004500 ,
        0x4e0055005f0044004100540045005f004400540032003d0046004f0052004d00 ,
        0x530021005b0046005f004d0045004e0055005f0044004100540045005d002100 ,
        0x5b004400540032005d00
    End
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        Begin Section
            Height =1834
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    ColumnOrder =2
                    Name ="TEDAD"
                    ControlSource ="TEDAD"
                    OnClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =570
                            Height =240
                            Name ="Label1"
                            Caption ="تعداد*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    Width =2490
                    ColumnOrder =0
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =453
                            Width =495
                            Height =240
                            Name ="Label3"
                            Caption ="كد*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    Width =2490
                    ColumnWidth =5070
                    ColumnOrder =1
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =795
                            Width =630
                            Height =240
                            Name ="Label5"
                            Caption ="نام كالا*"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    ColumnWidth =1065
                    TabIndex =3
                    Name ="MEGHk"
                    ControlSource ="expr1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =960
                            Height =240
                            Name ="Label7"
                            Caption ="مقدار كل كالا"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =178
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =1584
                    ColumnWidth =1995
                    TabIndex =4
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    AsianLineBreak =0
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =178
                            Left =2709
                            Top =1586
                            Width =870
                            Height =240
                            Name ="Label27"
                            Caption ="شماره فني"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "FROOSH_COUNTALL.cls"
