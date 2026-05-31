Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9694
    DatasheetFontHeight =10
    ItemSuffix =13
    Left =4515
    Top =1515
    Right =14205
    Bottom =8415
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa6a8f965b9a1e240
    End
    Caption ="راس گيري"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =6916
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6973
                    Top =170
                    Width =1641
                    Height =315
                    Name ="Text0"
                    Format ="0000\\/00\\/00"
                    AfterUpdate ="[Event Procedure]"
                    OnExit ="[Event Procedure]"
                    DefaultValue ="farsidate(Date())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8678
                            Top =177
                            Width =870
                            Height =300
                            Name ="Label1"
                            Caption ="تاريخ مبنا:"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =3272
                    Top =692
                    Width =6255
                    Height =5205
                    TabIndex =1
                    Name ="RAAS_SUB"
                    SourceObject ="Form.RAAS_SUB"
                    LinkChildFields ="USERID"
                    LinkMasterFields ="USERID"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Top =680
                    Width =1971
                    Height =315
                    TabIndex =2
                    BackColor =13434828
                    Name ="MABL"
                    ControlSource ="=[RAAS_SUB].[Form]![SMABK]/[RAAS_SUB].[Form]![text6]"
                    Format ="#,##0.00\" روز\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2250
                            Top =690
                            Width =915
                            Height =300
                            Name ="Label5"
                            Caption ="راس :"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =735
                    Top =1080
                    Width =1401
                    Height =315
                    TabIndex =3
                    BackColor =16764057
                    Name ="Text4"
                    ControlSource ="=RAAS_SUB.Form!SROOS/RAAS_SUB.Form.CID"
                    Format ="#,##0.00\" روز\""
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2245
                            Top =1090
                            Width =825
                            Height =300
                            Name ="Label6"
                            Caption ="راس مدت:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6519
                    Top =6066
                    Width =1971
                    Height =315
                    TabIndex =4
                    Name ="Text7"
                    ControlSource ="=RAAS_SUB.Form!Text6"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8599
                            Top =6076
                            Width =915
                            Height =300
                            Name ="Label8"
                            Caption ="جمع مبالغ:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =735
                    Top =1440
                    Width =1401
                    Height =315
                    TabIndex =5
                    BackColor =16764057
                    Name ="Text9"
                    ControlSource ="=RAAS_SUB.Form!SROOS/RAAS_SUB.Form.CID/30"
                    Format ="#,##0.00\" ماه\""
                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =2245
                            Top =1450
                            Width =825
                            Height =300
                            Name ="Label10"
                            Caption ="راس مدت:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3741
                    Top =6066
                    Width =1971
                    Height =315
                    TabIndex =6
                    Name ="Text11"
                    ControlSource ="=[RAAS_SUB].[Form]![SMABK]"
                    Format ="#,###"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =3
                    ScrollBarAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Top =90
                    Width =2394
                    Height =345
                    TabIndex =7
                    BackColor =13434828
                    Name ="USERID"
                    DefaultValue ="=[Forms]![BASEKNOW]![USERCOD]"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2433
                            Top =90
                            Width =1560
                            Height =345
                            Name ="Label18"
                            Caption ="کاربر"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "RAAS.cls"
