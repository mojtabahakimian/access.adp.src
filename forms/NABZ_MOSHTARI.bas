Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
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
    Width =15193
    DatasheetFontHeight =10
    ItemSuffix =59
    Right =15045
    Bottom =10200
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x41a653f29e1ae240
    End
    Caption ="پيغام"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
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
        Begin Tab
            TextFontCharSet =178
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin Page
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =10215
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Left =225
                    Top =6270
                    Width =14760
                    Height =3825
                    TabIndex =1
                    TabFixedHeight =403
                    Name ="TabCtl18"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =360
                            Top =6810
                            Width =14490
                            Height =3150
                            Name ="Page19"
                            Caption ="پيش فاکتور هاي مشتري"
                            Begin
                                Begin Subform
                                    OverlapFlags =215
                                    Left =503
                                    Top =6811
                                    Width =14190
                                    Height =2760
                                    Name ="grade_pish"
                                    SourceObject ="Form.grade_pish"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =360
                            Top =6810
                            Width =14490
                            Height =3150
                            Name ="Page20"
                            Caption ="حواله هاي مشتري"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =443
                                    Top =6811
                                    Width =14250
                                    Height =2761
                                    Name ="grade_havala"
                                    SourceObject ="Form.grade_havala"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =360
                            Top =6810
                            Width =14490
                            Height =3150
                            Name ="Page28"
                            Caption ="چکهاي برگشتي"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =503
                                    Top =6811
                                    Width =14190
                                    Height =2761
                                    Name ="grade_chek_bargashti"
                                    SourceObject ="Form.grade_chek_bargashti"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7484
                    Top =170
                    Width =6816
                    Height =330
                    BackColor =16764057
                    Name ="MOSHTARI"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =14366
                            Top =170
                            Width =615
                            Height =330
                            Name ="Label35"
                            Caption ="مشتري"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin OptionGroup
                    BackStyle =1
                    OverlapFlags =93
                    Left =7484
                    Top =729
                    Width =7372
                    Height =3345
                    TabIndex =2
                    Name ="Frame37"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9494
                    Top =1978
                    Width =2721
                    TabIndex =3
                    Name ="mhesab"
                    Format ="#,###"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =12275
                            Top =1981
                            Width =945
                            Height =240
                            Name ="Label42"
                            Caption ="مانده حساب:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9494
                    Top =2548
                    Width =2721
                    TabIndex =4
                    Name ="mpish"
                    Format ="#,##0\" ريال \";-#,##0\" ريال \""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =12275
                            Top =2548
                            Width =2295
                            Height =240
                            Name ="Label45"
                            Caption ="مانده پيش فاکتورهاي تاييد شده:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9494
                    Top =3103
                    Width =2721
                    TabIndex =5
                    Name ="mhav"
                    Format ="#,##0\" ريال \";-#,##0\" ريال \""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =12275
                            Top =3103
                            Width =2295
                            Height =240
                            Name ="Label47"
                            Caption ="جمع حواله هاي بارگيري نشده:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9494
                    Top =3673
                    Width =2721
                    TabIndex =6
                    Name ="masnad"
                    Format ="#,##0\" ريال \";-#,##0\" ريال \""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =12275
                            Top =3673
                            Width =2295
                            Height =240
                            Name ="Label53"
                            Caption ="مانده اسناد پاس نشده:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9494
                    Top =1243
                    Width =2721
                    Height =510
                    FontSize =16
                    TabIndex =7
                    Name ="grdn"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =12269
                            Top =1363
                            Width =1020
                            Height =240
                            Name ="Label59"
                            Caption ="گريد مشتري:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =215
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7650
                    Top =853
                    Width =4566
                    TabIndex =8
                    Name ="MNAM"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextFontCharSet =0
                            TextAlign =1
                            Left =12269
                            Top =856
                            Width =690
                            Height =240
                            Name ="Label84"
                            Caption ="مشتري:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2145
                    Top =390
                    Width =3464
                    Height =511
                    TabIndex =9
                    Name ="Command36"
                    Caption ="فرم اعتبار سنحي"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2145
                    Top =1021
                    Width =3464
                    Height =511
                    TabIndex =10
                    Name ="Command53"
                    Caption ="گزارش وضعيت پرداختهاي مشتري"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2145
                    Top =1652
                    Width =3464
                    Height =511
                    TabIndex =11
                    Name ="Command54"
                    Caption ="فرمهاي بايگاني و پرونده ها"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =10091
                    Top =4875
                    Width =3464
                    Height =511
                    TabIndex =12
                    Name ="Command55"
                    Caption ="ليست احناس فروخته نشده به مشتري"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Subform
                    OverlapFlags =85
                    SpecialEffect =0
                    Left =221
                    Top =2565
                    Width =7140
                    Height =3630
                    TabIndex =13
                    BorderColor =10079487
                    Name ="amar_FROOSH_MOSHTARI"
                    SourceObject ="Form.amar_FROOSH_MOSHTARI"
                End
            End
        End
    End
End
CodeBehindForm
' See "NABZ_MOSHTARI.cls"
