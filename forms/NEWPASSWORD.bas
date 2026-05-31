Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4139
    DatasheetFontHeight =10
    ItemSuffix =14
    Left =9990
    Top =3330
    Right =14145
    Bottom =5415
    HelpContextId =182
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5bbc800e1e1de240
    End
    Caption ="تعريف رمز عبور"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
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
        Begin ComboBox
            SpecialEffect =2
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2097
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2319
                    Top =120
                    Width =621
                    Height =300
                    TabIndex =4
                    Name ="usname"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="[Forms]![baseknow]![USERCOD]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3060
                            Top =165
                            Width =885
                            Height =285
                            Name ="Label1"
                            Caption ="آي دي كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =224
                    Top =832
                    Width =2721
                    Height =300
                    TabIndex =1
                    Name ="pass1"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3074
                            Top =881
                            Width =645
                            Height =285
                            Name ="Label5"
                            Caption ="رمز عبور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =224
                    Top =1177
                    Width =2721
                    Height =300
                    TabIndex =2
                    Name ="pass2"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3074
                            Top =1213
                            Width =1065
                            Height =285
                            Name ="Label7"
                            Caption ="مجدداً رمز عبور"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =1666
                    Top =1650
                    Width =1305
                    Height =390
                    FontWeight =700
                    TabIndex =3
                    Name ="Command8"
                    Caption ="ت&ائيد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =228
                    TextFontCharSet =0
                    Left =224
                    Top =1656
                    Width =1275
                    Height =390
                    FontWeight =700
                    TabIndex =5
                    Name ="Command9"
                    Caption ="ا&نصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =224
                    Top =487
                    Width =2721
                    Height =300
                    Name ="OLDPASS"
                    DefaultValue ="\"\""
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3068
                            Top =529
                            Width =795
                            Height =285
                            Name ="Label11"
                            Caption ="رمز قديمي"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =226
                    Top =113
                    Width =2046
                    Height =300
                    TabIndex =6
                    Name ="usna"
                    DefaultValue ="UCurrentUser()"
                End
            End
        End
    End
End
CodeBehindForm
' See "NEWPASSWORD.cls"
