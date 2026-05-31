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
    Width =4251
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =6585
    Top =3225
    Right =10845
    Bottom =5085
    HelpContextId =182
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5bbc800e1e1de240
    End
    Caption ="تعريف نام كاربر يا حذف آن"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
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
            Height =1870
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =325
                    Top =120
                    Width =2721
                    Height =300
                    Name ="usname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3168
                            Top =163
                            Width =630
                            Height =285
                            Name ="Label1"
                            Caption ="نام كاربر"
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
                    Left =330
                    Top =832
                    Width =2721
                    Height =300
                    TabIndex =2
                    Name ="pass1"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3180
                            Top =881
                            Width =645
                            Height =285
                            Name ="Label5"
                            Caption ="رمز عبور"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =2008
                    Top =1303
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
                    Left =566
                    Top =1309
                    Width =1275
                    Height =390
                    FontWeight =700
                    TabIndex =4
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
                    Left =330
                    Top =487
                    Width =2721
                    Height =300
                    TabIndex =1
                    Name ="NEWUSER"
                    DefaultValue ="\"\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3181
                            Top =525
                            Width =990
                            Height =285
                            Name ="Label11"
                            Caption ="نام كاربر جديد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "USER_CHANGE.cls"
