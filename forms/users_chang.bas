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
    Width =4081
    DatasheetFontHeight =10
    ItemSuffix =12
    Left =4020
    Top =2055
    Right =8115
    Bottom =4065
    HelpContextId =182
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x5bbc800e1e1de240
    End
    Caption ="تعريف كاربر جديد"
    OnCurrent ="[Event Procedure]"
    HelpFile ="AMIN98.HLP>MAIN"
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
            Height =2032
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
                    Left =165
                    Top =90
                    Width =2721
                    Height =300
                    Name ="usname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3008
                            Top =133
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
                    Left =165
                    Top =767
                    Width =2721
                    Height =300
                    TabIndex =2
                    Name ="pass1"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3015
                            Top =816
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
                    Left =165
                    Top =1112
                    Width =2721
                    Height =300
                    TabIndex =3
                    Name ="pass2"
                    InputMask ="Password"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3015
                            Top =1148
                            Width =1065
                            Height =285
                            Name ="Label7"
                            Caption ="مجدداً رمز عبور"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =199
                    TextFontCharSet =0
                    Left =1762
                    Top =1590
                    Width =1133
                    Height =412
                    FontWeight =700
                    TabIndex =4
                    Name ="Command8"
                    Caption ="ت&ائيد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =215
                    AccessKey =228
                    TextFontCharSet =0
                    Left =630
                    Top =1592
                    Width =1133
                    Height =397
                    FontWeight =700
                    TabIndex =5
                    Name ="Command9"
                    Caption ="ا&نصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =165
                    Top =435
                    Width =2736
                    Height =300
                    TabIndex =1
                    Name ="GRP"
                    RowSourceType ="Value List"
                    RowSource ="1;\"مديران\";2;\"حسابداران\";3;\"بازرگاني\";4;\"انبار\";5;\"كارگزيني\";6;\"هتل\""
                    ColumnWidths ="0"
                    DefaultValue ="2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3015
                            Top =482
                            Width =735
                            Height =285
                            Name ="Label11"
                            Caption ="گروه كاري"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "users_chang.cls"
