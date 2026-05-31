Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
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
    Width =7029
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =4110
    Top =2850
    Right =11145
    Bottom =4935
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x358ea6b70e6fe240
    End
    Caption ="تعداد كالا"
    DatasheetFontName ="Arial (Arabic)"
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
            Height =2097
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =187
                    Top =56
                    Width =1731
                    Height =765
                    FontSize =20
                    Name ="megh"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            TextFontFamily =10
                            Left =2040
                            Top =60
                            Width =4965
                            Height =750
                            FontSize =20
                            Name ="Label1"
                            Caption ="تعداد كالا بر حسب واحد اصلي  :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =3645
                    Top =1474
                    Width =870
                    Height =345
                    FontWeight =700
                    TabIndex =1
                    Name ="Command5"
                    Caption ="تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2777
                    Top =1474
                    Width =870
                    Height =347
                    FontWeight =700
                    TabIndex =2
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "VISTEDAD.cls"
