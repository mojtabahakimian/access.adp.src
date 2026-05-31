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
    Width =7500
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =4110
    Top =2985
    Right =11145
    Bottom =4665
    HelpContextId =17
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x095a33e52d41e240
    End
    Caption ="دريافت موجودي پايان سال قبل"
    HelpFile ="secret.hlp"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =1927
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =3727
                    Top =1475
                    Width =1095
                    Height =405
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="Command3"
                    Caption ="&تاييد"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontCharSet =0
                    Left =2632
                    Top =1475
                    Width =1095
                    Height =405
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="Command4"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =120
                    Top =915
                    Width =3119
                    Height =435
                    FontSize =16
                    FontWeight =700
                    ForeColor =128
                    Name ="YEA"
                    RowSourceType ="Value List"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =3360
                            Top =915
                            Width =3675
                            Height =345
                            Name ="Label1"
                            Caption ="سال مالي  مورد نظر جهت انتقال موجودي از آن:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Width =7500
                    Height =855
                    FontSize =20
                    ForeColor =255
                    Name ="Label14"
                    Caption ="اين عمليات فقط بايد پشت سرور انجام شود"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "GETFIRSTMOG.cls"
