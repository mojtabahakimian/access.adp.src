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
    Width =9184
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =2940
    Top =2190
    Right =8970
    Bottom =4050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x41a653f29e1ae240
    End
    Caption =" ."
    DatasheetFontName ="Arial (Arabic)"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
        End
        Begin Line
            Width =1701
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =1870
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontCharSet =0
                    Left =6135
                    Top =1350
                    Width =1281
                    Height =418
                    FontWeight =700
                    TabIndex =2
                    Name ="Command1"
                    Caption ="&تائيد"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =26
                    Left =113
                    Width =915
                    Height =1020
                    FontSize =48
                    Name ="Label13"
                    Caption ="V"
                    FontName ="Webdings"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =223
                    Left =963
                    Top =56
                    Width =170
                    Height =227
                    Name ="Line14"
                End
                Begin Line
                    LineSlant = NotDefault
                    OverlapFlags =215
                    Left =1020
                    Top =283
                    Width =226
                    Height =113
                    Name ="Line15"
                End
                Begin Line
                    OverlapFlags =87
                    Left =1028
                    Top =505
                    Width =283
                    Height =57
                    Name ="Line16"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =225
                    TextFontCharSet =0
                    Left =4845
                    Top =1350
                    Width =1281
                    Height =418
                    FontWeight =700
                    TabIndex =3
                    Name ="Command17"
                    Caption ="&لغو"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    Left =4907
                    Top =113
                    Width =3645
                    Height =345
                    Name ="Label18"
                    Caption ="واگذاري چك  به حساب  ذيل انجام شود؟"
                    FontName ="Tahoma"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    AllowedText =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =3
                    Left =1140
                    Top =915
                    Width =6666
                    Height =300
                    TabIndex =1
                    Name ="HHMOIN"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0;5103;567"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7845
                            Top =915
                            Width =1275
                            Height =300
                            Name ="Label4"
                            Caption ="واگذار به حساب :"
                            FontName ="Tahoma"
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
                    Left =6570
                    Top =525
                    Width =1256
                    Height =315
                    Name ="DTS"
                    ValidationRule ="Not Is Null"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Date())"
                    FontName ="Tahoma"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7873
                            Top =529
                            Width =1035
                            Height =300
                            Name ="Label20"
                            Caption ="تاريخ واگذاري:"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "BEHESABDIALOG.cls"
