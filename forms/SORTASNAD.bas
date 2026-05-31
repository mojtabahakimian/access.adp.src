Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
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
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =2
    Left =11520
    Top =3975
    Right =17190
    Bottom =6000
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x803b8a68b86ce240
    End
    Caption ="مرتب سازي اسناد"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
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
            Height =2040
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =2553
                    Top =1485
                    Width =1260
                    Height =465
                    FontWeight =700
                    ForeColor =0
                    Name ="Command2"
                    Caption ="انجام عمليات"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="اين گزينه را جهت دريافت  اطلاعات از ديسك  تاييد نماييد"
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =340
                    Top =226
                    Width =4740
                    Height =855
                    FontWeight =700
                    ForeColor =255
                    Name ="Label9"
                    Caption ="شماره اسناد در انجام اين عمل تغيير مي كند ولي شماره مبنا ثابت مي ماند . حتما دقت"
                        " كنيد كه تمام پنجره ها بسته باشد و كاربر ديگري از برنامه استفاده نمي كند."
                    FontName ="Titr"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =1260
                    Top =1485
                    Width =1290
                    Height =465
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="Command1"
                    Caption ="انصراف "
                    OnClick ="[Event Procedure]"
                    ControlTipText ="اين گزينه را جهت دريافت  اطلاعات از ديسك  تاييد نماييد"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextFontCharSet =2
                    TextFontFamily =10
                    BackStyle =0
                    Left =367
                    Top =1215
                    Width =4743
                    Height =180
                    FontSize =6
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text0"
                    DefaultValue ="\"n\""
                    FontName ="Wingdings"
                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =360
                    Top =1216
                    Width =4743
                    Height =185
                    Name ="Box7"
                End
            End
        End
    End
End
CodeBehindForm
' See "SORTASNAD.cls"
