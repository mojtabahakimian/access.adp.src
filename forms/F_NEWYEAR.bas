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
    Width =14400
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =3555
    Top =2580
    Right =17955
    Bottom =7215
    HelpContextId =17
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x095a33e52d41e240
    End
    Caption ="ايجاد سال"
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
        Begin Section
            Height =4648
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9503
                    Top =2635
                    Width =1829
                    Height =435
                    FontSize =16
                    Name ="YEA"
                    Format ="#"
                    ValidationRule ="Not Is Null"
                    BeforeUpdate ="[Event Procedure]"
                    DefaultValue ="=[Forms]![Baseknow]![YEA]+1"
                    FontName ="WeblogmaYekan"
                    InputMask ="####"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =11399
                            Top =2640
                            Width =1530
                            Height =435
                            FontSize =12
                            Name ="Label1"
                            Caption ="سال مالي جديد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontFamily =2
                    Left =3192
                    Top =3975
                    Width =1485
                    Height =525
                    FontSize =12
                    TabIndex =1
                    ForeColor =0
                    Name ="Command3"
                    Caption ="&تاييد"
                    OnClick ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    AccessKey =199
                    TextFontFamily =2
                    Left =1692
                    Top =3971
                    Width =1485
                    Height =525
                    FontSize =12
                    TabIndex =2
                    ForeColor =0
                    Name ="Command4"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =675
                    Top =915
                    Width =10634
                    Height =435
                    FontSize =16
                    TabIndex =3
                    BackColor =13434828
                    Name ="opath"
                    ValidationRule ="Not Is Null"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =11363
                            Top =915
                            Width =2820
                            Height =435
                            FontSize =12
                            Name ="Label3"
                            Caption ="مسير ديتا بيس در سال جاري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6795
                    Top =1425
                    Width =4499
                    Height =435
                    FontSize =16
                    TabIndex =4
                    BackColor =13434828
                    Name ="odbs"
                    ValidationRule ="Not Is Null"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =11374
                            Top =1408
                            Width =2625
                            Height =435
                            FontSize =12
                            Name ="Label5"
                            Caption ="نام ديتا بيس در سال جاري:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =675
                    Top =3135
                    Width =10649
                    Height =435
                    FontSize =16
                    TabIndex =5
                    Name ="npathf"
                    ValidationRule ="Not Is Null"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =11385
                            Top =3120
                            Width =3015
                            Height =435
                            FontSize =12
                            Name ="Label7"
                            Caption ="مسير ديتا بيس براي سال جديد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6795
                    Top =3645
                    Width =4514
                    Height =435
                    FontSize =16
                    TabIndex =6
                    Name ="ndbs"
                    ValidationRule ="Not Is Null"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =11385
                            Top =3615
                            Width =2895
                            Height =435
                            FontSize =12
                            Name ="Label9"
                            Caption ="نام ديتا بيس  براي سال جديد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =240
                    Top =3165
                    Width =397
                    Height =340
                    TabIndex =7
                    Name ="Command10"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =675
                    Top =1935
                    Width =10649
                    Height =435
                    FontSize =16
                    TabIndex =8
                    BackColor =10092543
                    Name ="filesf"
                    ValidationRule ="Not Is Null"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =11385
                            Top =1950
                            Width =2670
                            Height =435
                            FontSize =12
                            Name ="Label12"
                            Caption ="مسير ديتا بيس خالي(Files):"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =240
                    Top =1965
                    Width =397
                    Height =340
                    TabIndex =9
                    Name ="Command13"
                    Caption ="..."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =2
                    Left =3118
                    Width =7500
                    Height =855
                    FontSize =26
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
' See "F_NEWYEAR.cls"
