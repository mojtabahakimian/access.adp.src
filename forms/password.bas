Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4308
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =11970
    Top =4275
    Right =16530
    Bottom =6270
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x23e774f29e1ae240
    End
    Caption ="كلمه عبــور"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
        Begin Section
            Height =1757
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =510
                    Top =340
                    Height =345
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Text0"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =2
                            Left =2407
                            Top =344
                            Width =1815
                            Height =390
                            FontWeight =700
                            Name ="Label1"
                            Caption ="كلمه  عبور را وارد كنيد"
                            FontName ="Titr"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =202
                    TextFontFamily =2
                    Left =850
                    Top =1191
                    Width =1071
                    Height =343
                    FontWeight =700
                    TabIndex =1
                    Name ="Command2"
                    Caption ="&تاييد"
                    OnClick ="[Event Procedure]"
                    FontName ="Titr"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =199
                    TextFontFamily =2
                    Left =2273
                    Top =1191
                    Width =1071
                    Height =343
                    FontWeight =700
                    TabIndex =2
                    Name ="Command3"
                    Caption ="&انصراف"
                    OnClick ="[Event Procedure]"
                    FontName ="Titr"
                End
            End
        End
    End
End
CodeBehindForm
' See "password.cls"
