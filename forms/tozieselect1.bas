Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =10
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14626
    ItemSuffix =25
    Left =3240
    Top =3255
    Right =17865
    Bottom =7770
    DatasheetGridlinesColor =12632256
    MaxRecords =100
    RecSrcDt = Begin
        0xbe1d93b0aeade240
    End
    Caption ="."
    DatasheetFontName ="Arabic Transparent"
    OnKeyPress ="[Event Procedure]"
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
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
        Begin ListBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            Height =1417
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
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =4535
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Top =573
                    Width =14580
                    Height =3795
                    Name ="tozieselect"
                    SourceObject ="Form.tozieselect"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12193
                    Top =120
                    Width =1296
                    Height =300
                    TabIndex =1
                    BackColor =13434828
                    Name ="dt1"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13585
                            Top =120
                            Width =645
                            Height =300
                            Name ="Label21"
                            Caption ="از تاريخ:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9075
                    Top =120
                    Width =1296
                    Height =300
                    TabIndex =2
                    BackColor =13408767
                    Name ="dt2"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="=farsidate(Now())"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =10467
                            Top =120
                            Width =645
                            Height =300
                            Name ="Label23"
                            Caption ="تا تاريخ:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =116
                    Top =127
                    Width =8160
                    Height =285
                    FontSize =10
                    ForeColor =16711680
                    Name ="Label24"
                    Caption ="براي انتخاب فاكتور روي آن اينتر را بزنيد تا به حواله اضافه گردد . ميتوانيد محدود"
                        "ه تاريخي را تغيير دهيد"
                End
            End
        End
    End
End
CodeBehindForm
' See "tozieselect1.cls"
