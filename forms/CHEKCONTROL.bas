Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11700
    DatasheetFontHeight =10
    ItemSuffix =8
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0x79bf0bd0923be340
    End
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin Section
            CanGrow = NotDefault
            Height =7140
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =60
                    Top =300
                    Width =11460
                    Name ="Child0"
                    SourceObject ="Form.FASNAD_WITHOUT_CHEK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6360
                            Width =5160
                            Height =240
                            ForeColor =26367
                            Name ="Label1"
                            Caption ="اسناد حسابداري دريافت چك كه چك آنها در دفتر چك موجود نيست"
                            ControlTipText ="شسيبلصسثيبشچ"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Top =2100
                    Width =11460
                    TabIndex =1
                    Name ="Child2"
                    SourceObject ="Form.FCHEK_W_PASS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4320
                            Top =1800
                            Width =7140
                            Height =240
                            ForeColor =26367
                            Name ="Label3"
                            Caption ="چكهاي دريافتي وصول شده يا واگذار شده يا برگشت شده كه سند بستانكار آن وجود ندارند"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Top =3900
                    Width =11460
                    TabIndex =2
                    Name ="Child4"
                    SourceObject ="Form.FCHEK_W_PASSP"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =4140
                            Top =3600
                            Width =7320
                            Height =240
                            ForeColor =26367
                            Name ="Label5"
                            Caption ="سندهاي بستانكار اسناد دريافتني كه چك آنها در دفتر چك وصولي يا واگذاري يا برگشتي "
                                "نخورده است"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Top =5700
                    Width =11460
                    TabIndex =3
                    Name ="Child6"
                    SourceObject ="Form.FCHEK_WITHOUT_SANAD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =6300
                            Top =5400
                            Width =5160
                            Height =240
                            ForeColor =26367
                            Name ="Label7"
                            Caption ="چكهاي دريافتي كه در اسناد وجود ندارند"
                        End
                    End
                End
            End
        End
    End
End
