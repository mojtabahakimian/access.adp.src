Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
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
    Width =10674
    DatasheetFontHeight =10
    ItemSuffix =47
    Left =2265
    Top =690
    Right =12945
    Bottom =4110
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbace0f5a1a4ce240
    End
    Caption ="دفتر معين"
    OnOpen ="[Event Procedure]"
    HelpFile ="dena.hlp"
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
            Width =4536
            Height =2835
        End
        Begin Section
            Height =3437
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    TextFontCharSet =0
                    Left =5403
                    Top =2948
                    Width =870
                    Height =315
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
                    Left =4535
                    Top =2948
                    Width =870
                    Height =317
                    FontWeight =700
                    TabIndex =2
                    Name ="Command6"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =3900
                    Top =1365
                    Width =2057
                    Height =373
                    Name ="BEDBES"
                    DefaultValue ="1"
                    Begin
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =4621
                            Top =1439
                            OptionValue =2
                            Name ="Option43"
                            Begin
                                Begin Label
                                    OverlapFlags =215
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =3901
                                    Top =1411
                                    Width =720
                                    Height =285
                                    Name ="Label44"
                                    Caption ="بستانكار"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            ReadingOrder =1
                            Left =5512
                            Top =1444
                            OptionValue =1
                            Name ="Option41"
                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    TextFontCharSet =0
                                    TextAlign =3
                                    Left =4985
                                    Top =1411
                                    Width =510
                                    Height =285
                                    Name ="Label42"
                                    Caption ="بدهكار"
                                End
                            End
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6031
                    Top =227
                    Width =1431
                    Height =300
                    TabIndex =3
                    Name ="Combo34"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.NUM"
                        "BER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes FROM"
                        " TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HE"
                        "S.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KO"
                        "L"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7547
                            Top =226
                            Width =840
                            Height =285
                            Name ="Label35"
                            Caption =" از حساب :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1077
                    Top =241
                    Width =4851
                    Height =300
                    TabIndex =4
                    Name ="Combo36"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.NUM"
                        "BER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDE"
                        "TA_HES.NAME FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.N"
                        "UMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER"
                        " = DETA_HES.N_KOL"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6034
                    Top =795
                    Width =1431
                    Height =300
                    TabIndex =5
                    Name ="Combo39"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.NUM"
                        "BER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes FROM"
                        " TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HE"
                        "S.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KO"
                        "L ORDER BY RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES"
                        ".NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar))"
                    AfterUpdate ="[Event Procedure]"
                    OnNotInList ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =7550
                            Top =795
                            Width =780
                            Height =285
                            Name ="Label40"
                            Caption =" تاحساب :"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =1080
                    Top =809
                    Width =4851
                    Height =300
                    TabIndex =6
                    Name ="Combo41"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT RTRIM(CAST(TDETA_HES.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.NUM"
                        "BER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDE"
                        "TA_HES.NAME FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.N"
                        "UMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER"
                        " = DETA_HES.N_KOL"
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    ColumnCount =2
                    Left =566
                    Top =2211
                    Width =8781
                    Height =317
                    TabIndex =7
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="visit"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.ROUTE_NAME + N' - ' + CUST_HESAB.NAME"
                        " + N' - ' + CUST_HESAB.hes AS Expr1 FROM Visit_route INNER JOIN CUST_HESAB ON Vi"
                        "sit_route.HES = CUST_HESAB.hes WHERE (Visit_route.RACTIVE = 1)"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =9425
                            Top =2211
                            Width =915
                            Height =240
                            ForeColor =8388608
                            Name ="Label26"
                            Caption ="مسير ويزيت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "F_MENU_KOL_MOIN_DATE_AI.cls"
