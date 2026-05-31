Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5775
    RowHeight =292
    ItemSuffix =16
    Left =825
    Top =1500
    Right =11625
    Bottom =6795
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    Tag ="@DT =FORMS![F_MENU_DATE]![DT2]"
    OrderBy ="SFRAV_MOG.sudkol DESC"
    RecSrcDt = Begin
        0xef6c671014e1e440
    End
    InputParameters ="@DT =FORMS![F_MENU_DATE]![DT2]"
    RecordSource ="SFRAV_MOG"
    Caption ="سود ناويژه حاصل از فروش كالا ها"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,60 } Begin
        0x40004400540020003d0046004f0052004d00530021005b0046005f004d004500 ,
        0x4e0055005f0044004100540045005d0021005b004400540032005d00
    End
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
        Begin Section
            Height =3146
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1360
                    Top =170
                    Width =1440
                    Height =255
                    ColumnWidth =1185
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4161
                            Top =170
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1077
                    Top =456
                    Height =255
                    ColumnWidth =1095
                    ColumnOrder =2
                    TabIndex =1
                    Name ="P1"
                    ControlSource ="feekh"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4389
                            Top =453
                            Width =945
                            Height =240
                            Name ="Label3"
                            Caption ="ميانگين خريد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1077
                    Top =793
                    Height =255
                    ColumnWidth =1245
                    ColumnOrder =4
                    TabIndex =2
                    Name ="P2"
                    ControlSource ="feefr"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4389
                            Top =795
                            Width =1050
                            Height =240
                            Name ="Label5"
                            Caption ="ميانگين فروش"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1077
                    Top =1077
                    Height =255
                    ColumnWidth =1380
                    ColumnOrder =6
                    TabIndex =3
                    Name ="P3"
                    ControlSource ="sudone"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4387
                            Top =1084
                            Width =765
                            Height =240
                            Name ="Label7"
                            Caption ="سود  واحد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1077
                    Top =1421
                    Height =255
                    ColumnWidth =1665
                    ColumnOrder =7
                    TabIndex =4
                    Name ="P4"
                    ControlSource ="sudkol"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4389
                            Top =1417
                            Width =555
                            Height =405
                            Name ="Label9"
                            Caption ="سود كل"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1080
                    Top =1740
                    Height =255
                    ColumnWidth =1020
                    ColumnOrder =5
                    TabIndex =5
                    Name ="SumOfMEGHk"
                    ControlSource ="meghfr"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3602
                            Top =1740
                            Width =1050
                            Height =240
                            Name ="Label11"
                            Caption ="مقدار فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =2721
                    Width =3450
                    Height =255
                    ColumnWidth =4305
                    ColumnOrder =1
                    TabIndex =6
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4934
                            Top =2727
                            Width =540
                            Height =240
                            Name ="Label13"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1080
                    Top =2115
                    Height =255
                    ColumnWidth =960
                    ColumnOrder =3
                    TabIndex =7
                    Name ="Text14"
                    ControlSource ="meghkh"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3602
                            Top =2115
                            Width =1050
                            Height =240
                            Name ="Label15"
                            Caption ="مقدار  خريد"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "SOUD_FROOSH.cls"
