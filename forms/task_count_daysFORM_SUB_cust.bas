Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    AllowUpdating =4
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12642
    RowHeight =315
    ItemSuffix =6
    Left =135
    Top =2265
    Right =8325
    Bottom =9165
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x01e2118e3380e540
    End
    RecordSource ="SELECT COUNT(TASKS.COMP_COD) AS CountOfCOMP_COD, TASKS.STDATE, TASKS.COMP_COD, C"
        "UST_HESAB.NAME FROM TASKS INNER JOIN CUST_HESAB ON TASKS.COMP_COD = CUST_HESAB.h"
        "es GROUP BY TASKS.STDATE, TASKS.COMP_COD, CUST_HESAB.NAME ORDER BY COUNT(TASKS.C"
        "OMP_COD) DESC, TASKS.STDATE DESc"
    DatasheetFontName ="Tahoma"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3344
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9808
                    Top =396
                    ColumnWidth =1050
                    ColumnOrder =2
                    Name ="CountOfCOMP_COD"
                    ControlSource ="CountOfCOMP_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8107
                            Top =396
                            Width =1560
                            Height =240
                            Name ="Label0"
                            Caption ="تعداد  ورودي"
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
                    IMESentenceMode =3
                    Left =9808
                    Top =736
                    ColumnWidth =1020
                    ColumnOrder =0
                    TabIndex =1
                    Name ="STDATE"
                    ControlSource ="STDATE"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8107
                            Top =736
                            Width =705
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9808
                    Top =1076
                    ColumnWidth =5730
                    ColumnOrder =1
                    TabIndex =2
                    Name ="USERNAME"
                    ControlSource ="NAME"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8107
                            Top =1076
                            Width =945
                            Height =240
                            Name ="Label2"
                            Caption ="مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =1
                    IMESentenceMode =3
                    Left =5442
                    Top =1757
                    ColumnWidth =945
                    TabIndex =3
                    Name ="COMP_COD"
                    ControlSource ="COMP_COD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7854
                            Top =1757
                            Width =990
                            Height =240
                            Name ="Label5"
                            Caption ="كد مشتري"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1134
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =623
                    Name ="sc"
                    ControlSource ="=Sum([CountOfCOMP_COD])"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =623
                            Width =1560
                            Height =240
                            Name ="Label4"
                            Caption ="تعداد كار  ورودي"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "task_count_daysFORM_SUB_cust.cls"
