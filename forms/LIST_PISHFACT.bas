Version =20
VersionRequired =20
Begin Form
    KeyPreview = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    RowHeight =285
    ItemSuffix =7
    Left =270
    Top =615
    Right =8205
    Bottom =4590
    TimerInterval =15000
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    RecSrcDt = Begin
        0x9d4de5304f2be440
    End
    RecordSource ="SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_"
        "LST.OKF, dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, \015\012                      "
        "dbo.HEAD_LST.MOLAH\015\012FROM         dbo.HEAD_LST INNER JOIN\015\012          "
        "            dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes\015\012W"
        "HERE     (dbo.HEAD_LST.OKF = 0) AND (dbo.HEAD_LST.TAG = 20)"
    Caption ="ليست پيش فاكتورهاي تبديل نشده"
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnTimer ="[Event Procedure]"
    AllowFormView =0
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
        Begin Section
            Height =2961
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =680
                    ColumnWidth =675
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =680
                            Width =1020
                            Height =240
                            Name ="Label0"
                            Caption ="شماره "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =1360
                    ColumnWidth =1095
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =1360
                            Width =885
                            Height =240
                            Name ="Label2"
                            Caption ="تاريخ "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =2040
                    ColumnWidth =1260
                    TabIndex =2
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2160
                            Top =2040
                            Width =885
                            Height =240
                            Name ="Label4"
                            Caption ="كد مشتري "
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =2381
                    ColumnWidth =5565
                    TabIndex =3
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =2381
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3855
                    Top =2721
                    ColumnWidth =3750
                    TabIndex =4
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2154
                            Top =2721
                            Width =750
                            Height =240
                            Name ="Label6"
                            Caption ="ملاحظات"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "LIST_PISHFACT.cls"
