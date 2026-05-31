Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10674
    RowHeight =450
    ItemSuffix =5
    Left =3495
    Top =3450
    Right =17820
    Bottom =6960
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xcedf84225411e640
    End
    RecordSource ="SELECT HEAD_LST.NUMBER, HEAD_LST.DATE_N, CUST_HESAB.NAME, CUST_HESAB.hes, CUST_H"
        "ESAB.ADDRESS FROM CUST_HESAB INNER JOIN HEAD_LST ON CUST_HESAB.hes = HEAD_LST.CU"
        "ST_NO WHERE (HEAD_LST.DATE_N BETWEEN 14020703 AND 14020703) AND (HEAD_LST.TAG = "
        "2) ORDER BY HEAD_LST.NUMBER"
    Caption ="انتخاب مسير"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
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
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =3288
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2267
                    Top =1870
                    ColumnWidth =1065
                    ColumnOrder =0
                    TabIndex =1
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5010
                            Top =1875
                            Width =960
                            Height =240
                            Name ="Label0"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2380
                    Top =2154
                    ColumnWidth =990
                    ColumnOrder =1
                    TabIndex =2
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5295
                            Top =2160
                            Width =495
                            Height =240
                            Name ="Label1"
                            Caption ="تاريخ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1587
                    Top =2607
                    ColumnOrder =2
                    TabIndex =3
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4365
                            Top =2610
                            Width =840
                            Height =240
                            Name ="Label2"
                            Caption ="كد مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =8781
                    Height =317
                    ColumnWidth =3615
                    ColumnOrder =3
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8859
                            Width =1815
                            Height =330
                            ForeColor =8388608
                            Name ="Label26"
                            Caption ="نام"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Top =450
                    Width =8781
                    Height =317
                    ColumnWidth =6750
                    TabIndex =4
                    BackColor =13434828
                    ForeColor =8388608
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =0
                            Left =8859
                            Top =450
                            Width =1815
                            Height =330
                            ForeColor =8388608
                            Name ="Label4"
                            Caption ="آدرس"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "tozieselect.cls"
