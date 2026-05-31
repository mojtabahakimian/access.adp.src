Version =20
VersionRequired =20
Begin Form
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5669
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =600
    Top =720
    Right =8055
    Bottom =4485
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x8620e0decf0ce340
    End
    InputParameters ="@Forms___PARA___ANBAR = [Forms]![PARA]![ANBAR] , @Forms___PARA___DT BIGINT= [For"
        "ms]![PARA]![DT] , @Forms___PARA___id = [Forms]![PARA]![id]  , @Forms___PARA___CO"
        "D  =  [Forms]![PARA]![COD]"
    RecordSource ="AVRAGE_MOG"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,372 } Begin
        0x400046006f0072006d0073005f005f005f0050004100520041005f005f005f00 ,
        0x41004e0042004100520020003d0020005b0046006f0072006d0073005d002100 ,
        0x5b0050004100520041005d0021005b0041004e004200410052005d0020002c00 ,
        0x2000400046006f0072006d0073005f005f005f0050004100520041005f005f00 ,
        0x5f0044005400200042004900470049004e0054003d0020005b0046006f007200 ,
        0x6d0073005d0021005b0050004100520041005d0021005b00440054005d002000 ,
        0x2c002000400046006f0072006d0073005f005f005f0050004100520041005f00 ,
        0x5f005f006900640020003d0020005b0046006f0072006d0073005d0021005b00 ,
        0x50004100520041005d0021005b00690064005d00200020002c00200040004600 ,
        0x6f0072006d0073005f005f005f0050004100520041005f005f005f0043004f00 ,
        0x4400200020003d00200020005b0046006f0072006d0073005d0021005b005000 ,
        0x4100520041005d0021005b0043004f0044005d00
    End
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
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
            Height =1714
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =906
                    Top =793
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3753
                            Top =793
                            Width =555
                            Height =240
                            Name ="Label0"
                            Caption ="CODE:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =906
                    Top =1133
                    TabIndex =1
                    Name ="MOG"
                    ControlSource ="MOG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3828
                            Top =1133
                            Width =480
                            Height =240
                            Name ="Label1"
                            Caption ="MOG:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =906
                    Top =1474
                    TabIndex =2
                    Name ="mabl"
                    ControlSource ="mabl"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3843
                            Top =1474
                            Width =465
                            Height =240
                            Name ="Label2"
                            Caption ="mabl:"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AVRAGE.cls"
