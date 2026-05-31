Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4875
    RowHeight =375
    ItemSuffix =16
    Left =7710
    Top =45
    Right =14490
    Bottom =3390
    TimerInterval =5000
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xbb39346df6bde540
    End
    RecordSource ="SELECT     SUM(FACTOREFROOSH.MABL_K - FACTOREFROOSH.N_MOIN + FACTOREFROOSH.IMBAA"
        ") AS mab, dbo.UIIF(HEAD_LST.MODAT_PPID, N'=', 0, HEAD_LST.MAS, PRICE_PAYNO.MODAT"
        ") AS MODAT, FACTOREFROOSH.NUMBER, FACTOREFROOSH.TAG, HEAD_LST.DATE_N,dbo.Umonth("
        "HEAD_LST.DATE_N) AS MMO, HEAD_LST.CUST_NO FROM         HEAD_LST INNER JOIN  FACT"
        "OREFROOSH ON HEAD_LST.NUMBER = FACTOREFROOSH.NUMBER AND HEAD_LST.TAG - 11 = FACT"
        "OREFROOSH.TAG LEFT OUTER JOIN  PRICE_PAYNO ON HEAD_LST.MODAT_PPID = PRICE_PAYNO."
        "PPID WHERE (HEAD_LST.TAG = 13) GROUP BY dbo.UIIF(HEAD_LST.MODAT_PPID, N'=', 0, H"
        "EAD_LST.MAS, PRICE_PAYNO.MODAT), FACTOREFROOSH.NUMBER, FACTOREFROOSH.TAG, HEAD_L"
        "ST.DATE_N , dbo.UMonth(HEAD_LST.DATE_N), HEAD_LST.CUST_NO HAVING      (HEAD_LST."
        "CUST_NO = N'100-1-3')"
    Caption ="آمار فروش"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnTimer ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:DataAxisEmpty/>\015"
        "\012  <x:PivotField>\015\012   <x:Name>mab</x:Name>\015\012   <x:DataType>Number"
        "</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Com"
        "pareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFi"
        "eld>\015\012  <x:PivotField>\015\012   <x:Name>MODAT</x:Name>\015\012   <x:DataT"
        "ype>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </"
        "x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>NUMBER</x:Name>\015\012 "
        "  <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedTy"
        "pe>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>TAG</x:Name>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>DATE_N</"
        "x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adBig"
        "Int</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOr"
        "deredMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Na"
        "me>MMO</x:Name>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedT"
        "ype>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x"
        ":CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>CUST_NO</x:Name>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012"
        "   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015"
        "\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotView>\015\012   "
        "<x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:Caption>Microsoft Office Pi"
        "votTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012"
        "   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:Allow"
        "Additions>\015\012   <x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:Piv"
        "otView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>12594<"
        "/x:Width>\015\012  <x:Height>6403</x:Height>\015\012  <x:AllowPropertyBrowser/>\015"
        "\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataSourc"
        "e>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>0</x:D"
        "ataSourceIndex>\015\012  </x:Category>\015\012  <x:BoundCharts>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012  <x:PivotAg"
        "gOrientation>None</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012 "
        " <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Ent"
        "ry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00"
        "0000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</"
        "x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entr"
        "y>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:En"
        "try>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#6"
        "00080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080<"
        "/x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Ent"
        "ry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015"
        "\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012  "
        " <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:En"
        "try>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:Def"
        "aultFont>\015\012  <x:Border>\015\012   <x:Color>#2F4F4F</x:Color>\015\012  </x:"
        "Border>\015\012  <x:Interior>\015\012   <x:Color>#FFCC99</x:Color>\015\012   <x:"
        "FillEffect>\015\012    <x:fill x:type=\"Solid\" x:color=\"#FFCC99\"/>\015\012   "
        "</x:FillEffect>\015\012  </x:Interior>\015\012  <x:Chart>\015\012   <x:PlotArea>"
        "\015\012    <x:Interior>\015\012     <x:Color>#FFFFFF</x:Color>\015\012     <x:F"
        "illEffect>\015\012      <x:fill x:type=\"Solid\" x:color=\"#FFFFFF\"/>\015\012  "
        "   </x:FillEffect>\015\012    </x:Interior>\015\012    <x:Graph>\015\012     <x:"
        "SubType>Smooth</x:SubType>\015\012     <x:Type>Line</x:Type>\015\012     <x:SubT"
        "ype>Standard</x:SubType>\015\012     <x:Series>\015\012      <x:Line>\015\012   "
        "    <x:Color>#000000</x:Color>\015\012      </x:Line>\015\012      <x:FormatMap>"
        "\015\012      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <"
        "x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012      "
        " <x:Data>&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:I"
        "ndex>0</x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x"
        ":DataSourceIndex>\015\012      </x:Category>\015\012      <x:Marker>\015\012    "
        "   <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0<"
        "/x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:Smooth/>\015"
        "\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!<"
        "/x:Identifier>\015\012     </x:Series>\015\012     <x:Dimension>\015\012      <x"
        ":ScaleID>255615680</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012"
        "     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>255615884<"
        "/x:ScaleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015"
        "\012     <x:Dimension>\015\012      <x:ScaleID>255616088</x:ScaleID>\015\012    "
        "  <x:Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:Overl"
        "ap>100</x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:First"
        "SliceAngle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012"
        "     <x:AxisID>255619448</x:AxisID>\015\012     <x:ScaleID>255615680</x:ScaleID>"
        "\015\012     <x:Type>Category</x:Type>\015\012     <x:Font>\015\012      <x:Font"
        "Name>Tahoma</x:FontName>\015\012      <x:Size>7</x:Size>\015\012      <x:B>Autom"
        "atic</x:B>\015\012      <x:I>Automatic</x:I>\015\012      <x:U>Automatic</x:U>\015"
        "\012     </x:Font>\015\012     <x:Number>\015\012      <x:FormatString>mm/dd/yy<"
        "/x:FormatString>\015\012     </x:Number>\015\012     <x:MajorTick>Outside</x:Maj"
        "orTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Placement>Bot"
        "tom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    <"
        "/x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>255620156</x:AxisID>\015\012 "
        "    <x:ScaleID>255615884</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015\012 "
        "    <x:Font>\015\012      <x:Size>7</x:Size>\015\012      <x:B>Automatic</x:B>\015"
        "\012      <x:I>Automatic</x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x"
        ":Font>\015\012     <x:Number>\015\012      <x:FormatString>#,###</x:FormatString"
        ">\015\012     </x:Number>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012"
        "     <x:MinorTick>None</x:MinorTick>\015\012     <x:Placement>Left</x:Placement>"
        "\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identi"
        "fier>\015\012  </x:Chart>\015\012  <x:Title>\015\012   <x:Font>\015\012    <x:Fo"
        "ntName>B Titr</x:FontName>\015\012    <x:Size>7</x:Size>\015\012    <x:Color>#A5"
        "2A2A</x:Color>\015\012    <x:B>Automatic</x:B>\015\012    <x:I>Automatic</x:I>\015"
        "\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012   <x:Border>\015\012   "
        " <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>\015\012   <x:Caption>\015"
        "\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012    <x:Data>&quot;نمودا"
        "ر فروش 6 ماهه&quot;</x:Data>\015\012   </x:Caption>\015\012  </x:Title>\015\012 "
        " <x:Scaling>\015\012   <x:ScaleID>255615680</x:ScaleID>\015\012  </x:Scaling>\015"
        "\012  <x:Scaling>\015\012   <x:ScaleID>255615884</x:ScaleID>\015\012  </x:Scalin"
        "g>\015\012  <x:Scaling>\015\012   <x:ScaleID>255616088</x:ScaleID>\015\012  </x:"
        "Scaling>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideScreenTips/>\015\012 </x"
        ":ChartSpace>\015\012</xml>"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Section
            Height =3131
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1701
                    Top =1474
                    ColumnWidth =1575
                    ColumnOrder =4
                    Name ="DD"
                    ControlSource ="MODAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1470
                            Width =1470
                            Height =240
                            Name ="Label3"
                            Caption ="مدت پرداخت توافقي"
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
                    Left =1701
                    Top =1834
                    ColumnWidth =1365
                    ColumnOrder =0
                    TabIndex =1
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1834
                            Width =480
                            Height =240
                            Name ="Label5"
                            Caption ="تاريخ"
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
                    Left =283
                    Top =2324
                    ColumnWidth =1590
                    ColumnOrder =1
                    TabIndex =2
                    Name ="MABL_K"
                    ControlSource ="mab"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3010
                            Top =2324
                            Width =675
                            Height =240
                            Name ="Label8"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1701
                    ColumnWidth =435
                    ColumnOrder =2
                    TabIndex =3
                    Name ="MMO"
                    ControlSource ="MMO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =480
                            Height =240
                            Name ="Label10"
                            Caption ="ماه "
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =566
                    Top =850
                    ColumnOrder =3
                    TabIndex =4
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2783
                            Top =850
                            Width =1185
                            Height =240
                            Name ="Label14"
                            Caption ="شماره مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =453
                    Top =2891
                    ColumnWidth =1530
                    TabIndex =5
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2835
                            Top =2891
                            Width =1020
                            Height =240
                            Name ="Label15"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "NABZ-khareed_MOSHTARI.cls"
