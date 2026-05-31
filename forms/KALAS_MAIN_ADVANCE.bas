Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    FilterOn = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7635
    RowHeight =300
    ItemSuffix =136
    Left =4890
    Top =1110
    Right =15285
    Bottom =6150
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000000
    RecSrcDt = Begin
        0x8f414fba50f7e540
    End
    Caption ="ليست گردشها ي كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>مقدار كل كالا</x:Name>\015\012   <x:SourceName>MEGHk</x:SourceName>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:NumberFormat>Fixed</x:NumberFormat>\015\012   <x:CompareOr"
        "deredMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>مبلغ كل</x:Name>\015\012   <x:SourceName>"
        "MABL_K</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Enc"
        "odedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>#,###</x:NumberForma"
        "t>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>مبلغ خالص</x:Nam"
        "e>\015\012   <x:SourceName>KHFR</x:SourceName>\015\012   <x:DataType>Number</x:D"
        "ataType>\015\012   <x:PLDataOrientation/>\015\012   <x:EncodedType>adDouble</x:E"
        "ncodedType>\015\012   <x:NumberFormat>#,###</x:NumberFormat>\015\012   <x:Compar"
        "eOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField"
        ">\015\012  <x:PivotField>\015\012   <x:Name>مبلغ تخفيف</x:Name>\015\012   <x:Sou"
        "rceName>N_MOIN</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>#,###</x:Num"
        "berFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMemb"
        "ersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ارزش ا"
        "فزوده</x:Name>\015\012   <x:SourceName>IMBAA</x:SourceName>\015\012   <x:DataTyp"
        "e>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012 "
        "  <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x"
        ":PivotField>\015\012  <x:PivotField>\015\012   <x:Name>واحد</x:Name>\015\012   <"
        "x:SourceName>DEPATMAN</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015"
        "\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembe"
        "rsBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x"
        ":PivotField>\015\012   <x:Name>ماه</x:Name>\015\012   <x:SourceName>MM</x:Source"
        "Name>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:Orientation>Row</x"
        ":Orientation>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:Co"
        "mpareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expand"
        "ed/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:"
        "Name>\015\012   <x:Orientation>Row</x:Orientation>\015\012   <x:Position>-1</x:P"
        "osition>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotFiel"
        "d>\015\012   <x:Name>Sum of مبلغ خالص</x:Name>\015\012   <x:PLName>Total1</x:PLN"
        "ame>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x"
        ":Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x:ParentField>مبلغ"
        " خالص</x:ParentField>\015\012   <x:NumberFormat>#,###</x:NumberFormat>\015\012  "
        "</x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:Top>\015\012   <x"
        ":TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOffset>"
        "0</x:LeftOffset>\015\012   <x:SeqNum>116</x:SeqNum>\015\012  </x:PivotData>\015\012"
        "  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:C"
        "aption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015"
        "\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:Allo"
        "wAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDe"
        "letions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>36036<"
        "/x:Width>\015\012  <x:Height>12303</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012"
        "  <x:PivotAggOrientation>Categories</x:PivotAggOrientation>\015\012  <x:HidePivo"
        "tFields/>\015\012  <x:NoGrouping/>\015\012  <x:NoFiltering/>\015\012  <x:Palette"
        ">\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015"
        "\012   <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012  "
        " <x:Entry>#600080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:En"
        "try>#008080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#0"
        "00080</x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF<"
        "/x:Entry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Ent"
        "ry>\015\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015"
        "\012   <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>"
        "Arial</x:DefaultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Bo"
        "rder>\015\012     <x:LineStyle>Solid</x:LineStyle>\015\012    </x:Border>\015\012"
        "    <x:Interior>\015\012     <x:Color>#F0FFFF</x:Color>\015\012     <x:FillEffec"
        "t>\015\012      <x:fill x:type=\"Solid\" x:color=\"#F0FFFF\"/>\015\012     </x:F"
        "illEffect>\015\012    </x:Interior>\015\012    <x:Graph>\015\012     <x:Type>Lin"
        "e</x:Type>\015\012     <x:SubType>Standard</x:SubType>\015\012     <x:SubType>Ma"
        "rker</x:SubType>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      "
        "</x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Caption>\015"
        "\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot"
        ";Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</x:Inde"
        "x>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceInd"
        "ex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourc"
        "eIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:DataLabels>"
        "\015\012       <x:Border>\015\012        <x:ColorIndex>None</x:ColorIndex>\015\012"
        "       </x:Border>\015\012       <x:Number>\015\012        <x:FormatString>#,###"
        "</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowValue/>\015\012"
        "       <x:Separator>;</x:Separator>\015\012       <x:Position>Top</x:Position>\015"
        "\012      </x:DataLabels>\015\012      <x:Explode>0</x:Explode>\015\012      <x:"
        "Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>"
        "\015\012      <x:Identifier>!</x:Identifier>\015\012     </x:Series>\015\012    "
        " <x:Dimension>\015\012      <x:ScaleID>314050724</x:ScaleID>\015\012      <x:Ind"
        "ex>Categories</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015"
        "\012      <x:ScaleID>314050944</x:ScaleID>\015\012      <x:Index>Value</x:Index>"
        "\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>31"
        "4051164</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015\012     </x:"
        "Dimension>\015\012     <x:Overlap>100</x:Overlap>\015\012     <x:GapWidth>150</x"
        ":GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    </x:G"
        "raph>\015\012    <x:Axis>\015\012     <x:AxisID>314078872</x:AxisID>\015\012    "
        " <x:ScaleID>314050724</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012 "
        "    <x:Line>\015\012      <x:Weight>0</x:Weight>\015\012     </x:Line>\015\012  "
        "   <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick"
        ">\015\012     <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto"
        "</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID"
        ">314080108</x:AxisID>\015\012     <x:ScaleID>314050944</x:ScaleID>\015\012     <"
        "x:Type>Value</x:Type>\015\012     <x:MajorGridlines>\015\012      <x:Line>\015\012"
        "       <x:Weight>0</x:Weight>\015\012       <x:LineJoin>JoinRound</x:LineJoin>\015"
        "\012      </x:Line>\015\012     </x:MajorGridlines>\015\012     <x:MajorTick>Out"
        "side</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Pl"
        "acement>Left</x:Placement>\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012 "
        "  <x:Identifier></x:Identifier>\015\012  </x:Chart>\015\012  <x:Legend>\015\012 "
        "  <x:Placement>Right</x:Placement>\015\012  </x:Legend>\015\012  <x:Scaling>\015"
        "\012   <x:ScaleID>314050724</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scalin"
        "g>\015\012   <x:ScaleID>314050944</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:"
        "Scaling>\015\012   <x:ScaleID>314051164</x:ScaleID>\015\012  </x:Scaling>\015\012"
        "  <x:HidePassiveAlert/>\015\012  <x:HideScreenTips/>\015\012 </x:ChartSpace>\015"
        "\012</xml>"
    AllowFormView =0
    OnViewChange ="[Event Procedure]"
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
            Height =31654
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =113
                    Width =2268
                    ColumnWidth =1125
                    ColumnOrder =3
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =113
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره برگه"
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
                    Left =1929
                    Top =453
                    Width =2268
                    ColumnWidth =1410
                    ColumnOrder =1
                    TabIndex =1
                    Name ="BARGAH"
                    ControlSource ="BARGAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =450
                            Width =735
                            Height =240
                            Name ="Label3"
                            Caption ="نوع برگه"
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
                    Left =1929
                    Top =793
                    Width =2268
                    ColumnWidth =1890
                    ColumnOrder =56
                    TabIndex =2
                    Name ="ANBNAME"
                    ControlSource ="ANBNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =793
                            Width =570
                            Height =240
                            Name ="Label5"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =1133
                    Width =2268
                    ColumnWidth =1650
                    ColumnOrder =57
                    TabIndex =3
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =1133
                            Width =1530
                            Height =240
                            Name ="Label7"
                            Caption ="شماره فاكتور برگشت"
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
                    Left =1929
                    Top =1474
                    Width =2268
                    ColumnWidth =1080
                    ColumnOrder =0
                    TabIndex =4
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    OnDblClick ="[Event Procedure]"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =1474
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="تاريخ برگه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =2010
                    Width =2268
                    ColumnWidth =990
                    ColumnOrder =58
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =2010
                            Width =885
                            Height =240
                            Name ="Label11"
                            Caption ="شماره سند"
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
                    Left =1935
                    Top =8850
                    Width =2268
                    ColumnWidth =4305
                    ColumnOrder =2
                    TabIndex =6
                    Name ="CUSTNAME"
                    ControlSource ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4770
                            Top =8850
                            Width =1185
                            Height =240
                            Name ="Label13"
                            Caption ="نام شخص"
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
                    Left =1929
                    Top =2577
                    Width =2268
                    ColumnWidth =4260
                    ColumnOrder =59
                    TabIndex =7
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4770
                            Top =2580
                            Width =1110
                            Height =240
                            Name ="Label15"
                            Caption ="توضيحات فاكتور"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2042
                    Top =2891
                    Width =2268
                    ColumnWidth =1770
                    ColumnOrder =60
                    TabIndex =8
                    Name ="FNUMCO"
                    ControlSource ="FNUMCO"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4877
                            Top =2891
                            Width =1620
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور فروشنده"
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
                    Left =1929
                    Top =3514
                    Width =2268
                    ColumnWidth =1035
                    ColumnOrder =34
                    TabIndex =9
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =3514
                            Width =675
                            Height =240
                            Name ="Label21"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =3855
                    Width =2268
                    ColumnWidth =1140
                    ColumnOrder =11
                    TabIndex =10
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =3855
                            Width =960
                            Height =240
                            Name ="Label23"
                            Caption ="مقدار كل كالا"
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
                    Left =1929
                    Top =4195
                    Width =2268
                    ColumnWidth =1245
                    ColumnOrder =61
                    TabIndex =11
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MAR"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =4195
                            Width =1110
                            Height =240
                            Name ="Label25"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =4535
                    Width =2268
                    ColumnWidth =930
                    ColumnOrder =12
                    TabIndex =12
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =4535
                            Width =375
                            Height =240
                            Name ="Label27"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =4875
                    Width =2268
                    ColumnWidth =4605
                    ColumnOrder =5
                    TabIndex =13
                    Name ="KALA"
                    ControlSource ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =4875
                            Width =1305
                            Height =240
                            Name ="Label29"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1929
                    Top =5215
                    Width =2268
                    ColumnWidth =1905
                    ColumnOrder =54
                    TabIndex =14
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =5215
                            Width =615
                            Height =240
                            Name ="Label31"
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
                    Left =1929
                    Top =6236
                    Width =2268
                    ColumnWidth =735
                    ColumnOrder =35
                    TabIndex =15
                    Name ="VAHEDNAME"
                    ControlSource ="VAHEDNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =6236
                            Width =630
                            Height =240
                            Name ="Label37"
                            Caption ="نام واحد"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1985
                    Top =2267
                    Width =2268
                    ColumnWidth =615
                    ColumnOrder =4
                    TabIndex =17
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4820
                            Top =2267
                            Width =555
                            Height =240
                            Name ="Label40"
                            Caption ="كد كالا"
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
                    Left =2098
                    Top =3174
                    Width =2268
                    ColumnWidth =1417
                    ColumnOrder =63
                    TabIndex =18
                    Name ="hes"
                    ControlSource ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4933
                            Top =3178
                            Width =750
                            Height =240
                            Name ="Label43"
                            Caption ="كد شخص"
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
                    Left =1985
                    Top =6973
                    Width =2268
                    ColumnWidth =1417
                    ColumnOrder =64
                    TabIndex =19
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4820
                            Top =6973
                            Width =1035
                            Height =240
                            Name ="Label44"
                            Caption ="نام كاربر"
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
                    Left =1985
                    Top =7313
                    Width =2268
                    ColumnWidth =1417
                    ColumnOrder =65
                    TabIndex =20
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4820
                            Top =7313
                            Width =600
                            Height =240
                            Name ="Label45"
                            Caption ="شيفت"
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
                    Left =1985
                    Top =7653
                    Width =2268
                    ColumnWidth =1417
                    ColumnOrder =66
                    TabIndex =21
                    Name ="CUSTKNAME"
                    ControlSource ="CUSTKNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4820
                            Top =7653
                            Width =960
                            Height =240
                            Name ="Label46"
                            Caption ="نوع مشتري"
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
                    Left =1985
                    Top =7994
                    Width =2268
                    ColumnWidth =1417
                    ColumnOrder =67
                    TabIndex =22
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4820
                            Top =7994
                            Width =1140
                            Height =240
                            Name ="Label47"
                            Caption ="واحد"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =215
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3633
                    Top =8505
                    Width =852
                    ColumnWidth =1417
                    ColumnOrder =68
                    TabIndex =23
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            ReadingOrder =2
                            NumeralShapes =2
                            Left =2880
                            Top =8505
                            Width =1095
                            Height =240
                            Name ="Label48"
                            Caption ="ملاحظات سطر"
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
                    Left =2552
                    Top =9127
                    ColumnWidth =1417
                    ColumnOrder =69
                    TabIndex =24
                    Name ="MAS"
                    ControlSource ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5504
                            Top =9127
                            Width =450
                            Height =240
                            Name ="Label49"
                            Caption ="مدت"
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
                    Left =2495
                    Top =9474
                    ColumnWidth =1417
                    ColumnOrder =70
                    TabIndex =25
                    Name ="N_RASID"
                    ControlSource ="N_RASID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5448
                            Top =9467
                            Width =885
                            Height =240
                            Name ="Label51"
                            Caption ="محل مصرف"
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
                    Left =2495
                    Top =9841
                    ColumnWidth =1417
                    ColumnOrder =55
                    TabIndex =26
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5448
                            Top =9834
                            Width =885
                            Height =240
                            Name ="Label53"
                            Caption ="شماره فني"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2495
                    Top =5669
                    ColumnWidth =795
                    ColumnOrder =71
                    TabIndex =27
                    Name ="SHARAYET"
                    ControlSource ="SHARAYET"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4982
                            Top =5669
                            Width =915
                            Height =240
                            Name ="Label54"
                            Caption ="ملاحظات"
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
                    Left =2438
                    Top =11111
                    ColumnWidth =1417
                    ColumnOrder =14
                    TabIndex =28
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5280
                            Top =11115
                            Width =900
                            Height =240
                            Name ="Label55"
                            Caption ="ارزش افزوده"
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
                    Left =2438
                    Top =11451
                    ColumnWidth =1515
                    ColumnOrder =15
                    TabIndex =29
                    Name ="HMBAA"
                    ControlSource ="HMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5175
                            Top =11445
                            Width =1410
                            Height =240
                            Name ="Label56"
                            Caption ="حساب ارزش افزوده"
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
                    Left =2438
                    Top =11792
                    ColumnWidth =690
                    ColumnOrder =16
                    TabIndex =30
                    Name ="TAMIR"
                    ControlSource ="TAMIR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5225
                            Top =11792
                            Width =615
                            Height =240
                            Name ="Label57"
                            Caption ="بارگيري"
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
                    Left =2438
                    Top =12132
                    ColumnWidth =1605
                    ColumnOrder =17
                    TabIndex =31
                    Name ="TICMBAA"
                    ControlSource ="TICMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5025
                            Top =12135
                            Width =1500
                            Height =240
                            Name ="Label58"
                            Caption ="محاسبه ارزش افزوده"
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
                    Left =2438
                    Top =12472
                    ColumnWidth =795
                    ColumnOrder =18
                    TabIndex =32
                    Name ="OKF"
                    ControlSource ="OKF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5400
                            Top =12465
                            Width =690
                            Height =240
                            Name ="Label59"
                            Caption ="تاييد برگه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =13950
                    ColumnWidth =3300
                    ColumnOrder =19
                    TabIndex =33
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4977
                            Top =13950
                            Width =600
                            Height =240
                            Name ="Label60"
                            Caption ="توضيح"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =14290
                    ColumnWidth =1417
                    ColumnOrder =6
                    TabIndex =34
                    Name ="B_SEF"
                    ControlSource ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4990
                            Top =14292
                            Width =1020
                            Height =240
                            Name ="Label61"
                            Caption ="بهينه سفارش"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =14630
                    ColumnWidth =1417
                    ColumnOrder =7
                    TabIndex =35
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4977
                            Top =14630
                            Width =600
                            Height =240
                            Name ="Label62"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =14971
                    ColumnWidth =1417
                    ColumnOrder =8
                    TabIndex =36
                    Name ="MIN_M"
                    ControlSource ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4947
                            Top =14971
                            Width =630
                            Height =240
                            Name ="Label63"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =15311
                    ColumnWidth =1417
                    ColumnOrder =9
                    TabIndex =37
                    Name ="MAX_M"
                    ControlSource ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4917
                            Top =15311
                            Width =660
                            Height =240
                            Name ="Label64"
                            Caption ="حد اكثر موجودي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =15651
                    ColumnWidth =1417
                    ColumnOrder =20
                    TabIndex =38
                    Name ="RADAH"
                    ControlSource ="RADAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4917
                            Top =15651
                            Width =660
                            Height =240
                            Name ="Label65"
                            Caption ="گروه"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =15991
                    ColumnWidth =1417
                    ColumnOrder =21
                    TabIndex =39
                    Name ="KINDK"
                    ControlSource ="KINDK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4992
                            Top =15991
                            Width =585
                            Height =240
                            Name ="Label66"
                            Caption ="نوع"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =16331
                    ColumnWidth =1417
                    ColumnOrder =22
                    TabIndex =40
                    Name ="MABL_F"
                    ControlSource ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4875
                            Top =16335
                            Width =735
                            Height =240
                            Name ="Label67"
                            Caption ="في عمده"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =17012
                    ColumnWidth =1417
                    ColumnOrder =24
                    TabIndex =42
                    Name ="CMBAA"
                    ControlSource ="CMBAA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4917
                            Top =17012
                            Width =660
                            Height =240
                            Name ="Label69"
                            Caption ="مشمول ماليات"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2175
                    Top =17352
                    ColumnWidth =780
                    ColumnOrder =25
                    TabIndex =43
                    Name ="VAZN"
                    ControlSource ="VAZN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5052
                            Top =17352
                            Width =675
                            Height =240
                            Name ="Label70"
                            Caption ="وزن "
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
                    Left =2438
                    Top =17801
                    ColumnWidth =1417
                    ColumnOrder =26
                    TabIndex =44
                    Name ="N_TAF"
                    ControlSource ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5225
                            Top =17801
                            Width =615
                            Height =240
                            Name ="Label71"
                            Caption ="price"
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
                    Left =2438
                    Top =18141
                    ColumnWidth =1417
                    ColumnOrder =27
                    TabIndex =45
                    Name ="TOTALARZ"
                    ControlSource ="TOTALARZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4925
                            Top =18141
                            Width =915
                            Height =240
                            Name ="Label72"
                            Caption ="total price"
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
                    Left =2268
                    Top =13096
                    ColumnWidth =1417
                    ColumnOrder =28
                    TabIndex =46
                    Name ="ANBARF"
                    ControlSource ="ANBARF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4815
                            Top =13095
                            Width =1140
                            Height =240
                            Name ="Label73"
                            Caption ="شماره صادراتي"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2835
                    Top =10601
                    ColumnWidth =1635
                    ColumnOrder =29
                    TabIndex =47
                    Name ="TAGCODE"
                    ControlSource ="TAGCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, BARGAH FROM TAGCOD"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5382
                            Top =10601
                            Width =855
                            Height =240
                            Name ="Label74"
                            Caption ="نوع برگه"
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
                    Left =2325
                    Top =18595
                    ColumnWidth =1417
                    ColumnOrder =30
                    TabIndex =48
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5106
                            Top =18600
                            Width =930
                            Height =240
                            Name ="Label75"
                            Caption ="درصد تخفيف"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2325
                    Top =18935
                    ColumnWidth =1417
                    ColumnOrder =73
                    TabIndex =49
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4986
                            Top =18930
                            Width =870
                            Height =240
                            Name ="Label76"
                            Caption ="مبلغ تخفيف"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2325
                    Top =19502
                    ColumnWidth =1417
                    ColumnOrder =33
                    TabIndex =50
                    Name ="MM"
                    ControlSource ="MM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4954
                            Top =19502
                            Width =375
                            Height =240
                            Name ="Label77"
                            Caption ="ماه"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =3062
                    Top =5896
                    Height =255
                    ColumnWidth =1417
                    ColumnOrder =31
                    TabIndex =51
                    Name ="VAHCODE"
                    ControlSource ="VAHCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5804
                            Top =5896
                            Width =660
                            Height =240
                            Name ="Label78"
                            Caption ="كد واحد"
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
                    IMESentenceMode =3
                    Left =1928
                    Top =19904
                    ColumnWidth =1905
                    ColumnOrder =13
                    TabIndex =52
                    Name ="KHFR"
                    ControlSource ="KHFR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4811
                            Top =19899
                            Width =810
                            Height =240
                            Name ="Label79"
                            Caption ="مبلغ خالص"
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
                    IMESentenceMode =3
                    Left =1928
                    Top =20245
                    ColumnWidth =1440
                    ColumnOrder =74
                    TabIndex =53
                    Name ="GHFR"
                    ControlSource ="GHFR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4796
                            Top =20244
                            Width =900
                            Height =240
                            Name ="Label80"
                            Caption ="قابل پرداخت"
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
                    Left =2608
                    Top =20585
                    ColumnWidth =1417
                    ColumnOrder =75
                    TabIndex =54
                    Name ="AVRAGE"
                    ControlSource ="AVRAGE"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =5276
                            Top =20589
                            Width =780
                            Height =405
                            Name ="Label81"
                            Caption ="في ميانگين"
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
                    Left =2608
                    Top =20925
                    ColumnWidth =1590
                    ColumnOrder =32
                    TabIndex =55
                    Name ="mabrial"
                    ControlSource ="mabrial"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =5366
                            Top =20919
                            Width =705
                            Height =405
                            Name ="Label82"
                            Caption ="مبلغ ريالي"
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
                    Left =2949
                    Top =21429
                    ColumnWidth =1417
                    ColumnOrder =76
                    TabIndex =56
                    Name ="ANBARCODE"
                    ControlSource ="ANBARCODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5751
                            Top =21429
                            Width =600
                            Height =240
                            Name ="Label83"
                            Caption ="كد انبار"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2175
                    Top =16671
                    ColumnWidth =1417
                    ColumnOrder =23
                    TabIndex =41
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPATMAN, DEPNAME FROM DEPART"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4857
                            Top =16671
                            Width =720
                            Height =240
                            Name ="Label68"
                            Caption ="واحد ارائه كننده"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1929
                    Top =6576
                    Width =2268
                    ColumnWidth =1695
                    ColumnOrder =62
                    TabIndex =16
                    Name ="GRPCODE"
                    ControlSource ="GRPCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     CODE, NAMES\015\012FROM         dbo.TCOD_STUFGROUP"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4764
                            Top =6576
                            Width =885
                            Height =240
                            Name ="Label39"
                            Caption ="نام گروه كالا"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1475
                    Top =5499
                    ColumnWidth =1530
                    ColumnOrder =36
                    TabIndex =57
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT     DEPATMAN, DEPNAME\015\012FROM         dbo.DEPART"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4007
                            Top =5499
                            Width =870
                            Height =240
                            Name ="Label84"
                            Caption ="واحد"
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
                    Top =27269
                    Width =1878
                    Height =318
                    ColumnWidth =1260
                    ColumnOrder =44
                    TabIndex =67
                    Name ="coln1"
                    ControlSource ="coln1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =27269
                            Width =1905
                            Height =240
                            Name ="LCOLn1"
                            Caption ="1"
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
                    Top =27721
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =45
                    TabIndex =68
                    Name ="coln2"
                    ControlSource ="coln2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =27721
                            Width =1905
                            Height =240
                            Name ="LCOLn2"
                            Caption ="2"
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
                    Top =28173
                    Width =1878
                    Height =318
                    ColumnWidth =1140
                    ColumnOrder =47
                    TabIndex =69
                    Name ="coln3"
                    ControlSource ="coln3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =28173
                            Width =1905
                            Height =240
                            Name ="LCOLn3"
                            Caption ="3"
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
                    Top =28625
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =48
                    TabIndex =70
                    Name ="coln4"
                    ControlSource ="coln4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =28625
                            Width =1905
                            Height =240
                            Name ="LCOLn4"
                            Caption ="4"
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
                    Top =29077
                    Width =1878
                    Height =318
                    ColumnWidth =540
                    ColumnOrder =49
                    TabIndex =71
                    Name ="coln5"
                    ControlSource ="coln5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =29077
                            Width =1905
                            Height =240
                            Name ="LCOLn5"
                            Caption ="5"
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
                    Top =29529
                    Width =1878
                    Height =318
                    ColumnWidth =810
                    ColumnOrder =50
                    TabIndex =72
                    Name ="coln6"
                    ControlSource ="coln6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =29529
                            Width =1905
                            Height =240
                            Name ="LCOLn6"
                            Caption ="6"
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
                    Top =29981
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =51
                    TabIndex =73
                    Name ="coln7"
                    ControlSource ="coln7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =29981
                            Width =1905
                            Height =240
                            Name ="LCOLn7"
                            Caption ="7"
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
                    Top =30433
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =52
                    TabIndex =74
                    Name ="coln8"
                    ControlSource ="coln8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =30433
                            Width =1905
                            Height =240
                            Name ="LCOLn8"
                            Caption ="8"
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
                    Top =30885
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =53
                    TabIndex =75
                    Name ="coln9"
                    ControlSource ="coln9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1947
                            Top =30885
                            Width =1905
                            Height =240
                            Name ="LCOLn9"
                            Caption ="9"
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
                    Left =795
                    Top =22050
                    Width =1878
                    Height =318
                    ColumnWidth =480
                    ColumnOrder =37
                    TabIndex =58
                    Name ="col1"
                    ControlSource ="col1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =22050
                            Width =1905
                            Height =240
                            Name ="LCOL1"
                            Caption ="1"
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
                    Left =795
                    Top =22502
                    Width =1878
                    Height =318
                    ColumnWidth =645
                    ColumnOrder =38
                    TabIndex =59
                    Name ="col2"
                    ControlSource ="col2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =22502
                            Width =1905
                            Height =240
                            Name ="LCOL2"
                            Caption ="2"
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
                    Left =795
                    Top =22954
                    Width =1878
                    Height =318
                    ColumnWidth =705
                    ColumnOrder =46
                    TabIndex =60
                    Name ="col3"
                    ControlSource ="col3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =22954
                            Width =1905
                            Height =240
                            Name ="LCOL3"
                            Caption ="3"
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
                    Left =795
                    Top =23406
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =39
                    TabIndex =61
                    Name ="col4"
                    ControlSource ="col4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =23406
                            Width =1905
                            Height =240
                            Name ="LCOL4"
                            Caption ="4"
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
                    Left =795
                    Top =23858
                    Width =1878
                    Height =318
                    ColumnWidth =1950
                    ColumnOrder =40
                    TabIndex =62
                    Name ="col5"
                    ControlSource ="col5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =23858
                            Width =1905
                            Height =240
                            Name ="LCOL5"
                            Caption ="5"
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
                    Left =795
                    Top =24310
                    Width =1878
                    Height =318
                    ColumnWidth =510
                    ColumnOrder =10
                    TabIndex =63
                    Name ="col6"
                    ControlSource ="col6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =24310
                            Width =1905
                            Height =240
                            Name ="LCOL6"
                            Caption ="6"
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
                    Left =795
                    Top =24762
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =41
                    TabIndex =64
                    Name ="col7"
                    ControlSource ="col7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =24762
                            Width =1905
                            Height =240
                            Name ="LCOL7"
                            Caption ="7"
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
                    Left =795
                    Top =25214
                    Width =1878
                    Height =318
                    ColumnWidth =1417
                    ColumnOrder =42
                    TabIndex =65
                    Name ="col8"
                    ControlSource ="col8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2742
                            Top =25214
                            Width =1905
                            Height =240
                            Name ="LCOL8"
                            Caption ="8"
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
                    Left =795
                    Top =25666
                    Width =1878
                    Height =318
                    ColumnWidth =645
                    ColumnOrder =43
                    TabIndex =66
                    Name ="col9"
                    ControlSource ="col9"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2742
                            Top =25666
                            Width =1905
                            Height =240
                            Name ="LCOL9"
                            Caption ="9"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2212
                    Top =21713
                    ColumnWidth =1417
                    ColumnOrder =72
                    TabIndex =76
                    Name ="TKHN"
                    ControlSource ="TKHN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5085
                            Top =21720
                            Width =1350
                            Height =240
                            Name ="Label123"
                            Caption ="درصد تخفيف نقدي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =28346
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =79
                    TabIndex =77
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6720
                            Top =28346
                            Width =825
                            Height =240
                            Name ="Label124"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =28686
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =80
                    TabIndex =78
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7155
                            Top =28680
                            Width =435
                            Height =240
                            Name ="Label125"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =29026
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =81
                    TabIndex =79
                    Name ="CODE_E"
                    ControlSource ="CODE_E"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6810
                            Top =29025
                            Width =795
                            Height =405
                            Name ="Label126"
                            Caption ="کد اقتصادي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =29366
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =82
                    TabIndex =80
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6750
                            Top =29366
                            Width =795
                            Height =240
                            Name ="Label127"
                            Caption ="کد ملي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =29707
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =83
                    TabIndex =81
                    Name ="MOBILE"
                    ControlSource ="MOBILE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6855
                            Top =29707
                            Width =690
                            Height =240
                            Name ="Label128"
                            Caption ="موبايل"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =30047
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =84
                    TabIndex =82
                    Name ="Longitude"
                    ControlSource ="Longitude"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6705
                            Top =30045
                            Width =915
                            Height =405
                            Name ="Label129"
                            Caption ="طول جغرافيايي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =255
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4479
                    Top =30387
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =85
                    TabIndex =83
                    Name ="Latitude"
                    ControlSource ="Latitude"
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =6825
                            Top =30390
                            Width =810
                            Height =405
                            Name ="Label130"
                            Caption ="عرض جغرافيايي"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =95
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4195
                    Top =27779
                    Height =570
                    ColumnWidth =810
                    ColumnOrder =87
                    TabIndex =85
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6451
                            Top =27779
                            Width =810
                            Height =240
                            Name ="Label132"
                            Caption ="کد استان"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4648
                    Top =25738
                    Height =570
                    ColumnWidth =720
                    ColumnOrder =88
                    TabIndex =86
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6824
                            Top =25738
                            Width =810
                            Height =240
                            Name ="Label133"
                            Caption ="کد شهر"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4251
                    Top =27212
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =77
                    TabIndex =87
                    Name ="OSNAME"
                    ControlSource ="OSNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6392
                            Top =27212
                            Width =765
                            Height =240
                            Name ="Label134"
                            Caption ="استان"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4081
                    Top =26532
                    Height =570
                    ColumnWidth =1417
                    ColumnOrder =78
                    TabIndex =88
                    Name ="CITYNAME"
                    ControlSource ="CITYNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6007
                            Top =26532
                            Width =900
                            Height =240
                            Name ="Label135"
                            Caption =" شهر"
                        End
                    End
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4479
                    Top =30727
                    ColumnWidth =4035
                    ColumnOrder =86
                    TabIndex =84
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.ROUTE_NAME + N' - ' + CUST_HESAB.NAME"
                        " + N' - ' + CUST_HESAB.hes AS Expr1 FROM Visit_route INNER JOIN CUST_HESAB ON Vi"
                        "sit_route.HES = CUST_HESAB.hes WHERE (Visit_route.RACTIVE = 1)"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =6390
                            Top =30727
                            Width =1155
                            Height =240
                            Name ="Label131"
                            Caption ="مسير ويزيت"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KALAS_MAIN_ADVANCE.cls"
