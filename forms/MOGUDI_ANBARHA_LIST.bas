Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5730
    RowHeight =285
    ItemSuffix =22
    Left =825
    Top =1110
    Right =8850
    Bottom =5745
    DatasheetGridlinesColor =12632256
    MaxRecords =100000000
    RecSrcDt = Begin
        0xe64e747ea8dee340
    End
    RecordSource ="MOGUDI_KOL_ANBARHA"
    Caption ="موجودي كلي كالاها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>كد كالا</x:Name>\015\012   <x:SourceName>CODE</x:SourceName>\015\012 "
        "  <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy"
        ">UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Piv"
        "otField>\015\012   <x:Name>موجودي</x:Name>\015\012   <x:SourceName>MANDAH</x:Sou"
        "rceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientatio"
        "n/>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrdere"
        "dMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>MEGF</x:Name>\015\012   <x:DataType>Number</x"
        ":DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Compar"
        "eOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayI"
        "nFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>"
        "MABLK</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedTyp"
        "e>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Co"
        "mpareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotFie"
        "ld>\015\012  <x:PivotField>\015\012   <x:Name>SumOfMEG</x:Name>\015\012   <x:Dat"
        "aType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <"
        "x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>نام كالا</x:Name>\015\012   <x:SourceName>NAME</x:SourceName>\015\012"
        "   <x:Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adVarWChar</x:Enc"
        "odedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembe"
        "rsBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>نام واحد</x:Name>\015\012   <x:SourceName>NAMES</x:SourceName>\015"
        "\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMemb"
        "ersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <"
        "x:PivotField>\015\012   <x:Name>KK</x:Name>\015\012   <x:EncodedType>adVarChar</"
        "x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdered"
        "MembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012 "
        " <x:PivotField>\015\012   <x:Name>VCOD</x:Name>\015\012   <x:DataType>Integer</x"
        ":DataType>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:Compa"
        "reOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplay"
        "InFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name"
        ">KINDK</x:Name>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedT"
        "ype>adSmallInt</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</"
        "x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:Pivo"
        "tField>\015\012  <x:PivotField>\015\012   <x:Name>GHEMAT</x:Name>\015\012   <x:D"
        "ataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>FI</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <"
        "x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Uniq"
        "ueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  "
        "</x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>N_FANI</x:Name>\015\012"
        "   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersB"
        "y>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015"
        "\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Positio"
        "n>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>Sum of موجودي</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012"
        "   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation"
        ">\015\012   <x:Position>1</x:Position>\015\012   <x:ParentField>موجودي</x:Parent"
        "Field>\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0</x:T"
        "op>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012 "
        "  <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>4</x:SeqNum>\015\012  </x:P"
        "ivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label"
        ">\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    "
        "<x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits"
        ">\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletio"
        "ns>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012"
        "</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.6854         </x:OWCVersion>\015\012  <x:Width>26988<"
        "/x:Width>\015\012  <x:Height>15161</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012"
        "  <x:FormatValue>\015\012   <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   "
        "<x:Data>2</x:Data>\015\012  </x:FormatValue>\015\012  <x:PivotAggOrientation>Ser"
        "ies</x:PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping"
        "/>\015\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:"
        "Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>"
        "\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012  "
        " <x:Entry>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:En"
        "try>#600080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#0"
        "08080</x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080<"
        "/x:Entry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Ent"
        "ry>\015\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015"
        "\012   <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012  "
        " <x:Entry>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial<"
        "/x:DefaultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015"
        "\012     <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015"
        "\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012"
        "      <x:Name>Series</x:Name>\015\012      <x:Caption>\015\012       <x:DataSour"
        "ceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Series&quot;</x:Data>"
        "\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012      <x:Cat"
        "egory>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:C"
        "ategory>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x:Dat"
        "aSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012    "
        "  </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbo"
        "l>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <"
        "x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012      <x:Identifier>!.Sum of موجودي</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Dimension>\015\012      <x:ScaleID>324469416</x:ScaleID>\015"
        "\012      <x:Index>Categories</x:Index>\015\012     </x:Dimension>\015\012     <"
        "x:Dimension>\015\012      <x:ScaleID>324470024</x:ScaleID>\015\012      <x:Index"
        ">Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012   "
        "   <x:ScaleID>324470228</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015"
        "\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012     <x:GapW"
        "idth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012"
        "    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>324473416</x:AxisID>\015"
        "\012     <x:ScaleID>324469416</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015"
        "\012     <x:Font>\015\012      <x:FontName>Tahoma</x:FontName>\015\012      <x:B"
        ">Automatic</x:B>\015\012      <x:I>Automatic</x:I>\015\012      <x:U>Automatic</"
        "x:U>\015\012     </x:Font>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012"
        "     <x:MinorTick>None</x:MinorTick>\015\012     <x:Title>\015\012      <x:Font>"
        "\015\012       <x:FontName>Tahoma</x:FontName>\015\012       <x:Size>8</x:Size>\015"
        "\012      </x:Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:"
        "ColorIndex>\015\012      </x:Border>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;نام كالا&quot;"
        "</x:Data>\015\012      </x:Caption>\015\012      <x:Position>Bottom</x:Position>"
        "\015\012     </x:Title>\015\012     <x:Placement>Bottom</x:Placement>\015\012   "
        "  <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>"
        "\015\012     <x:AxisID>324474812</x:AxisID>\015\012     <x:ScaleID>324470024</x:"
        "ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     <x:Number>\015\012      "
        "<x:FormatString>General</x:FormatString>\015\012     </x:Number>\015\012     <x:"
        "MajorGridlines>\015\012     </x:MajorGridlines>\015\012     <x:MajorTick>Outside"
        "</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012     <x:Title>"
        "\015\012      <x:Font>\015\012       <x:FontName>Tahoma</x:FontName>\015\012    "
        "   <x:Size>8</x:Size>\015\012      </x:Font>\015\012      <x:Border>\015\012    "
        "   <x:ColorIndex>None</x:ColorIndex>\015\012      </x:Border>\015\012      <x:Ca"
        "ption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:"
        "Data>&quot;موجودي&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Posit"
        "ion>Left</x:Position>\015\012     </x:Title>\015\012     <x:Placement>Left</x:Pl"
        "acement>\015\012    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:Identifier></"
        "x:Identifier>\015\012  </x:Chart>\015\012  <x:Title>\015\012   <x:Font>\015\012 "
        "   <x:FontName>Titr Mazar</x:FontName>\015\012    <x:B>Automatic</x:B>\015\012  "
        "  <x:I>Automatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012"
        "   <x:Border>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>"
        "\015\012   <x:Caption>\015\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012"
        "    <x:Data>&quot;موجودي كل انبار ها&quot;</x:Data>\015\012   </x:Caption>\015\012"
        "  </x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>324469416</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>324470024</x:ScaleI"
        "D>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>324470228</x:"
        "ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  <x:HideSc"
        "reenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
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
            Height =3194
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1104
                    Top =113
                    Width =2490
                    ColumnWidth =1080
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="CODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4965
                            Top =120
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =3
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1893
                    Top =453
                    ColumnWidth =1890
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4950
                            Top =450
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="موجودي"
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
                    Left =1893
                    Top =1474
                    ColumnWidth =1785
                    ColumnOrder =4
                    TabIndex =2
                    Name ="FI"
                    ControlSource ="FI"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4115
                            Top =1474
                            Width =1350
                            Height =240
                            Name ="Label9"
                            Caption ="مبلغ ميانگين"
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
                    Left =1893
                    Top =1814
                    ColumnWidth =1500
                    ColumnOrder =5
                    TabIndex =3
                    Name ="MABLK"
                    ControlSource ="MABLK"
                    Format ="Standard"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4185
                            Top =1815
                            Width =1545
                            Height =225
                            Name ="Label11"
                            Caption ="مبلغ موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =144
                    Top =2154
                    Width =3450
                    ColumnWidth =6135
                    ColumnOrder =1
                    TabIndex =4
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4965
                            Top =2160
                            Width =540
                            Height =240
                            Name ="Label13"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =804
                    Top =2494
                    Width =2790
                    ColumnWidth =1185
                    ColumnOrder =3
                    TabIndex =5
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4835
                            Top =2494
                            Width =630
                            Height =240
                            Name ="Label15"
                            Caption ="نام واحد"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1110
                    Top =900
                    Width =2490
                    ColumnWidth =615
                    TabIndex =6
                    Name ="Text20"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4971
                            Top =907
                            Width =510
                            Height =240
                            Name ="Label21"
                            Caption ="كد كالا"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "MOGUDI_ANBARHA_LIST.cls"
