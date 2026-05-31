Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    DefaultView =4
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    TabularFamily =0
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =4940
    DatasheetFontHeight =10
    ItemSuffix =12
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xbd7cd8f81b6ee540
    End
    RecordSource ="TEST5"
    Caption ="نمودار سود فروش"
    DatasheetFontName ="Arial"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>ماه</x:Name>\015\012   <x:SourceName>MM</x:SourceName>\015\012   <x:D"
        "ataType>Integer</x:DataType>\015\012   <x:Orientation>Row</x:Orientation>\015\012"
        "   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy"
        ">UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012  </x:Piv"
        "otField>\015\012  <x:PivotField>\015\012   <x:Name>سود ماه</x:Name>\015\012   <x"
        ":SourceName>sudday</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:PLDataOrientation/>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </"
        "x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>فروش</x:Name>\015\012   "
        "<x:SourceName>FROOSH</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:PLDataOrientation/>\015\012   <x:PLPosition>2</x:PLPosition>\015\012  "
        " <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Un"
        "iqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>هزينه</x:Name>\015\012   <x:SourceName>HAZINEH</x:Source"
        "Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/>"
        "\015\012   <x:PLPosition>4</x:PLPosition>\015\012   <x:EncodedType>adDouble</x:E"
        "ncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMem"
        "bersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ماه1<"
        "/x:Name>\015\012   <x:SourceName>MON</x:SourceName>\015\012   <x:PLCaption>ماه</"
        "x:PLCaption>\015\012   <x:FilterCaption>ماه</x:FilterCaption>\015\012   <x:Orien"
        "tation>Row</x:Orientation>\015\012   <x:Position>2</x:Position>\015\012   <x:Enc"
        "odedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueN"
        "ame</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>"
        "\015\012  <x:PivotField>\015\012   <x:Name>قيمت تمام شده</x:Name>\015\012   <x:S"
        "ourceName>GHEYMAT</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:PLDataOrientation/>\015\012   <x:PLPosition>3</x:PLPosition>\015\012   <x:"
        "EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Unique"
        "Name</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation"
        ">\015\012   <x:Position>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:Pi"
        "votField>\015\012  <x:PivotField>\015\012   <x:Name>Sum of سود ماه</x:Name>\015\012"
        "   <x:PLName>Total1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012"
        "   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012"
        "   <x:ParentField>سود ماه</x:ParentField>\015\012  </x:PivotField>\015\012  <x:P"
        "ivotField>\015\012   <x:Name>Sum of فروش</x:Name>\015\012   <x:PLName>Total2</x:"
        "PLName>\015\012   <x:TotalNumber>1</x:TotalNumber>\015\012   <x:Orientation>Data"
        "</x:Orientation>\015\012   <x:Position>2</x:Position>\015\012   <x:ParentField>ف"
        "روش</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <"
        "x:Name>Sum of قيمت تمام شده</x:Name>\015\012   <x:PLName>Total4</x:PLName>\015\012"
        "   <x:TotalNumber>2</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation"
        ">\015\012   <x:Position>3</x:Position>\015\012   <x:ParentField>قيمت تمام شده</x"
        ":ParentField>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name"
        ">Sum of هزينه</x:Name>\015\012   <x:PLName>Total3</x:PLName>\015\012   <x:TotalN"
        "umber>3</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   "
        "<x:Position>4</x:Position>\015\012   <x:ParentField>هزينه</x:ParentField>\015\012"
        "  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0.0.0</x:Top>\015\012 "
        "  <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0</x:Left>\015\012   <x:LeftOff"
        "set>0</x:LeftOffset>\015\012   <x:SeqNum>106</x:SeqNum>\015\012  </x:PivotData>\015"
        "\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    "
        "<x:Caption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible"
        "/>\015\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <"
        "x:AllowAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:A"
        "llowDeletions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>36036<"
        "/x:Width>\015\012  <x:Height>16060</x:Height>\015\012  <x:DataSource>\015\012   "
        "<x:Type>PivotList</x:Type>\015\012  </x:DataSource>\015\012  <x:BoundSeries>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundSeries>\015\012"
        "  <x:Category>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:C"
        "ategory>\015\012  <x:Value>\015\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015"
        "\012  </x:Value>\015\012  <x:BoundCharts>\015\012   <x:DataSourceIndex>0</x:Data"
        "SourceIndex>\015\012  </x:BoundCharts>\015\012  <x:PivotAggOrientation>Series</x"
        ":PivotAggOrientation>\015\012  <x:HidePivotFields/>\015\012  <x:NoGrouping/>\015"
        "\012  <x:NoFiltering/>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:Entry>"
        "\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012"
        "   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:"
        "Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>"
        "#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#00000"
        "0</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#8080FF</x:Entry>\015\012   <x:Entry>#802060</x:Entry>\015\012   <x:En"
        "try>#FFFFA0</x:Entry>\015\012   <x:Entry>#A0E0E0</x:Entry>\015\012   <x:Entry>#6"
        "00080</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012   <x:Entry>#008080<"
        "/x:Entry>\015\012   <x:Entry>#C0C0FF</x:Entry>\015\012   <x:Entry>#000080</x:Ent"
        "ry>\015\012   <x:Entry>#FF00FF</x:Entry>\015\012   <x:Entry>#80FFFF</x:Entry>\015"
        "\012   <x:Entry>#0080FF</x:Entry>\015\012   <x:Entry>#FF8080</x:Entry>\015\012  "
        " <x:Entry>#C0FF80</x:Entry>\015\012   <x:Entry>#FFC0FF</x:Entry>\015\012   <x:En"
        "try>#FF80FF</x:Entry>\015\012  </x:Palette>\015\012  <x:DefaultFont>Arial</x:Def"
        "aultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x:Graph>\015\012"
        "     <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column</x:Type>\015\012"
        "     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012  "
        "    <x:Name>Sum of سود ماه</x:Name>\015\012      <x:Caption>\015\012       <x:Da"
        "taSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of سود ماه&"
        "quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012"
        "      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x"
        ":DataSourceIndex>\015\012      </x:Value>\015\012      <x:DataLabels>\015\012   "
        "    <x:Border>\015\012        <x:ColorIndex>None</x:ColorIndex>\015\012       </"
        "x:Border>\015\012       <x:Number>\015\012        <x:FormatString>General</x:For"
        "matString>\015\012       </x:Number>\015\012       <x:ShowValue/>\015\012       "
        "<x:Separator>;</x:Separator>\015\012       <x:Position>OutsideEnd</x:Position>\015"
        "\012      </x:DataLabels>\015\012      <x:Marker>\015\012       <x:Symbol>None</"
        "x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012"
        "      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      <x:Identifier>!.Sum of سود ماه</x:Identifier>\015\012    "
        " </x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:"
        "FormatMap>\015\012      <x:Name>Sum of فروش</x:Name>\015\012      <x:Caption>\015"
        "\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot"
        ";Sum of فروش&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>2</x"
        ":Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSour"
        "ceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:Data"
        "SourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:DataLa"
        "bels>\015\012       <x:Border>\015\012        <x:ColorIndex>None</x:ColorIndex>\015"
        "\012       </x:Border>\015\012       <x:Number>\015\012        <x:FormatString>G"
        "eneral</x:FormatString>\015\012       </x:Number>\015\012       <x:ShowValue/>\015"
        "\012       <x:Separator>;</x:Separator>\015\012       <x:Position>OutsideEnd</x:"
        "Position>\015\012      </x:DataLabels>\015\012      <x:Marker>\015\012       <x:"
        "Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Exp"
        "lode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex"
        ">0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of فروش</x:Identifier>\015"
        "\012     </x:Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012  "
        "    </x:FormatMap>\015\012      <x:Name>Sum of قيمت تمام شده</x:Name>\015\012   "
        "   <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012  "
        "     <x:Data>&quot;Sum of قيمت تمام شده&quot;</x:Data>\015\012      </x:Caption>"
        "\015\012      <x:Index>3</x:Index>\015\012      <x:Category>\015\012       <x:Da"
        "taSourceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:V"
        "alue>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Va"
        "lue>\015\012      <x:DataLabels>\015\012       <x:Border>\015\012        <x:Colo"
        "rIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012       <x:Number>\015"
        "\012        <x:FormatString>General</x:FormatString>\015\012       </x:Number>\015"
        "\012       <x:ShowValue/>\015\012       <x:Separator>;</x:Separator>\015\012    "
        "   <x:Position>OutsideEnd</x:Position>\015\012      </x:DataLabels>\015\012     "
        " <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015"
        "\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015"
        "\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!."
        "Sum of قيمت تمام شده</x:Identifier>\015\012     </x:Series>\015\012     <x:Serie"
        "s>\015\012      <x:FormatMap>\015\012      </x:FormatMap>\015\012      <x:Name>S"
        "um of هزينه</x:Name>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-"
        "1</x:DataSourceIndex>\015\012       <x:Data>&quot;Sum of هزينه&quot;</x:Data>\015"
        "\012      </x:Caption>\015\012      <x:Index>1</x:Index>\015\012      <x:Categor"
        "y>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Categ"
        "ory>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012      </x:Value>\015\012      <x:DataLabels>\015\012       <x:Border>\015"
        "\012        <x:ColorIndex>None</x:ColorIndex>\015\012       </x:Border>\015\012 "
        "      <x:Number>\015\012        <x:FormatString>General</x:FormatString>\015\012"
        "       </x:Number>\015\012       <x:ShowValue/>\015\012       <x:Separator>;</x:"
        "Separator>\015\012       <x:Position>OutsideEnd</x:Position>\015\012      </x:Da"
        "taLabels>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012"
        "      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012      <x:Thicknes"
        "s>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012"
        "      <x:Identifier>!.Sum of هزينه</x:Identifier>\015\012     </x:Series>\015\012"
        "     <x:Dimension>\015\012      <x:ScaleID>272712376</x:ScaleID>\015\012      <x"
        ":Index>Categories</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>"
        "\015\012      <x:ScaleID>272712596</x:ScaleID>\015\012      <x:Index>Value</x:In"
        "dex>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleI"
        "D>272712816</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015\012     "
        "</x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012     <x:GapWidth>150<"
        "/x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012    </x"
        ":Graph>\015\012    <x:Axis>\015\012     <x:AxisID>272722580</x:AxisID>\015\012  "
        "   <x:ScaleID>272712376</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015\012"
        "     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTi"
        "ck>\015\012     <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Au"
        "to</x:GroupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:Axis"
        "ID>272723052</x:AxisID>\015\012     <x:ScaleID>272712596</x:ScaleID>\015\012    "
        " <x:Type>Value</x:Type>\015\012     <x:Number>\015\012      <x:FormatString>#,##"
        "#</x:FormatString>\015\012     </x:Number>\015\012     <x:MajorGridlines>\015\012"
        "     </x:MajorGridlines>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012 "
        "    <x:MinorTick>None</x:MinorTick>\015\012     <x:Placement>Left</x:Placement>\015"
        "\012    </x:Axis>\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier"
        ">\015\012  </x:Chart>\015\012  <x:Scaling>\015\012   <x:ScaleID>272712376</x:Sca"
        "leID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>272712596<"
        "/x:ScaleID>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>2727"
        "12816</x:ScaleID>\015\012  </x:Scaling>\015\012  <x:HidePassiveAlert/>\015\012  "
        "<x:HideScreenTips/>\015\012 </x:ChartSpace>\015\012</xml>"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
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
            Height =2167
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =113
                    Name ="MM"
                    ControlSource ="MM"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =113
                            Width =315
                            Height =240
                            Name ="Label1"
                            Caption ="ماه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =453
                    ColumnWidth =2160
                    TabIndex =1
                    Name ="sudday"
                    ControlSource ="sudday"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =450
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="سود ماه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =793
                    TabIndex =2
                    Name ="FROOSH"
                    ControlSource ="FROOSH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =793
                            Width =705
                            Height =240
                            Name ="Label5"
                            Caption ="فروش"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1133
                    TabIndex =3
                    Name ="HAZINEH"
                    ControlSource ="HAZINEH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1133
                            Width =735
                            Height =240
                            Name ="Label7"
                            Caption ="هزينه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1474
                    Width =2490
                    TabIndex =4
                    Name ="MON"
                    ControlSource ="MON"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1474
                            Width =420
                            Height =240
                            Name ="Label9"
                            Caption ="ماه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1927
                    Top =1927
                    TabIndex =5
                    Name ="GHEYMAT"
                    ControlSource ="GHEYMAT"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =1927
                            Width =1185
                            Height =240
                            Name ="Label11"
                            Caption ="قيمت تمام شده"
                        End
                    End
                End
            End
        End
    End
End
