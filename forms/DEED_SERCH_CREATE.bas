Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =7237
    RowHeight =375
    ItemSuffix =44
    Left =1260
    Top =900
    Right =12270
    Bottom =5745
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    OnApplyFilter ="[Event Procedure]"
    RecSrcDt = Begin
        0xb5bd8e8eaeb8e540
    End
    ServerFilter ="( ((HES_K = 128) AND (HES_M = 1)) ) or (((HES_K = 128) AND (HES_M = 1)) )"
    RecordSource ="SEARCH_ON SANAD"
    Caption ="جستجو در شرح سند"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:Dis"
        "playFieldList/>\015\012  <x:FieldListTop>66</x:FieldListTop>\015\012  <x:FieldLi"
        "stLeft>223</x:FieldListLeft>\015\012  <x:FieldListBottom>435</x:FieldListBottom>"
        "\015\012  <x:FieldListRight>423</x:FieldListRight>\015\012  <x:CacheDetails/>\015"
        "\012  <x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:DataAxisEmpt"
        "y/>\015\012  <x:PivotField>\015\012   <x:Name> سند</x:Name>\015\012   <x:SourceN"
        "ame>N_S</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:En"
        "codedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueNa"
        "me</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>RADIF</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name> كل</x:Name>\015\012"
        "   <x:SourceName>HES_K</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>"
        "\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:EncodedType>adInte"
        "ger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOr"
        "deredMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:Piv"
        "otField>\015\012   <x:Name> معين</x:Name>\015\012   <x:SourceName>HES_M</x:Sourc"
        "eName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:EncodedType>adInt"
        "eger</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareO"
        "rderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:N"
        "ame>شرح</x:Name>\015\012   <x:SourceName>SHARH</x:SourceName>\015\012   <x:Encod"
        "edType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueNam"
        "e</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>بدهكار</x:Name>\015\012   <x:SourceName>BED</x:SourceName>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:NumberFormat>#,##0.00;-#,##0.00</x:NumberFormat>\015\012   <x:"
        "CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Pivo"
        "tField>\015\012  <x:PivotField>\015\012   <x:Name>بستانكار</x:Name>\015\012   <x"
        ":SourceName>BES</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>#,##0.00;-#"
        ",##0.00</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Comp"
        "areOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   "
        "<x:Name>شماره سري</x:Name>\015\012   <x:SourceName>N_SERI</x:SourceName>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>بانك</x:Name"
        ">\015\012   <x:SourceName>BANK</x:SourceName>\015\012   <x:DataType>Integer</x:D"
        "ataType>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:Compare"
        "OrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>"
        "\015\012  <x:PivotField>\015\012   <x:Name>شماره فاكتور</x:Name>\015\012   <x:So"
        "urceName>NUMBER</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>برچسب</x:Name>\015\012   <x:SourceName>TAG</x:SourceNa"
        "me>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble<"
        "/x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdere"
        "dMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>ت"
        "اريخ سند</x:Name>\015\012   <x:SourceName>DATE_S</x:SourceName>\015\012   <x:Dat"
        "aType>Number</x:DataType>\015\012   <x:EncodedType>adBigInt</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </"
        "x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name> تفصيلي</x:Name>\015\012"
        "   <x:SourceName>TNUMBER</x:SourceName>\015\012   <x:DataType>Integer</x:DataTyp"
        "e>\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrdere"
        "dMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>شرح حساب</x:Name>\015\012   <x:SourceName>NAM"
        "E</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <"
        "x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:Pi"
        "votField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:Name>\015\012   <x:O"
        "rientation>Column</x:Orientation>\015\012   <x:Position>-1</x:Position>\015\012 "
        "  <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:N"
        "ame>Sum of  كل</x:Name>\015\012   <x:PLName>Total1</x:PLName>\015\012   <x:Total"
        "Number>0</x:TotalNumber>\015\012   <x:Orientation>Data</x:Orientation>\015\012  "
        " <x:Position>1</x:Position>\015\012   <x:ParentField> كل</x:ParentField>\015\012"
        "  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:Top>0</x:Top>\015\012   <x"
        ":TopOffset>0</x:TopOffset>\015\012   <x:Left>0.0</x:Left>\015\012   <x:LeftOffse"
        "t>0</x:LeftOffset>\015\012   <x:SeqNum>3</x:SeqNum>\015\012  </x:PivotData>\015\012"
        "  <x:PivotView>\015\012   <x:IsNotFiltered/>\015\012   <x:Label>\015\012    <x:C"
        "aption>Microsoft Office PivotTable 10.0</x:Caption>\015\012    <x:NotVisible/>\015"
        "\012   </x:Label>\015\012   <x:AllowEdits>false</x:AllowEdits>\015\012   <x:Allo"
        "wAdditions>false</x:AllowAdditions>\015\012   <x:AllowDeletions>false</x:AllowDe"
        "letions>\015\012  </x:PivotView>\015\012 </x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>21061<"
        "/x:Width>\015\012  <x:Height>10742</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012"
        "  <x:FormatValue>\015\012   <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   "
        "<x:Data>2</x:Data>\015\012  </x:FormatValue>\015\012  <x:PivotAggOrientation>Ser"
        "ies</x:PivotAggOrientation>\015\012  <x:DisplayFieldList/>\015\012  <x:DefaultFo"
        "nt>Arial</x:DefaultFont>\015\012  <x:Chart>\015\012   <x:PlotArea>\015\012    <x"
        ":Graph>\015\012     <x:SubType>Clustered</x:SubType>\015\012     <x:Type>Column<"
        "/x:Type>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:Form"
        "atMap>\015\012      <x:Name>115</x:Name>\015\012      <x:Caption>\015\012       "
        "<x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;115&quot;<"
        "/x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012    "
        "  <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012   "
        "   </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:Da"
        "taSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012     "
        "  <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015"
        "\012      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None<"
        "/x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012"
        "      <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      <x:Identifier>!._80000073</x:Identifier>\015\012     </x:"
        "Series>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:Forma"
        "tMap>\015\012      <x:Name>121</x:Name>\015\012      <x:Caption>\015\012       <"
        "x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;121&quot;</"
        "x:Data>\015\012      </x:Caption>\015\012      <x:Index>1</x:Index>\015\012     "
        " <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012    "
        "  </x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:Dat"
        "aSourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012      "
        " <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._80000079</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>122</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;122&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>2</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000007A</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>123</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;123&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>3</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000007B</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>124</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;124&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>4</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000007C</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>125</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;125&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>5</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000007D</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>126</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;126&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>6</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000007E</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>127</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;127&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>7</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000007F</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>129</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;129&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>8</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._80000081</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>211</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;211&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>9</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800000D3</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>212</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;212&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>10</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800000D4</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>214</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;214&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>11</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800000D6</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>300</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;300&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>12</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000012C</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>313</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;313&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>13</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._80000139</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>314</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;314&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>14</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000013A</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>413</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;413&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>15</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000019D</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>414</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;414&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>16</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000019E</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>511</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;511&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>17</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800001FF</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>512</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;512&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>18</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002</x:Identifier>\015\012     </x:Serie"
        "s>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap>"
        "\015\012      <x:Name>513</x:Name>\015\012      <x:Caption>\015\012       <x:Dat"
        "aSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;513&quot;</x:Dat"
        "a>\015\012      </x:Caption>\015\012      <x:Index>19</x:Index>\015\012      <x:"
        "Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </"
        "x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x:"
        "DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012 "
        "     </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Sy"
        "mbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012    "
        "  <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceI"
        "ndex>\015\012      <x:Identifier>!._80000201</x:Identifier>\015\012     </x:Seri"
        "es>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMap"
        ">\015\012      <x:Name>514</x:Name>\015\012      <x:Caption>\015\012       <x:Da"
        "taSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;514&quot;</x:Da"
        "ta>\015\012      </x:Caption>\015\012      <x:Index>20</x:Index>\015\012      <x"
        ":Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <"
        "/x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSo"
        "urceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x"
        ":DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._80000202</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>611</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;611&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>21</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._80000263</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>711</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;711&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>22</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002C7</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>712</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;712&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>23</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002C8</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>713</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;713&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>24</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002C9</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>714</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;714&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>25</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002CA</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>717</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;717&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>26</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002CD</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>718</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;718&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>27</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002CE</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>719</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;719&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>28</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002CF</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>721</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;721&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>29</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002D1</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>722</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;722&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>30</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002D2</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>723</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;723&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>31</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002D3</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>724</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;724&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>32</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002D4</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>725</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;725&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>33</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002D5</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>726</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;726&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>34</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._800002D6</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012      </x:FormatMa"
        "p>\015\012      <x:Name>811</x:Name>\015\012      <x:Caption>\015\012       <x:D"
        "ataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;811&quot;</x:D"
        "ata>\015\012      </x:Caption>\015\012      <x:Index>35</x:Index>\015\012      <"
        "x:Category>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      "
        "</x:Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataS"
        "ourceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <"
        "x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012"
        "      </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:S"
        "ymbol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012   "
        "   <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSource"
        "Index>\015\012      <x:Identifier>!._8000032B</x:Identifier>\015\012     </x:Ser"
        "ies>\015\012     <x:Dimension>\015\012      <x:ScaleID>200323456</x:ScaleID>\015"
        "\012      <x:Index>Categories</x:Index>\015\012     </x:Dimension>\015\012     <"
        "x:Dimension>\015\012      <x:ScaleID>200323660</x:ScaleID>\015\012      <x:Index"
        ">Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012   "
        "   <x:ScaleID>200323864</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015"
        "\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012     <x:GapW"
        "idth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012"
        "    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>31580096</x:AxisID>\015"
        "\012     <x:ScaleID>200323456</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015"
        "\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:Min"
        "orTick>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Ar"
        "ial</x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Font>\015\012"
        "      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </"
        "x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSo"
        "urceIndex>\015\012       <x:Data>&quot;Axis Title&quot;</x:Data>\015\012      </"
        "x:Caption>\015\012      <x:Position>Bottom</x:Position>\015\012     </x:Title>\015"
        "\012     <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:G"
        "roupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>3158"
        "0568</x:AxisID>\015\012     <x:ScaleID>200323660</x:ScaleID>\015\012     <x:Type"
        ">Value</x:Type>\015\012     <x:Number>\015\012      <x:FormatString>General</x:F"
        "ormatString>\015\012     </x:Number>\015\012     <x:MajorGridlines>\015\012     "
        "</x:MajorGridlines>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <"
        "x:MinorTick>None</x:MinorTick>\015\012     <x:Title>\015\012      <x:Font>\015\012"
        "       <x:FontName>Arial</x:FontName>\015\012       <x:Size>10</x:Size>\015\012 "
        "     </x:Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:Color"
        "Index>\015\012      </x:Border>\015\012      <x:Caption>\015\012       <x:DataSo"
        "urceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Axis Title&quot;</x"
        ":Data>\015\012      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012"
        "     </x:Title>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>"
        "\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Ch"
        "art>\015\012  <x:Scaling>\015\012   <x:ScaleID>200323456</x:ScaleID>\015\012  </"
        "x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>200323660</x:ScaleID>\015\012"
        "  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>200323864</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:HasSelectionFeedback>True</x:HasSelectionFeedback"
        ">\015\012  <x:HidePassiveAlert/>\015\012 </x:ChartSpace>\015\012</xml>"
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
            Height =7269
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3405
                    Top =480
                    ColumnWidth =990
                    ColumnOrder =0
                    Name ="N_S"
                    ControlSource ="N_S"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6092
                            Top =480
                            Width =885
                            Height =240
                            Name ="Label1"
                            Caption =" سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =150
                    Top =1830
                    Width =4965
                    ColumnWidth =6480
                    ColumnOrder =7
                    TabIndex =3
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6566
                            Top =1830
                            Width =420
                            Height =240
                            Name ="Label9"
                            Caption ="شرح"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =2170
                    ColumnWidth =1560
                    ColumnOrder =8
                    TabIndex =4
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6461
                            Top =2170
                            Width =525
                            Height =240
                            Name ="Label11"
                            Caption ="بدهكار"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =2510
                    ColumnWidth =1545
                    ColumnOrder =9
                    TabIndex =5
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,##0.00;-#,##0.00"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6341
                            Top =2510
                            Width =645
                            Height =240
                            Name ="Label13"
                            Caption ="بستانكار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =2850
                    ColumnWidth =990
                    ColumnOrder =11
                    TabIndex =6
                    Name ="N_SERI"
                    ControlSource ="N_SERI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6026
                            Top =2850
                            Width =960
                            Height =240
                            Name ="Label15"
                            Caption ="شماره سري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =3190
                    ColumnWidth =465
                    ColumnOrder =12
                    TabIndex =7
                    Name ="BANK"
                    ControlSource ="BANK"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6611
                            Top =3190
                            Width =375
                            Height =240
                            Name ="Label17"
                            Caption ="بانك"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =3530
                    ColumnWidth =1125
                    ColumnOrder =13
                    TabIndex =8
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6026
                            Top =3530
                            Width =960
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3414
                    Top =3870
                    ColumnWidth =645
                    ColumnOrder =14
                    TabIndex =9
                    Name ="TAG"
                    ControlSource ="TAG"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6356
                            Top =3870
                            Width =630
                            Height =240
                            Name ="Label21"
                            Caption ="برچسب"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1611
                    Top =4169
                    Width =3516
                    ColumnWidth =5250
                    ColumnOrder =10
                    TabIndex =10
                    Name ="hes"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5802
                            Top =4166
                            Width =1200
                            Height =240
                            Name ="Label22"
                            Caption ="شرح حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2437
                    Top =113
                    ColumnWidth =1170
                    ColumnOrder =1
                    TabIndex =11
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5029
                            Top =113
                            Width =810
                            Height =240
                            Name ="Label27"
                            Caption ="تاريخ سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1143
                    Top =793
                    Width =3966
                    ColumnWidth =525
                    ColumnOrder =4
                    TabIndex =1
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6200
                            Top =793
                            Width =780
                            Height =240
                            Name ="Label5"
                            Caption =" كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1133
                    Top =1190
                    Width =3966
                    ColumnWidth =585
                    ColumnOrder =5
                    TabIndex =2
                    Name ="HES_M"
                    ControlSource ="HES_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6025
                            Top =1190
                            Width =945
                            Height =240
                            Name ="Label7"
                            Caption =" معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1143
                    Top =1493
                    Width =3966
                    ColumnWidth =825
                    ColumnOrder =6
                    TabIndex =12
                    Name ="TNUMBER"
                    ControlSource ="TNUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6030
                            Top =1500
                            Width =1140
                            Height =240
                            Name ="Label29"
                            Caption =" تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3435
                    Top =4485
                    ColumnOrder =15
                    TabIndex =13
                    Name ="NO_S"
                    ControlSource ="NO_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6303
                            Top =4488
                            Width =660
                            Height =240
                            Name ="Label30"
                            Caption ="نوع سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =4755
                    Width =3966
                    ColumnWidth =870
                    ColumnOrder =16
                    TabIndex =14
                    Name ="HES_T2"
                    ControlSource ="HES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6057
                            Top =4762
                            Width =1140
                            Height =240
                            Name ="Label32"
                            Caption =" تفصيلي2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =5122
                    Width =3966
                    ColumnWidth =870
                    ColumnOrder =17
                    TabIndex =15
                    Name ="HES_T3"
                    ControlSource ="HES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6057
                            Top =5129
                            Width =1140
                            Height =240
                            Name ="Label35"
                            Caption =" تفصيلي3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =5489
                    Width =3966
                    ColumnWidth =870
                    ColumnOrder =18
                    TabIndex =16
                    Name ="HES_T4"
                    ControlSource ="HES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6057
                            Top =5496
                            Width =1140
                            Height =240
                            Name ="Label37"
                            Caption =" تفصيلي4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =5856
                    Width =3966
                    ColumnWidth =1530
                    ColumnOrder =2
                    TabIndex =17
                    Name ="HESab"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6057
                            Top =5863
                            Width =1140
                            Height =240
                            Name ="Label39"
                            Caption ="حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2267
                    Top =6349
                    ColumnWidth =4560
                    ColumnOrder =21
                    TabIndex =18
                    Name ="namcust"
                    ControlSource ="namcust"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4935
                            Top =6345
                            Width =870
                            Height =240
                            Name ="Label40"
                            Caption ="نام تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2267
                    Top =6689
                    ColumnWidth =1785
                    ColumnOrder =20
                    TabIndex =19
                    Name ="namemoin"
                    ControlSource ="namemoin"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4814
                            Top =6689
                            Width =855
                            Height =240
                            Name ="Label41"
                            Caption ="نام معين"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2267
                    Top =7029
                    ColumnWidth =1905
                    ColumnOrder =19
                    TabIndex =20
                    Name ="namkol"
                    ControlSource ="namkol"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5039
                            Top =7029
                            Width =630
                            Height =240
                            Name ="Label42"
                            Caption ="نام کل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =247
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1643
                    Top =6179
                    ColumnWidth =7185
                    ColumnOrder =3
                    TabIndex =21
                    Name ="namhes"
                    ControlSource ="namhes"
                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4350
                            Top =6180
                            Width =780
                            Height =240
                            Name ="Label43"
                            Caption ="نام حساب"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DEED_SERCH_CREATE.cls"
