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
    Width =6916
    RowHeight =360
    ItemSuffix =78
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xf0b17f2ffccde340
    End
    RecordSource ="SELECT NAME, SUM([1]) AS [1], SUM([2]) AS [2], SUM([3]) AS [3], SUM([4]) AS [4],"
        " SUM([5]) AS [5], SUM([6]) AS [6], SUM([7]) AS [7], SUM([8]) AS [8], SUM([9]) AS"
        " [9], SUM([10]) AS [10], SUM([11]) AS [11], SUM([12]) AS [12], SUM([13]) AS [13]"
        ", SUM([14]) AS [14], SUM([15]) AS [15], SUM([16]) AS [16], SUM([17]) AS [17], SU"
        "M([18]) AS [18], SUM([19]) AS [19], SUM([20]) AS [20], SUM([21]) AS [21], SUM([2"
        "2]) AS [22], SUM([23]) AS [23], SUM([24]) AS [24], SUM([25]) AS [25], SUM([26]) "
        "AS [26], SUM([27]) AS [27], SUM([28]) AS [28], SUM([29]) AS [29], SUM([30]) AS ["
        "30], SUM([31]) AS [31], VH, GRP, vahed, NAMES FROM (SELECT NAME, CASE DAYY WHEN "
        "1 THEN MEGHT ELSE 0 END AS [1], CASE DAYY WHEN 2 THEN MEGHT ELSE 0 END AS [2], C"
        "ASE DAYY WHEN 3 THEN MEGHT ELSE 0 END AS [3], CASE DAYY WHEN 4 THEN MEGHT ELSE 0"
        " END AS [4], CASE DAYY WHEN 5 THEN MEGHT ELSE 0 END AS [5], CASE DAYY WHEN 6 THE"
        "N MEGHT ELSE 0 END AS [6], CASE DAYY WHEN 7 THEN MEGHT ELSE 0 END AS [7], CASE D"
        "AYY WHEN 8 THEN MEGHT ELSE 0 END AS [8], CASE DAYY WHEN 9 THEN MEGHT ELSE 0 END "
        "AS [9], CASE DAYY WHEN 10 THEN MEGHT ELSE 0 END AS [10], CASE DAYY WHEN 11 THEN "
        "MEGHT ELSE 0 END AS [11], CASE DAYY WHEN 12 THEN MEGHT ELSE 0 END AS [12], CASE "
        "DAYY WHEN 13 THEN MEGHT ELSE 0 END AS [13], CASE DAYY WHEN 14 THEN MEGHT ELSE 0 "
        "END AS [14], CASE DAYY WHEN 15 THEN MEGHT ELSE 0 END AS [15], CASE DAYY WHEN 16 "
        "THEN MEGHT ELSE 0 END AS [16], CASE DAYY WHEN 17 THEN MEGHT ELSE 0 END AS [17], "
        "CASE DAYY WHEN 18 THEN MEGHT ELSE 0 END AS [18], CASE DAYY WHEN 19 THEN MEGHT EL"
        "SE 0 END AS [19], CASE DAYY WHEN 20 THEN MEGHT ELSE 0 END AS [20], CASE DAYY WHE"
        "N 21 THEN MEGHT ELSE 0 END AS [21], CASE DAYY WHEN 22 THEN MEGHT ELSE 0 END AS ["
        "22], CASE DAYY WHEN 23 THEN MEGHT ELSE 0 END AS [23], CASE DAYY WHEN 24 THEN MEG"
        "HT ELSE 0 END AS [24], CASE DAYY WHEN 25 THEN MEGHT ELSE 0 END AS [25], CASE DAY"
        "Y WHEN 26 THEN MEGHT ELSE 0 END AS [26], CASE DAYY WHEN 27 THEN MEGHT ELSE 0 END"
        " AS [27], CASE DAYY WHEN 28 THEN MEGHT ELSE 0 END AS [28], CASE DAYY WHEN 29 THE"
        "N MEGHT ELSE 0 END AS [29], CASE DAYY WHEN 30 THEN MEGHT ELSE 0 END AS [30], CAS"
        "E DAYY WHEN 31 THEN MEGHT ELSE 0 END AS [31], NAMES, vahed, GRP, VH, DAYY, date_"
        "n FROM dbo.AMAR_FROOSH_DAYLY WHERE (date_n BETWEEN 890101 AND 891229)) DERIVEDTB"
        "L GROUP BY NAME, VH, GRP, vahed, NAMES"
    Caption ="آمار فروش كالا ها به تفكيك روز"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:Cac"
        "heDetails/>\015\012  <x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  "
        "<x:DataAxisEmpty/>\015\012  <x:PivotField>\015\012   <x:Name>نام كالا</x:Name>\015"
        "\012   <x:SourceName>NAME</x:SourceName>\015\012   <x:Orientation>Column</x:Orie"
        "ntation>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Compar"
        "eOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>1</x:Name>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:Orientation>Column</x:Orient"
        "ation>\015\012   <x:Position>2</x:Position>\015\012   <x:EncodedType>adDouble</x"
        ":EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedM"
        "embersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotFiel"
        "d>\015\012   <x:Name>2</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adDouble</x:Encod"
        "edType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembers"
        "By>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>3</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x"
        ":Orientation>Row</x:Orientation>\015\012   <x:Position>2</x:Position>\015\012   "
        "<x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Uni"
        "queName</x:CompareOrderedMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotFi"
        "eld>\015\012  <x:PivotField>\015\012   <x:Name>4</x:Name>\015\012   <x:DataType>"
        "Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   "
        "<x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:P"
        "ivotField>\015\012  <x:PivotField>\015\012   <x:Name>5</x:Name>\015\012   <x:Dat"
        "aType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </"
        "x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>6</x:Name>\015\012   <x:"
        "DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>7</x:Name>\015\012  "
        " <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedTyp"
        "e>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>8</x:Name>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>9</x:Name>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>10</x:Na"
        "me>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble<"
        "/x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdere"
        "dMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>1"
        "1</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>ad"
        "Double</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compar"
        "eOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x"
        ":Name>12</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encoded"
        "Type>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x"
        ":CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>14</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:En"
        "codedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueNa"
        "me</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>15</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <"
        "x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Uniq"
        "ueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotFie"
        "ld>\015\012   <x:Name>16</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>17</x:Name>\015\012   <x:DataType>Number</x:DataType>\015"
        "\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMember"
        "sBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:"
        "PivotField>\015\012   <x:Name>18</x:Name>\015\012   <x:DataType>Number</x:DataTy"
        "pe>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrdere"
        "dMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>19</x:Name>\015\012   <x:DataType>Number</x:D"
        "ataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareO"
        "rderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>20</x:Name>\015\012   <x:DataType>Number<"
        "/x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Comp"
        "areOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFie"
        "ld>\015\012  <x:PivotField>\015\012   <x:Name>21</x:Name>\015\012   <x:DataType>"
        "Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   "
        "<x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:P"
        "ivotField>\015\012  <x:PivotField>\015\012   <x:Name>22</x:Name>\015\012   <x:Da"
        "taType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015"
        "\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>23</x:Name>\015\012 "
        "  <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedTy"
        "pe>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>24</x:Name>\015\012"
        "   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedT"
        "ype>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>25</x:Name>\015"
        "\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble</x:Enco"
        "dedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMember"
        "sBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>26</x:Na"
        "me>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adDouble<"
        "/x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrdere"
        "dMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>2"
        "7</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>ad"
        "Double</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Compar"
        "eOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x"
        ":Name>28</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Encoded"
        "Type>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x"
        ":CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012"
        "   <x:Name>29</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:En"
        "codedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueNa"
        "me</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotField>\015"
        "\012   <x:Name>30</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <"
        "x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>Uniq"
        "ueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotFie"
        "ld>\015\012   <x:Name>31</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivo"
        "tField>\015\012   <x:Name>گروه كالا</x:Name>\015\012   <x:SourceName>NAMES</x:So"
        "urceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Compa"
        "reOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFiel"
        "d>\015\012  <x:PivotField>\015\012   <x:Name>واحد</x:Name>\015\012   <x:SourceNa"
        "me>vahed</x:SourceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012"
        "   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </"
        "x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>DATE_N</x:Name>\015\012 "
        "  <x:DataType>Number</x:DataType>\015\012   <x:EncodedType>adBigInt</x:EncodedTy"
        "pe>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015"
        "\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField"
        ">\015\012   <x:Name>GRP</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012"
        "   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>"
        "UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012"
        "  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>VH</x:Name>\015\012 "
        "  <x:DataType>Integer</x:DataType>\015\012   <x:EncodedType>adInteger</x:Encoded"
        "Type>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy"
        ">\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>DAYY</x:Name>\015\012   <x:DataType>Integer</x:DataType"
        ">\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrdered"
        "MembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldLi"
        "st/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>Data</x:"
        "Name>\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Position>-1</"
        "x:Position>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012  <x:PivotF"
        "ield>\015\012   <x:Name>Count of نام كالا</x:Name>\015\012   <x:PLName>Total1</x"
        ":PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientation>Dat"
        "a</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x:ParentField>"
        "نام كالا</x:ParentField>\015\012   <x:Function>Count</x:Function>\015\012  </x:P"
        "ivotField>\015\012  <x:PivotField>\015\012   <x:Name>Count of نام كالا (2)</x:Na"
        "me>\015\012   <x:PLName>Total2</x:PLName>\015\012   <x:TotalNumber>1</x:TotalNum"
        "ber>\015\012   <x:Orientation>Data</x:Orientation>\015\012   <x:Position>2</x:Po"
        "sition>\015\012   <x:ParentField>نام كالا</x:ParentField>\015\012   <x:Function>"
        "Count</x:Function>\015\012  </x:PivotField>\015\012  <x:PivotData>\015\012   <x:"
        "Top>0.0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Left>0.0.0"
        "</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:SeqNum>16</x:Seq"
        "Num>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>"
        "\015\012   <x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:C"
        "aption>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>"
        "false</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012"
        "   <x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </"
        "x:PivotTable>\015\012</xml>"
    PivotChartXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:ChartSpace>\015"
        "\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  <x:Width>21061<"
        "/x:Width>\015\012  <x:Height>10689</x:Height>\015\012  <x:AllowPropertyBrowser/>"
        "\015\012  <x:DataSource>\015\012   <x:Type>PivotList</x:Type>\015\012  </x:DataS"
        "ource>\015\012  <x:BoundSeries>\015\012   <x:DataSourceIndex>0</x:DataSourceInde"
        "x>\015\012  </x:BoundSeries>\015\012  <x:Category>\015\012   <x:DataSourceIndex>"
        "0</x:DataSourceIndex>\015\012  </x:Category>\015\012  <x:Value>\015\012   <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012  </x:Value>\015\012  <x:BoundCharts>\015"
        "\012   <x:DataSourceIndex>0</x:DataSourceIndex>\015\012  </x:BoundCharts>\015\012"
        "  <x:FormatValue>\015\012   <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012   "
        "<x:Data>2</x:Data>\015\012  </x:FormatValue>\015\012  <x:PivotAggOrientation>Ser"
        "ies</x:PivotAggOrientation>\015\012  <x:Palette>\015\012   <x:Entry>#000000</x:E"
        "ntry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
        "\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012  "
        " <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:En"
        "try>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#0"
        "00000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000<"
        "/x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Ent"
        "ry>\015\012   <x:Entry>#000000</x:Entry>\015\012   <x:Entry>#000000</x:Entry>\015"
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
        "      <x:Name>كت و شلوار پشمي سايز 48 - 0 - Count of نام كالا</x:Name>\015\012  "
        "    <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012 "
        "      <x:Data>&quot;كت و شلوار پشمي سايز 48 - 0 - Count of نام كالا&quot;</x:Dat"
        "a>\015\012      </x:Caption>\015\012      <x:Index>0</x:Index>\015\012      <x:C"
        "ategory>\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x"
        ":Category>\015\012      <x:Value>\015\012       <x:DataSourceIndex>0</x:DataSour"
        "ceIndex>\015\012      </x:Value>\015\012      <x:FormatValue>\015\012       <x:D"
        "ataSourceIndex>-3</x:DataSourceIndex>\015\012       <x:Data>2</x:Data>\015\012  "
        "    </x:FormatValue>\015\012      <x:Marker>\015\012       <x:Symbol>None</x:Sym"
        "bol>\015\012      </x:Marker>\015\012      <x:Explode>0</x:Explode>\015\012     "
        " <x:Thickness>10</x:Thickness>\015\012      <x:DataSourceIndex>0</x:DataSourceIn"
        "dex>\015\012      <x:Identifier>!.كت و شلوار پشمي سايز 48._8.Count of نام كالا</"
        "x:Identifier>\015\012     </x:Series>\015\012     <x:Series>\015\012      <x:For"
        "matMap>\015\012      </x:FormatMap>\015\012      <x:Name>كت و شلوار پشمي سايز 48"
        " - 0 - Count of نام كالا (2)</x:Name>\015\012      <x:Caption>\015\012       <x:"
        "DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;كت و شلوار پش"
        "مي سايز 48 - 0 - Count of نام كالا (2)&quot;</x:Data>\015\012      </x:Caption>\015"
        "\012      <x:Index>1</x:Index>\015\012      <x:Category>\015\012       <x:DataSo"
        "urceIndex>0</x:DataSourceIndex>\015\012      </x:Category>\015\012      <x:Value"
        ">\015\012       <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>"
        "\015\012      <x:FormatValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceI"
        "ndex>\015\012       <x:Data>2</x:Data>\015\012      </x:FormatValue>\015\012    "
        "  <x:Marker>\015\012       <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015"
        "\012      <x:Explode>0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015"
        "\012      <x:DataSourceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!."
        "كت و شلوار پشمي سايز 48._8.Count of نام كالا (2)</x:Identifier>\015\012     </x:"
        "Series>\015\012     <x:Dimension>\015\012      <x:ScaleID>259470036</x:ScaleID>\015"
        "\012      <x:Index>Categories</x:Index>\015\012     </x:Dimension>\015\012     <"
        "x:Dimension>\015\012      <x:ScaleID>259470256</x:ScaleID>\015\012      <x:Index"
        ">Value</x:Index>\015\012     </x:Dimension>\015\012     <x:Dimension>\015\012   "
        "   <x:ScaleID>259470476</x:ScaleID>\015\012      <x:Index>FormatValue</x:Index>\015"
        "\012     </x:Dimension>\015\012     <x:Overlap>0</x:Overlap>\015\012     <x:GapW"
        "idth>150</x:GapWidth>\015\012     <x:FirstSliceAngle>0</x:FirstSliceAngle>\015\012"
        "    </x:Graph>\015\012    <x:Axis>\015\012     <x:AxisID>259473692</x:AxisID>\015"
        "\012     <x:ScaleID>259470036</x:ScaleID>\015\012     <x:Type>Category</x:Type>\015"
        "\012     <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:Min"
        "orTick>\015\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Ar"
        "ial</x:FontName>\015\012       <x:Size>10</x:Size>\015\012      </x:Font>\015\012"
        "      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015\012      </"
        "x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>-1</x:DataSo"
        "urceIndex>\015\012       <x:Data>&quot;Axis Title&quot;</x:Data>\015\012      </"
        "x:Caption>\015\012      <x:Position>Bottom</x:Position>\015\012     </x:Title>\015"
        "\012     <x:Placement>Bottom</x:Placement>\015\012     <x:GroupingEnum>Auto</x:G"
        "roupingEnum>\015\012    </x:Axis>\015\012    <x:Axis>\015\012     <x:AxisID>2594"
        "74440</x:AxisID>\015\012     <x:ScaleID>259470256</x:ScaleID>\015\012     <x:Typ"
        "e>Value</x:Type>\015\012     <x:Number>\015\012      <x:FormatString>General</x:"
        "FormatString>\015\012     </x:Number>\015\012     <x:MajorGridlines>\015\012    "
        " </x:MajorGridlines>\015\012     <x:MajorTick>Outside</x:MajorTick>\015\012     "
        "<x:MinorTick>None</x:MinorTick>\015\012     <x:Title>\015\012      <x:Font>\015\012"
        "       <x:FontName>Arial</x:FontName>\015\012       <x:Size>10</x:Size>\015\012 "
        "     </x:Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:Color"
        "Index>\015\012      </x:Border>\015\012      <x:Caption>\015\012       <x:DataSo"
        "urceIndex>-1</x:DataSourceIndex>\015\012       <x:Data>&quot;Axis Title&quot;</x"
        ":Data>\015\012      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012"
        "     </x:Title>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>"
        "\015\012   </x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Ch"
        "art>\015\012  <x:Scaling>\015\012   <x:ScaleID>259470036</x:ScaleID>\015\012  </"
        "x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>259470256</x:ScaleID>\015\012"
        "  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>259470476</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:HasSelectionFeedback>True</x:HasSelectionFeedback"
        ">\015\012  <x:HidePassiveAlert/>\015\012 </x:ChartSpace>\015\012</xml>"
    AllowFormView =0
    RecordSource ="SELECT NAME, SUM([1]) AS [1], SUM([2]) AS [2], SUM([3]) AS [3], SUM([4]) AS [4],"
        " SUM([5]) AS [5], SUM([6]) AS [6], SUM([7]) AS [7], SUM([8]) AS [8], SUM([9]) AS"
        " [9], SUM([10]) AS [10], SUM([11]) AS [11], SUM([12]) AS [12], SUM([13]) AS [13]"
        ", SUM([14]) AS [14], SUM([15]) AS [15], SUM([16]) AS [16], SUM([17]) AS [17], SU"
        "M([18]) AS [18], SUM([19]) AS [19], SUM([20]) AS [20], SUM([21]) AS [21], SUM([2"
        "2]) AS [22], SUM([23]) AS [23], SUM([24]) AS [24], SUM([25]) AS [25], SUM([26]) "
        "AS [26], SUM([27]) AS [27], SUM([28]) AS [28], SUM([29]) AS [29], SUM([30]) AS ["
        "30], SUM([31]) AS [31], VH, GRP, vahed, NAMES FROM (SELECT NAME, CASE DAYY WHEN "
        "1 THEN MEGHT ELSE 0 END AS [1], CASE DAYY WHEN 2 THEN MEGHT ELSE 0 END AS [2], C"
        "ASE DAYY WHEN 3 THEN MEGHT ELSE 0 END AS [3], CASE DAYY WHEN 4 THEN MEGHT ELSE 0"
        " END AS [4], CASE DAYY WHEN 5 THEN MEGHT ELSE 0 END AS [5], CASE DAYY WHEN 6 THE"
        "N MEGHT ELSE 0 END AS [6], CASE DAYY WHEN 7 THEN MEGHT ELSE 0 END AS [7], CASE D"
        "AYY WHEN 8 THEN MEGHT ELSE 0 END AS [8], CASE DAYY WHEN 9 THEN MEGHT ELSE 0 END "
        "AS [9], CASE DAYY WHEN 10 THEN MEGHT ELSE 0 END AS [10], CASE DAYY WHEN 11 THEN "
        "MEGHT ELSE 0 END AS [11], CASE DAYY WHEN 12 THEN MEGHT ELSE 0 END AS [12], CASE "
        "DAYY WHEN 13 THEN MEGHT ELSE 0 END AS [13], CASE DAYY WHEN 14 THEN MEGHT ELSE 0 "
        "END AS [14], CASE DAYY WHEN 15 THEN MEGHT ELSE 0 END AS [15], CASE DAYY WHEN 16 "
        "THEN MEGHT ELSE 0 END AS [16], CASE DAYY WHEN 17 THEN MEGHT ELSE 0 END AS [17], "
        "CASE DAYY WHEN 18 THEN MEGHT ELSE 0 END AS [18], CASE DAYY WHEN 19 THEN MEGHT EL"
        "SE 0 END AS [19], CASE DAYY WHEN 20 THEN MEGHT ELSE 0 END AS [20], CASE DAYY WHE"
        "N 21 THEN MEGHT ELSE 0 END AS [21], CASE DAYY WHEN 22 THEN MEGHT ELSE 0 END AS ["
        "22], CASE DAYY WHEN 23 THEN MEGHT ELSE 0 END AS [23], CASE DAYY WHEN 24 THEN MEG"
        "HT ELSE 0 END AS [24], CASE DAYY WHEN 25 THEN MEGHT ELSE 0 END AS [25], CASE DAY"
        "Y WHEN 26 THEN MEGHT ELSE 0 END AS [26], CASE DAYY WHEN 27 THEN MEGHT ELSE 0 END"
        " AS [27], CASE DAYY WHEN 28 THEN MEGHT ELSE 0 END AS [28], CASE DAYY WHEN 29 THE"
        "N MEGHT ELSE 0 END AS [29], CASE DAYY WHEN 30 THEN MEGHT ELSE 0 END AS [30], CAS"
        "E DAYY WHEN 31 THEN MEGHT ELSE 0 END AS [31], NAMES, vahed, GRP, VH, DAYY, date_"
        "n FROM dbo.AMAR_FROOSH_DAYLY WHERE (date_n BETWEEN 890101 AND 891229)) DERIVEDTB"
        "L GROUP BY NAME, VH, GRP, vahed, NAMES"
    RecordSourceQualifier ="dbo"
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
            Height =13165
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3902
                    Top =113
                    Width =2490
                    ColumnWidth =2565
                    ColumnOrder =0
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2095
                            Top =120
                            Width =540
                            Height =240
                            Name ="Label1"
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
                    IMESentenceMode =3
                    Left =3902
                    Top =453
                    ColumnWidth =405
                    ColumnOrder =1
                    TabIndex =1
                    Name ="1"
                    ControlSource ="1"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =453
                            Width =165
                            Height =240
                            Name ="Label3"
                            Caption ="1"
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
                    Left =3902
                    Top =793
                    ColumnWidth =405
                    ColumnOrder =2
                    TabIndex =2
                    Name ="2"
                    ControlSource ="2"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =793
                            Width =165
                            Height =240
                            Name ="Label5"
                            Caption ="2"
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
                    Left =3902
                    Top =1133
                    ColumnWidth =405
                    ColumnOrder =3
                    TabIndex =3
                    Name ="3"
                    ControlSource ="3"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =1133
                            Width =165
                            Height =240
                            Name ="Label7"
                            Caption ="3"
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
                    Left =3902
                    Top =1474
                    ColumnWidth =405
                    ColumnOrder =4
                    TabIndex =4
                    Name ="4"
                    ControlSource ="4"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =1474
                            Width =165
                            Height =240
                            Name ="Label9"
                            Caption ="4"
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
                    Left =3902
                    Top =1814
                    ColumnWidth =405
                    ColumnOrder =5
                    TabIndex =5
                    Name ="5"
                    ControlSource ="5"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl5"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =1814
                            Width =165
                            Height =240
                            Name ="Label11"
                            Caption ="5"
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
                    Left =3902
                    Top =2154
                    ColumnWidth =405
                    ColumnOrder =6
                    TabIndex =6
                    Name ="6"
                    ControlSource ="6"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl6"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =2154
                            Width =165
                            Height =240
                            Name ="Label13"
                            Caption ="6"
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
                    Left =3902
                    Top =2494
                    ColumnWidth =405
                    ColumnOrder =7
                    TabIndex =7
                    Name ="7"
                    ControlSource ="7"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl7"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =2494
                            Width =165
                            Height =240
                            Name ="Label15"
                            Caption ="7"
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
                    Left =3902
                    Top =2834
                    ColumnWidth =405
                    ColumnOrder =8
                    TabIndex =8
                    Name ="8"
                    ControlSource ="8"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl8"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =2834
                            Width =165
                            Height =240
                            Name ="Label17"
                            Caption ="8"
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
                    Left =3902
                    Top =3174
                    ColumnWidth =405
                    ColumnOrder =9
                    TabIndex =9
                    Name ="9"
                    ControlSource ="9"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl9"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =3174
                            Width =165
                            Height =240
                            Name ="Label19"
                            Caption ="9"
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
                    Left =3902
                    Top =3514
                    ColumnWidth =405
                    ColumnOrder =10
                    TabIndex =10
                    Name ="10"
                    ControlSource ="10"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl10"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =3514
                            Width =255
                            Height =240
                            Name ="Label21"
                            Caption ="10"
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
                    Left =3902
                    Top =3855
                    ColumnWidth =405
                    ColumnOrder =11
                    TabIndex =11
                    Name ="11"
                    ControlSource ="11"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl11"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =3855
                            Width =255
                            Height =240
                            Name ="Label23"
                            Caption ="11"
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
                    Left =3902
                    Top =4195
                    ColumnWidth =405
                    ColumnOrder =12
                    TabIndex =12
                    Name ="12"
                    ControlSource ="12"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl12"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =4195
                            Width =255
                            Height =240
                            Name ="Label25"
                            Caption ="12"
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
                    Left =3902
                    Top =4535
                    ColumnWidth =405
                    ColumnOrder =14
                    TabIndex =13
                    Name ="14"
                    ControlSource ="14"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl14"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =4535
                            Width =255
                            Height =240
                            Name ="Label27"
                            Caption ="14"
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
                    Left =3902
                    Top =4875
                    ColumnWidth =405
                    ColumnOrder =15
                    TabIndex =14
                    Name ="15"
                    ControlSource ="15"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl15"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =4875
                            Width =255
                            Height =240
                            Name ="Label29"
                            Caption ="15"
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
                    Left =3902
                    Top =5215
                    ColumnWidth =405
                    ColumnOrder =16
                    TabIndex =15
                    Name ="16"
                    ControlSource ="16"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl16"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =5215
                            Width =255
                            Height =240
                            Name ="Label31"
                            Caption ="16"
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
                    Left =3902
                    Top =5555
                    ColumnWidth =405
                    ColumnOrder =17
                    TabIndex =16
                    Name ="17"
                    ControlSource ="17"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl17"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =5555
                            Width =255
                            Height =240
                            Name ="Label33"
                            Caption ="17"
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
                    Left =3902
                    Top =5896
                    ColumnWidth =405
                    ColumnOrder =18
                    TabIndex =17
                    Name ="18"
                    ControlSource ="18"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl18"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =5896
                            Width =255
                            Height =240
                            Name ="Label35"
                            Caption ="18"
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
                    Left =3902
                    Top =6236
                    ColumnWidth =405
                    ColumnOrder =19
                    TabIndex =18
                    Name ="19"
                    ControlSource ="19"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl19"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =6236
                            Width =255
                            Height =240
                            Name ="Label37"
                            Caption ="19"
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
                    Left =3902
                    Top =6576
                    ColumnWidth =405
                    ColumnOrder =20
                    TabIndex =19
                    Name ="20"
                    ControlSource ="20"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl20"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =6576
                            Width =255
                            Height =240
                            Name ="Label39"
                            Caption ="20"
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
                    Left =3902
                    Top =6916
                    ColumnWidth =405
                    ColumnOrder =21
                    TabIndex =20
                    Name ="21"
                    ControlSource ="21"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl21"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =6916
                            Width =255
                            Height =240
                            Name ="Label41"
                            Caption ="21"
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
                    Left =3902
                    Top =7256
                    ColumnWidth =405
                    ColumnOrder =22
                    TabIndex =21
                    Name ="22"
                    ControlSource ="22"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl22"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =7256
                            Width =255
                            Height =240
                            Name ="Label43"
                            Caption ="22"
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
                    Left =3902
                    Top =7596
                    ColumnWidth =405
                    ColumnOrder =23
                    TabIndex =22
                    Name ="23"
                    ControlSource ="23"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl23"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =7596
                            Width =255
                            Height =240
                            Name ="Label45"
                            Caption ="23"
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
                    Left =3902
                    Top =7937
                    ColumnWidth =405
                    ColumnOrder =24
                    TabIndex =23
                    Name ="24"
                    ControlSource ="24"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl24"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =7937
                            Width =255
                            Height =240
                            Name ="Label47"
                            Caption ="24"
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
                    Left =3902
                    Top =8277
                    ColumnWidth =405
                    ColumnOrder =25
                    TabIndex =24
                    Name ="25"
                    ControlSource ="25"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl25"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =8277
                            Width =255
                            Height =240
                            Name ="Label49"
                            Caption ="25"
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
                    Left =3902
                    Top =8617
                    ColumnWidth =405
                    ColumnOrder =26
                    TabIndex =25
                    Name ="26"
                    ControlSource ="26"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl26"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =8617
                            Width =255
                            Height =240
                            Name ="Label51"
                            Caption ="26"
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
                    Left =3902
                    Top =8957
                    ColumnWidth =405
                    ColumnOrder =27
                    TabIndex =26
                    Name ="27"
                    ControlSource ="27"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl27"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =8957
                            Width =255
                            Height =240
                            Name ="Label53"
                            Caption ="27"
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
                    Left =3902
                    Top =9297
                    ColumnWidth =405
                    ColumnOrder =28
                    TabIndex =27
                    Name ="28"
                    ControlSource ="28"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl28"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =9297
                            Width =255
                            Height =240
                            Name ="Label55"
                            Caption ="28"
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
                    Left =3902
                    Top =9637
                    ColumnWidth =405
                    ColumnOrder =29
                    TabIndex =28
                    Name ="29"
                    ControlSource ="29"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl29"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =9637
                            Width =255
                            Height =240
                            Name ="Label57"
                            Caption ="29"
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
                    Left =3902
                    Top =9977
                    ColumnWidth =405
                    ColumnOrder =30
                    TabIndex =29
                    Name ="30"
                    ControlSource ="30"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl30"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =9977
                            Width =255
                            Height =240
                            Name ="Label59"
                            Caption ="30"
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
                    Left =3902
                    Top =10318
                    ColumnWidth =405
                    ColumnOrder =31
                    TabIndex =30
                    Name ="31"
                    ControlSource ="31"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl31"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =10318
                            Width =255
                            Height =240
                            Name ="Label61"
                            Caption ="31"
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
                    Left =3902
                    Top =10658
                    Width =2490
                    ColumnWidth =2475
                    ColumnOrder =33
                    TabIndex =31
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2095
                            Top =10665
                            Width =645
                            Height =240
                            Name ="Label63"
                            Caption ="گروه كالا"
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
                    Left =3902
                    Top =10998
                    Width =2490
                    ColumnOrder =34
                    TabIndex =32
                    Name ="vahed"
                    ControlSource ="vahed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2088
                            Top =10998
                            Width =525
                            Height =240
                            Name ="Label65"
                            Caption ="واحد"
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
                    Left =3845
                    Top =11792
                    ColumnOrder =32
                    TabIndex =33
                    Name ="Text74"
                    ControlSource ="=[1]+[2]+[3]+[4]+[5]+[6]+[7]+[8]+[9]+[10]+[11]+[12]+[13]+[14]+[15]+[16]+[17]+[18"
                        "]+[19]+[20]+[21]+[22]+[23]+[24]+[25]+[26]+[27]+[28]+[29]+[30]+[31]"
                    Format ="Fixed"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2031
                            Top =11792
                            Width =255
                            Height =240
                            Name ="Label75"
                            Caption ="جمع"
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
                    Left =517
                    Top =4478
                    Width =1521
                    ColumnWidth =555
                    ColumnOrder =13
                    TabIndex =34
                    Name ="13"
                    ControlSource ="13"
                    Format ="Fixed"
                    EventProcPrefix ="Ctl13"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =113
                            Top =4478
                            Width =255
                            Height =240
                            Name ="Label77"
                            Caption ="13"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AMAR_FROOSH_DAYLY_MAIN.cls"
