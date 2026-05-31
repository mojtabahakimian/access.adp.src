Version =20
VersionRequired =20
Begin Form
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5612
    RowHeight =330
    ItemSuffix =27
    Left =840
    Top =1890
    Right =8610
    Bottom =5970
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    ShortcutMenuBar ="MNEMUDAR"
    RecSrcDt = Begin
        0x81364c17539be340
    End
    InputParameters ="@Forms___F_MENU_ANBAR___DT2 BIGINT =[Forms]![F_MENU_ANBAR]![DT2],@Forms___F_MENU"
        "_ANBAR___MANBAR nvarchar(10)=[FormS]![F_MENU_ANBAR]![MANBAR]"
    RecordSource ="AKMOGUDI_KOL_ANBAR"
    Caption ="موجودي  كالا ها به تفكيك انبار"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,280 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f005f005f0044005400320020004200490047004900 ,
        0x4e00540020003d005b0046006f0072006d0073005d0021005b0046005f004d00 ,
        0x45004e0055005f0041004e004200410052005d0021005b004400540032005d00 ,
        0x2c00400046006f0072006d0073005f005f005f0046005f004d0045004e005500 ,
        0x5f0041004e004200410052005f005f005f004d0041004e004200410052002000 ,
        0x6e00760061007200630068006100720028003100300029003d005b0046006f00 ,
        0x72006d0053005d0021005b0046005f004d0045004e0055005f0041004e004200 ,
        0x410052005d0021005b004d0041004e004200410052005d00
    End
    PivotTableXML ="<xml xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\015\012 <x:WorksheetOpt"
        "ions>\015\012  <x:DisplayRightToLeft/>\015\012 </x:WorksheetOptions>\015\012 <x:"
        "PivotTable>\015\012  <x:OWCVersion>10.0.0.5605         </x:OWCVersion>\015\012  "
        "<x:NoDisplayAlerts/>\015\012  <x:NoDisplayToolbar/>\015\012  <x:DisplayScreenTip"
        "s/>\015\012  <x:NoAutoFit/>\015\012  <x:Height>384</x:Height>\015\012  <x:Width>"
        "541</x:Width>\015\012  <x:NoAllowGrouping/>\015\012  <x:NoAllowFiltering/>\015\012"
        "  <x:CubeProvider>msolap.2</x:CubeProvider>\015\012  <x:CacheDetails/>\015\012  "
        "<x:Name>Microsoft Office PivotTable 10.0</x:Name>\015\012  <x:PivotField>\015\012"
        "   <x:Name>كد كالا</x:Name>\015\012   <x:SourceName>CODE</x:SourceName>\015\012 "
        "  <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy"
        ">UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Piv"
        "otField>\015\012   <x:Name>موجودي</x:Name>\015\012   <x:SourceName>MAND</x:Sourc"
        "eName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:PLDataOrientation/"
        ">\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedM"
        "embersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>MEGF</x:Name>\015\012   <x:DataType>Number</x"
        ":DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Compar"
        "eOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDisplayI"
        "nFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>"
        "SMEGH</x:Name>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:EncodedTyp"
        "e>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:Co"
        "mpareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotFie"
        "ld>\015\012  <x:PivotField>\015\012   <x:Name>في</x:Name>\015\012   <x:SourceNam"
        "e>FII</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012   <x:Enco"
        "dedType>adDouble</x:EncodedType>\015\012   <x:NumberFormat>Standard</x:NumberFor"
        "mat>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>"
        "\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Name>مبلغ موجودي<"
        "/x:Name>\015\012   <x:SourceName>MABLK</x:SourceName>\015\012   <x:DataType>Numb"
        "er</x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:N"
        "umberFormat>Standard</x:NumberFormat>\015\012   <x:CompareOrderedMembersBy>Uniqu"
        "eName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:PivotFiel"
        "d>\015\012   <x:Name>نام كالا</x:Name>\015\012   <x:SourceName>NAME</x:SourceNam"
        "e>\015\012   <x:Orientation>Row</x:Orientation>\015\012   <x:EncodedType>adVarWC"
        "har</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:CompareOr"
        "deredMembersBy>\015\012   <x:Expanded/>\015\012  </x:PivotField>\015\012  <x:Piv"
        "otField>\015\012   <x:Name>نام واحد</x:Name>\015\012   <x:SourceName>NAMES</x:So"
        "urceName>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:Compa"
        "reOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotFiel"
        "d>\015\012  <x:PivotField>\015\012   <x:Name>KK</x:Name>\015\012   <x:EncodedTyp"
        "e>adVarChar</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>UniqueName</x:C"
        "ompareOrderedMembersBy>\015\012   <x:NoDisplayInFieldList/>\015\012  </x:PivotFi"
        "eld>\015\012  <x:PivotField>\015\012   <x:Name>كد انبار</x:Name>\015\012   <x:So"
        "urceName>ANBAR</x:SourceName>\015\012   <x:DataType>Number</x:DataType>\015\012 "
        "  <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:CompareOrderedMembersBy>U"
        "niqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012  <x:Pivot"
        "Field>\015\012   <x:Name>انبار</x:Name>\015\012   <x:SourceName>ANBARN</x:Source"
        "Name>\015\012   <x:EncodedType>adVarWChar</x:EncodedType>\015\012   <x:CompareOr"
        "deredMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015"
        "\012  <x:PivotField>\015\012   <x:Name>كد كالا1</x:Name>\015\012   <x:SourceName"
        ">VCOD</x:SourceName>\015\012   <x:DataType>Integer</x:DataType>\015\012   <x:PLC"
        "aption>كد كالا</x:PLCaption>\015\012   <x:FilterCaption>كد كالا</x:FilterCaption"
        ">\015\012   <x:EncodedType>adInteger</x:EncodedType>\015\012   <x:CompareOrdered"
        "MembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012  </x:PivotField>\015\012"
        "  <x:PivotField>\015\012   <x:Name>GHEMAT</x:Name>\015\012   <x:DataType>Number<"
        "/x:DataType>\015\012   <x:EncodedType>adDouble</x:EncodedType>\015\012   <x:Comp"
        "areOrderedMembersBy>UniqueName</x:CompareOrderedMembersBy>\015\012   <x:NoDispla"
        "yInFieldList/>\015\012  </x:PivotField>\015\012  <x:PivotField>\015\012   <x:Nam"
        "e>Data</x:Name>\015\012   <x:Orientation>Column</x:Orientation>\015\012   <x:Pos"
        "ition>-1</x:Position>\015\012   <x:DataField/>\015\012  </x:PivotField>\015\012 "
        " <x:PivotField>\015\012   <x:Name>Sum of موجودي</x:Name>\015\012   <x:PLName>Tot"
        "al1</x:PLName>\015\012   <x:TotalNumber>0</x:TotalNumber>\015\012   <x:Orientati"
        "on>Data</x:Orientation>\015\012   <x:Position>1</x:Position>\015\012   <x:Parent"
        "Field>موجودي</x:ParentField>\015\012  </x:PivotField>\015\012  <x:PivotData>\015"
        "\012   <x:Top>0.0</x:Top>\015\012   <x:TopOffset>0</x:TopOffset>\015\012   <x:Le"
        "ft>0</x:Left>\015\012   <x:LeftOffset>0</x:LeftOffset>\015\012   <x:InvertedRowM"
        "ember>!.پارچه انگليسي پشمي</x:InvertedRowMember>\015\012   <x:SeqNum>17</x:SeqNu"
        "m>\015\012  </x:PivotData>\015\012  <x:PivotView>\015\012   <x:IsNotFiltered/>\015"
        "\012   <x:Label>\015\012    <x:Caption>Microsoft Office PivotTable 10.0</x:Capti"
        "on>\015\012    <x:NotVisible/>\015\012   </x:Label>\015\012   <x:AllowEdits>fals"
        "e</x:AllowEdits>\015\012   <x:AllowAdditions>false</x:AllowAdditions>\015\012   "
        "<x:AllowDeletions>false</x:AllowDeletions>\015\012  </x:PivotView>\015\012 </x:P"
        "ivotTable>\015\012</xml>"
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
        "/x:DefaultFont>\015\012  <x:Border>\015\012   <x:Color>#2F4F4F</x:Color>\015\012"
        "  </x:Border>\015\012  <x:Interior>\015\012   <x:Color>#FFFFFF</x:Color>\015\012"
        "   <x:FillEffect>\015\012    <x:fill x:type=\"Solid\" x:color=\"#FFFFFF\"/>\015\012"
        "   </x:FillEffect>\015\012  </x:Interior>\015\012  <x:Chart>\015\012   <x:PlotAr"
        "ea>\015\012    <x:Graph>\015\012     <x:SubType>Clustered</x:SubType>\015\012   "
        "  <x:Type>Column</x:Type>\015\012     <x:Series>\015\012      <x:FormatMap>\015\012"
        "      </x:FormatMap>\015\012      <x:Name>Series</x:Name>\015\012      <x:Captio"
        "n>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data"
        ">&quot;Series&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Index>0</"
        "x:Index>\015\012      <x:Category>\015\012       <x:DataSourceIndex>0</x:DataSou"
        "rceIndex>\015\012      </x:Category>\015\012      <x:Value>\015\012       <x:Dat"
        "aSourceIndex>0</x:DataSourceIndex>\015\012      </x:Value>\015\012      <x:Forma"
        "tValue>\015\012       <x:DataSourceIndex>-3</x:DataSourceIndex>\015\012       <x"
        ":Data>2</x:Data>\015\012      </x:FormatValue>\015\012      <x:Marker>\015\012  "
        "     <x:Symbol>None</x:Symbol>\015\012      </x:Marker>\015\012      <x:Explode>"
        "0</x:Explode>\015\012      <x:Thickness>10</x:Thickness>\015\012      <x:DataSou"
        "rceIndex>0</x:DataSourceIndex>\015\012      <x:Identifier>!.Sum of موجودي</x:Ide"
        "ntifier>\015\012     </x:Series>\015\012     <x:Dimension>\015\012      <x:Scale"
        "ID>241196260</x:ScaleID>\015\012      <x:Index>Categories</x:Index>\015\012     "
        "</x:Dimension>\015\012     <x:Dimension>\015\012      <x:ScaleID>241196888</x:Sc"
        "aleID>\015\012      <x:Index>Value</x:Index>\015\012     </x:Dimension>\015\012 "
        "    <x:Dimension>\015\012      <x:ScaleID>241197092</x:ScaleID>\015\012      <x:"
        "Index>FormatValue</x:Index>\015\012     </x:Dimension>\015\012     <x:Overlap>0<"
        "/x:Overlap>\015\012     <x:GapWidth>150</x:GapWidth>\015\012     <x:FirstSliceAn"
        "gle>0</x:FirstSliceAngle>\015\012    </x:Graph>\015\012    <x:Axis>\015\012     "
        "<x:AxisID>241200192</x:AxisID>\015\012     <x:ScaleID>241196260</x:ScaleID>\015\012"
        "     <x:Type>Category</x:Type>\015\012     <x:Font>\015\012      <x:FontName>Tah"
        "oma</x:FontName>\015\012      <x:B>Automatic</x:B>\015\012      <x:I>Automatic</"
        "x:I>\015\012      <x:U>Automatic</x:U>\015\012     </x:Font>\015\012     <x:Majo"
        "rTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015\012 "
        "    <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Tahoma</x:FontName"
        ">\015\012       <x:Size>8</x:Size>\015\012       <x:Color>#000000</x:Color>\015\012"
        "       <x:B>Automatic</x:B>\015\012       <x:I>Automatic</x:I>\015\012       <x:"
        "U>Automatic</x:U>\015\012      </x:Font>\015\012      <x:Border>\015\012       <"
        "x:ColorIndex>None</x:ColorIndex>\015\012      </x:Border>\015\012      <x:Captio"
        "n>\015\012       <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012       <x:Data"
        ">&quot;نام كالا&quot;</x:Data>\015\012      </x:Caption>\015\012      <x:Positio"
        "n>Bottom</x:Position>\015\012     </x:Title>\015\012     <x:Placement>Bottom</x:"
        "Placement>\015\012     <x:GroupingEnum>Auto</x:GroupingEnum>\015\012    </x:Axis"
        ">\015\012    <x:Axis>\015\012     <x:AxisID>241201588</x:AxisID>\015\012     <x:"
        "ScaleID>241196888</x:ScaleID>\015\012     <x:Type>Value</x:Type>\015\012     <x:"
        "Number>\015\012      <x:FormatString>General</x:FormatString>\015\012     </x:Nu"
        "mber>\015\012     <x:MajorGridlines>\015\012     </x:MajorGridlines>\015\012    "
        " <x:MajorTick>Outside</x:MajorTick>\015\012     <x:MinorTick>None</x:MinorTick>\015"
        "\012     <x:Title>\015\012      <x:Font>\015\012       <x:FontName>Tahoma</x:Fon"
        "tName>\015\012       <x:Size>8</x:Size>\015\012       <x:B>Automatic</x:B>\015\012"
        "       <x:I>Automatic</x:I>\015\012       <x:U>Automatic</x:U>\015\012      </x:"
        "Font>\015\012      <x:Border>\015\012       <x:ColorIndex>None</x:ColorIndex>\015"
        "\012      </x:Border>\015\012      <x:Caption>\015\012       <x:DataSourceIndex>"
        "-1</x:DataSourceIndex>\015\012       <x:Data>&quot;موجودي&quot;</x:Data>\015\012"
        "      </x:Caption>\015\012      <x:Position>Left</x:Position>\015\012     </x:Ti"
        "tle>\015\012     <x:Placement>Left</x:Placement>\015\012    </x:Axis>\015\012   "
        "</x:PlotArea>\015\012   <x:Identifier></x:Identifier>\015\012  </x:Chart>\015\012"
        "  <x:Title>\015\012   <x:Font>\015\012    <x:FontName>Sina</x:FontName>\015\012 "
        "   <x:Color>#228B22</x:Color>\015\012    <x:B>Automatic</x:B>\015\012    <x:I>Au"
        "tomatic</x:I>\015\012    <x:U>Automatic</x:U>\015\012   </x:Font>\015\012   <x:B"
        "order>\015\012    <x:ColorIndex>None</x:ColorIndex>\015\012   </x:Border>\015\012"
        "   <x:Caption>\015\012    <x:DataSourceIndex>-1</x:DataSourceIndex>\015\012    <"
        "x:Data>&quot;نمودار موجودي كالا ها&quot;</x:Data>\015\012   </x:Caption>\015\012"
        "  </x:Title>\015\012  <x:Scaling>\015\012   <x:ScaleID>241196260</x:ScaleID>\015"
        "\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>241196888</x:ScaleI"
        "D>\015\012  </x:Scaling>\015\012  <x:Scaling>\015\012   <x:ScaleID>241197092</x:"
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
            Height =4398
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =960
                    Top =56
                    Width =2490
                    ColumnWidth =780
                    ColumnOrder =2
                    Name ="CODE"
                    ControlSource ="CODE"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4821
                            Top =63
                            Width =510
                            Height =240
                            Name ="Label1"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1749
                    Top =396
                    ColumnWidth =1035
                    ColumnOrder =5
                    TabIndex =1
                    Name ="MAND"
                    ControlSource ="MAND"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4806
                            Top =393
                            Width =645
                            Height =240
                            Name ="Label3"
                            Caption ="موجودي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2097
                    Width =3450
                    ColumnWidth =2730
                    ColumnOrder =3
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4821
                            Top =2103
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =660
                    Top =2437
                    Width =2790
                    ColumnWidth =1185
                    ColumnOrder =4
                    TabIndex =3
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4691
                            Top =2437
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1746
                    Top =738
                    ColumnWidth =690
                    ColumnOrder =1
                    TabIndex =4
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4503
                            Top =738
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="كد انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1746
                    Top =1098
                    ColumnWidth =1455
                    ColumnOrder =0
                    TabIndex =5
                    Name ="ANBARN"
                    ControlSource ="ANBARN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4503
                            Top =1098
                            Width =645
                            Height =240
                            Name ="Label20"
                            Caption ="انبار"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =876
                    Top =2777
                    Width =2490
                    ColumnWidth =1417
                    TabIndex =6
                    Name ="VCOD"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4737
                            Top =2784
                            Width =510
                            Height =240
                            Name ="Label22"
                            Caption ="كد كالا"
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
                    Left =1417
                    Top =3458
                    ColumnWidth =960
                    TabIndex =7
                    Name ="GRCOD"
                    ControlSource ="GRCOD"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3904
                            Top =3458
                            Width =915
                            Height =240
                            Name ="Label23"
                            Caption ="كد گروه كالا"
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
                    Left =1417
                    Top =3798
                    ColumnWidth =2400
                    TabIndex =8
                    Name ="GRNAME"
                    ControlSource ="GRNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3874
                            Top =3798
                            Width =945
                            Height =240
                            Name ="Label24"
                            Caption ="نام گروه كالا"
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
                    Left =1417
                    Top =4158
                    TabIndex =9
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3874
                            Top =4158
                            Width =945
                            Height =240
                            Name ="Label26"
                            Caption ="شماره فني"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AK_MOGUDI_ANBAR_LIST2.cls"
