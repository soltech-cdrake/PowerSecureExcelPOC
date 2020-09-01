<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:msxsl="urn:schemas-microsoft-com:xslt" exclude-result-prefixes="#default"
xmlns:html="http://www.w3.org/TR/REC-html40"
xmlns="http://www.soltech.net/PowerSecureOwnedAsset"
>
  <xsl:output method="xml" indent="no"/>
  <xsl:strip-space elements="*"/>
  <xsl:key name = "summary" match="Items/summaryTable" use="variableName"/>
  <xsl:template name="AddSheetData" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <row r="5" spans="1:4" ht="15.75">
      <c r="A5" s="13" t="s">
        <v>0</v>
      </c>
      <c r="B5" s="14" t="s">
        <v>1</v>
      </c>
      <c r="C5" s="15" t="s">
        <v>2</v>
      </c>
      <c r="D5" s="16" t="s">
        <v>3</v>
      </c>
    </row>
    <row r="6" spans="1:4">
      <c r="A6" s="17" t="s">
        <v>4</v>
      </c>
      <c r="B6" s="1" t="s">
        <v>1</v>
      </c>
      <c r="C6" s="1" t="s">
        <v>1</v>
      </c>
      <c r="D6" s="18" t="s">
        <v>1</v>
      </c>
    </row>
    <row r="7" spans="1:4">
      <c r="A7" s="19" t="s">
        <v>5</v>
      </c>
      <c r="B7" s="2" />
      <c r="C7" s="2" />
      <c r="D7" s="20" t="s">
        <v>1</v>
      </c>
    </row>
    <row r="8" spans="1:4">
      <c r="A8" s="19" t="s">
        <v>1</v>
      </c>
      <c r="B8" s="2" />
      <c r="C8" s="2" />
      <c r="D8" s="20" t="s">
        <v>1</v>
      </c>
    </row>
    
    <xsl:variable name="currentRowCount" select="8"/>
    <xsl:variable name="equipmentCount" select="count(Items/summaryTable[variableName='equipment']/tableItems)"/>
    <xsl:for-each select="key('summary', 'equipment')">
      <xsl:call-template name="AddOwnedAssetTable">
        <xsl:with-param name = "rowCount" select="$currentRowCount"/>
        <xsl:with-param name = "headerA" select="6"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>
    <xsl:variable name="assemblyCount" select="count(Items/summaryTable[variableName='assembly']/tableItems)"/> 
    <xsl:for-each select="key('summary', 'assembly')">
      <xsl:call-template name="AddOwnedAssetTable">
        <!-- Initially adding 2 to count b/c each previously inserted table adds a header and spacer-->
        <xsl:with-param name = "rowCount" select="2 + $currentRowCount + $equipmentCount"/>
        <xsl:with-param name = "headerA" select="15"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>
    <!--Add Data from the "Other" Summary table in the JSON file this represents misc estimate items-->>
    <xsl:variable name="otherCount" select="count(Items/summaryTable[variableName='other']/tableItems[not(item = 'Warranty')][not(item = 'Alliance')][not(item = 'Contingency')])"/>
    <xsl:for-each select ="key('summary', 'other')">
      <xsl:call-template name="AddOwnedAssetTable">
        <!-- Initially adding 4 to count b/c each previously inserted table adds a header and spacer-->
        <xsl:with-param name = "rowCount" select="4 + $currentRowCount + $equipmentCount + $assemblyCount"/>
        <xsl:with-param name = "headerA" select="21"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>
    <!--Add table for the PM and Engineering line items-->
    <xsl:variable name="pmAndEngineeringCount" select="count(Items/summaryTable[variableName='PME']/tableItems)"/>
    <xsl:for-each select="key('summary', 'PME')">
      <xsl:call-template name="AddOwnedAssetTable">
        <!-- Initially adding 6 to count b/c each previously inserted table adds a header and spacer-->
        <xsl:with-param name = "rowCount" select="6 + $currentRowCount + $equipmentCount + $assemblyCount + $otherCount"/>
        <xsl:with-param name = "headerA" select="30"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>
    <!--Add table for the AlliancePartnerFee line items-->
    <xsl:variable name="allianceCount" select="count(Items/summaryTable[variableName='other']/tableItems[item='Alliance'])"/>
    <xsl:for-each select="key('summary', 'other')">
      <xsl:call-template name="AddOwnedAssetAllianceTable">
        <!-- Initially adding 8 to count b/c each previously inserted table adds a header and spacer-->
        <xsl:with-param name = "rowCount" select="8 + $currentRowCount + $equipmentCount + $assemblyCount + $otherCount + $pmAndEngineeringCount"/>
        <xsl:with-param name = "headerA" select="35"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>  
    <!--Add table for the Warranty line items-->
    <xsl:variable name="warrantyCount" select="count(Items/summaryTable[variableName='other']/tableItems[item='Warranty'])"/>
    <xsl:for-each select="key('summary', 'other')">
      <xsl:call-template name="AddOwnedAssetWarrantyTable">
        <!-- Initially adding 10 to count b/c each previously inserted table adds a header and spacer-->
        <xsl:with-param name = "rowCount" select="10 + $currentRowCount + $equipmentCount + $assemblyCount + $otherCount + $pmAndEngineeringCount + $allianceCount"/>
        <xsl:with-param name = "headerA" select="36"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>
    <!--Add table for the Warranty line items-->
    <xsl:variable name="contingencyCount" select="count(Items/summaryTable[variableName='other']/tableItems[item='Contingency'])"/>
    <xsl:for-each select="key('summary', 'other')">
      <xsl:call-template name="AddOwnedAssetContingencyTable">
        <!-- Initially adding 10 to count b/c each previously inserted table adds a header and spacer-->
        <xsl:with-param name = "rowCount" select="12 + $currentRowCount + $equipmentCount + $assemblyCount + $otherCount + $pmAndEngineeringCount + $allianceCount + $warrantyCount"/>
        <xsl:with-param name = "headerA" select="38"/>
        <xsl:with-param name = "headerB" select="7"/>
        <xsl:with-param name = "headerC" select="8"/>
        <xsl:with-param name = "headerD" select="9"/>
      </xsl:call-template>
    </xsl:for-each>
    <!--Add Grand total cost summary line-->>
    <xsl:call-template name="AddOwnedAssetSummaryTable">
      <!-- Initially adding 6 to count b/c each previously inserted table adds a header and spacer-->
      <xsl:with-param name = "rowCount" select="14 + $currentRowCount + $equipmentCount + $assemblyCount + $otherCount + $pmAndEngineeringCount + $allianceCount + $warrantyCount + $contingencyCount"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="AddOwnedAssetTable" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="headerA" select="'Column A'"/>
    <xsl:param name="headerB" select="'Column B'"/>
    <xsl:param name="headerC" select="'Column C'"/>
    <xsl:param name="headerD" select="'Column D'"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <xsl:call-template name="AddOwnedAssetHeader">
      <xsl:with-param name = "rowCount" select="$rowCount"/>
      <xsl:with-param name = "headerA" select="$headerA"/>
      <xsl:with-param name = "headerB" select="$headerB"/>
      <xsl:with-param name = "headerC" select="$headerC"/>
      <xsl:with-param name = "headerD" select="$headerD"/>
    </xsl:call-template>
    <xsl:for-each select = "tableItems[not(item = 'Warranty')][not(item = 'Alliance')][not(item = 'Contingency')]">
      <xsl:call-template name="AddOwnedAssetCostTableItem">
        <xsl:with-param name="rowCount" select="(position()-1) + $insertIndex"/>
      </xsl:call-template>
    </xsl:for-each>
    <xsl:call-template name="AddOwnedAssetAdHocSpacer">
        <xsl:with-param name="rowCount" select="count(tableItems[not(item = 'Warranty')][not(item = 'Alliance')][not(item = 'Contingency')]) + $insertIndex"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="AddOwnedAssetSummaryTable" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <xsl:variable name="sumMaterial" select = "sum(Items/summaryTable/tableItems/cost)+sum(Items/summaryTable/tableItems/materialBurdon)"/>
    <xsl:variable name="sumRawLabor" select = "sum(Items/summaryTable/tableItems/raw)+sum(Items/summaryTable/tableItems/laborBurdon)"/>
    <xsl:variable name="sumDpeLabor" select = "sum(Items/summaryTable/tableItems/dpe)+sum(Items/summaryTable/tableItems/laborBurdon) + (Items/summaryTable[variableName='other']/tableItems[item='Alliance']/itemcost) + Items/summaryTable[variableName='other']/tableItems[item='Warranty']/itemcost + Items/summaryTable[variableName='other']/tableItems[item='Contingency']/itemcost"/>
    <xsl:variable name="sumTax" select = "sum(Items/summaryTable/tableItems/tax)"/>
    <xsl:call-template name="AddOwnedAssetSummaryItem">
      <xsl:with-param name = "rowCount" select="$rowCount"/>
      <xsl:with-param name = "valueA" select="40"/>
      <xsl:with-param name = "valueB" select="$sumMaterial"/>
      <xsl:with-param name = "valueC" select="$sumRawLabor + $sumDpeLabor"/>
      <xsl:with-param name = "valueD" select="$sumTax"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="AddOwnedAssetAllianceTable" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="headerA" select="'Column A'"/>
    <xsl:param name="headerB" select="'Column B'"/>
    <xsl:param name="headerC" select="'Column C'"/>
    <xsl:param name="headerD" select="'Column D'"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <xsl:call-template name="AddOwnedAssetHeader">
      <xsl:with-param name = "rowCount" select="$rowCount"/>
      <xsl:with-param name = "headerA" select="$headerA"/>
      <xsl:with-param name = "headerB" select="$headerB"/>
      <xsl:with-param name = "headerC" select="$headerC"/>
      <xsl:with-param name = "headerD" select="$headerD"/>
    </xsl:call-template>
    <xsl:for-each select = "tableItems[item='Alliance']">
      <xsl:call-template name="AddOwnedAssetLineItem">
        <xsl:with-param name = "rowCount" select="(position()-1) + $insertIndex"/>
        <xsl:with-param name = "valueA" select="item"/>
        <xsl:with-param name = "valueB" select="0.0"/>
        <xsl:with-param name = "valueC" select="itemcost"/>
        <xsl:with-param name = "valueD" select="0.0"/>
      </xsl:call-template>
    </xsl:for-each>
    <xsl:call-template name="AddOwnedAssetAdHocSpacer">
        <xsl:with-param name="rowCount" select="count(tableItems[item='Alliance']) + $insertIndex"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="AddOwnedAssetWarrantyTable" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="headerA" select="'Column A'"/>
    <xsl:param name="headerB" select="'Column B'"/>
    <xsl:param name="headerC" select="'Column C'"/>
    <xsl:param name="headerD" select="'Column D'"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <xsl:call-template name="AddOwnedAssetHeader">
      <xsl:with-param name = "rowCount" select="$rowCount"/>
      <xsl:with-param name = "headerA" select="$headerA"/>
      <xsl:with-param name = "headerB" select="$headerB"/>
      <xsl:with-param name = "headerC" select="$headerC"/>
      <xsl:with-param name = "headerD" select="$headerD"/>
    </xsl:call-template>
    <xsl:for-each select = "tableItems[item='Warranty']">
      <xsl:call-template name="AddOwnedAssetLineItem">
        <xsl:with-param name = "rowCount" select="(position()-1) + $insertIndex"/>
        <xsl:with-param name = "valueA" select="item"/>
        <xsl:with-param name = "valueB" select="0.0"/>
        <xsl:with-param name = "valueC" select="itemcost"/>
        <xsl:with-param name = "valueD" select="0.0"/>
      </xsl:call-template>
    </xsl:for-each>
    <xsl:call-template name="AddOwnedAssetAdHocSpacer">
        <xsl:with-param name="rowCount" select="count(tableItems[item='Warranty']) + $insertIndex"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="AddOwnedAssetContingencyTable" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="headerA" select="'Column A'"/>
    <xsl:param name="headerB" select="'Column B'"/>
    <xsl:param name="headerC" select="'Column C'"/>
    <xsl:param name="headerD" select="'Column D'"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <xsl:call-template name="AddOwnedAssetHeader">
      <xsl:with-param name = "rowCount" select="$rowCount"/>
      <xsl:with-param name = "headerA" select="$headerA"/>
      <xsl:with-param name = "headerB" select="$headerB"/>
      <xsl:with-param name = "headerC" select="$headerC"/>
      <xsl:with-param name = "headerD" select="$headerD"/>
    </xsl:call-template>
    <xsl:for-each select = "tableItems[item='Contingency']">
      <xsl:call-template name="AddOwnedAssetLineItem">
        <xsl:with-param name = "rowCount" select="(position()-1) + $insertIndex"/>
        <xsl:with-param name = "valueA" select="item"/>
        <xsl:with-param name = "valueB" select="0.0"/>
        <xsl:with-param name = "valueC" select="itemcost"/>
        <xsl:with-param name = "valueD" select="0.0"/>
      </xsl:call-template>
    </xsl:for-each>
    <xsl:call-template name="AddOwnedAssetAdHocSpacer">
        <xsl:with-param name="rowCount" select="count(tableItems[item='Contingency']) + $insertIndex"/>
    </xsl:call-template>
  </xsl:template>

  <xsl:template name="AddOwnedAssetHeader"  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="headerA" select="'Column A'"/>
    <xsl:param name="headerB" select="'Column B'"/>
    <xsl:param name="headerC" select="'Column C'"/>
    <xsl:param name="headerD" select="'Column D'"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <row r="{$insertIndex}" spans="1:4">
      <c r="A{$insertIndex}" s="12" t="s">
        <v>
          <xsl:value-of select="$headerA"/>
        </v>
      </c>
      <c r="B{$insertIndex}" s="3" t="s">
        <v>
          <xsl:value-of select="$headerB"/>
        </v>
      </c>
      <c r="C{$insertIndex}" s="3" t="s">
        <v>
          <xsl:value-of select="$headerC"/>
        </v>
      </c>
      <c r="D{$insertIndex}" s="21" t="s">
        <v>
          <xsl:value-of select="$headerD"/>
        </v>
      </c>
    </row>
  </xsl:template>

  <xsl:template name="AddOwnedAssetCostTableItem" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <row r="{$insertIndex}" spans="1:4">
      <c r="A{$insertIndex}" s="22">
        <is><t><xsl:value-of select="item"/></t></is>
      </c>
      <c r="B{$insertIndex}" s="6">
        <v>
          <xsl:choose>
            <xsl:when test="cost and materialBurdon">
              <xsl:variable name="cost" select="cost"/>
              <xsl:variable name="materialBurdon" select="materialBurdon"/>
              <xsl:value-of select="$cost + $materialBurdon"/>
            </xsl:when>  
            <xsl:otherwise>
              <xsl:value-of select="0"/>
            </xsl:otherwise>
          </xsl:choose>
        </v>
      </c>
      <c r="C{$insertIndex}" s="6">
        <v>
          <xsl:choose>
            <xsl:when test="laborBurdon">
              <xsl:variable name="laborBurdon" select="laborBurdon"/>
              <xsl:if test="raw">
                <xsl:variable name="raw" select="raw"/>
                <xsl:value-of select="$raw + $laborBurdon"/>
              </xsl:if>
              <xsl:if test="dpe">
                <xsl:variable name="dpe" select="dpe"/>
                <xsl:value-of select="$dpe + $laborBurdon"/>
              </xsl:if>
            </xsl:when>  
            <xsl:otherwise>
              <xsl:if test="raw">
                <xsl:value-of select="raw"/>
              </xsl:if>
              <xsl:if test="dpe">
                <xsl:value-of select="dpe"/>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </v>
      </c>
      <c r="D{$insertIndex}" s="23">
        <v>
          <xsl:value-of select="tax"/>
        </v>
      </c>
    </row>
  </xsl:template>

  <!-- TODO: Add the capability to parameterize style values and condense AddOwnedAssetSummaryItem, AddOwnedAssetLineItem, and AddOwnedAssetHeader into single template-->
  <xsl:template name="AddOwnedAssetSummaryItem"  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="valueA" select="'Item Text Not Available'"/>
    <xsl:param name="valueB" select="0"/>
    <xsl:param name="valueC" select="0"/>
    <xsl:param name="valueD" select="0"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <row r="{$insertIndex}" spans="1:4">
      <xsl:if test="number($valueA) = 'NaN'">
        <c r="A{$insertIndex}" s="39" t="inlineStr">
          <is>
            <t><xsl:value-of select="@valueA"/></t>
          </is>
        </c>
      </xsl:if>
      <xsl:if test="number($valueA) != 'NaN'">
        <c r="A{$insertIndex}" s="39" t="s">
          <v>
              <xsl:value-of select="$valueA"/>
          </v>
        </c>
      </xsl:if>
      <c r="B{$insertIndex}" s="9">
        <v>
          <xsl:value-of select="$valueB"/>
        </v>
      </c>
      <c r="C{$insertIndex}" s="10">
        <v>
          <xsl:value-of select="$valueC"/>
        </v>
      </c>
      <c r="D{$insertIndex}" s="38">
        <v>
          <xsl:value-of select="$valueD"/>
        </v>
      </c>
    </row>
  </xsl:template>

  <xsl:template name="AddOwnedAssetLineItem"  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:param name="valueA" select="'Item Text Not Available'"/>
    <xsl:param name="valueB" select="0"/>
    <xsl:param name="valueC" select="0"/>
    <xsl:param name="valueD" select="0"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <row r="{$insertIndex}" spans="1:4">
      <xsl:if test="number($valueA) = 'NaN'">
        <c r="A{$insertIndex}" s="22" t="inlineStr">
          <is>
            <t><xsl:value-of select="@valueA"/></t>
          </is>
        </c>
      </xsl:if>
      <xsl:if test="number($valueA) != 'NaN'">
        <c r="A{$insertIndex}" s="22" t="s">
          <v>
              <xsl:value-of select="$valueA"/>
          </v>
        </c>
      </xsl:if>
      <c r="B{$insertIndex}" s="6">
        <v>
          <xsl:value-of select="$valueB"/>
        </v>
      </c>
      <c r="C{$insertIndex}" s="6">
        <v>
          <xsl:value-of select="$valueC"/>
        </v>
      </c>
      <c r="D{$insertIndex}" s="23">
        <v>
          <xsl:value-of select="$valueD"/>
        </v>
      </c>
    </row>
  </xsl:template>

  <xsl:template name="AddOwnedAssetAdHocSpacer" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
    <xsl:param name="rowCount" select="0"/>
    <xsl:variable name="insertIndex" select="$rowCount + 1"/>
    <row r="{$insertIndex}" spans="1:4">
      <c r="{concat('A', $insertIndex)}" s="25" t="s">
        <v>1</v>
      </c>
      <c r="{concat('B', $insertIndex)}" s="7"/>
      <c r="{concat('C', $insertIndex)}" s="7"/>
      <c r="{concat('D', $insertIndex)}" s="26" t="s">
        <v>1</v>
      </c>
    </row>
  </xsl:template>

  <xsl:template match="/">
    <pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
      <pkg:part pkg:name="/docProps/app.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.extended-properties+xml">
        <pkg:xmlData>
          <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
            <Application>Microsoft Excel Online</Application>
            <Manager></Manager>
            <Company></Company>
            <HyperlinkBase></HyperlinkBase>
            <AppVersion>16.0300</AppVersion>
          </Properties>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/docProps/core.xml" pkg:contentType="application/vnd.openxmlformats-package.core-properties+xml">
        <pkg:xmlData>
          <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <dc:title></dc:title>
            <dc:subject></dc:subject>
            <dc:creator></dc:creator>
            <cp:keywords></cp:keywords>
            <dc:description></dc:description>
            <cp:lastModifiedBy>Chris Drake</cp:lastModifiedBy>
            <cp:revision></cp:revision>
            <dcterms:created xsi:type="dcterms:W3CDTF">2020-08-21T14:06:39Z</dcterms:created>
            <dcterms:modified xsi:type="dcterms:W3CDTF">2020-08-23T21:28:17Z</dcterms:modified>
            <cp:category></cp:category>
            <cp:contentStatus></cp:contentStatus>
          </cp:coreProperties>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/drawings/drawing1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.drawing+xml">
        <pkg:xmlData>
          <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <xdr:twoCellAnchor editAs="oneCell">
              <xdr:from>
                <xdr:col>0</xdr:col>
                <xdr:colOff>47625</xdr:colOff>
                <xdr:row>0</xdr:row>
                <xdr:rowOff>47625</xdr:rowOff>
              </xdr:from>
              <xdr:to>
                <xdr:col>0</xdr:col>
                <xdr:colOff>1771650</xdr:colOff>
                <xdr:row>3</xdr:row>
                <xdr:rowOff>47625</xdr:rowOff>
              </xdr:to>
              <xdr:pic>
                <xdr:nvPicPr>
                  <xdr:cNvPr id="2" name="Picture 1">
                    <a:extLst>
                      <a:ext uri="">
                        <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="" />
                      </a:ext>
                    </a:extLst>
                    <!--<a:extLst>
                      <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                        <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{64034014-C8B8-4123-A1D5-2A9156A98031}" />
                      </a:ext>
                    </a:extLst> -->
                  </xdr:cNvPr>
                  <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect="1" />
                  </xdr:cNvPicPr>
                </xdr:nvPicPr>
                <xdr:blipFill>
                  <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" />
                  <a:stretch>
                    <a:fillRect />
                  </a:stretch>
                </xdr:blipFill>
                <xdr:spPr>
                  <a:xfrm>
                    <a:off x="47625" y="47625" />
                    <a:ext cx="1724025" cy="571500" />
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst />
                  </a:prstGeom>
                </xdr:spPr>
              </xdr:pic>
              <xdr:clientData />
            </xdr:twoCellAnchor>
          </xdr:wsDr>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/drawings/_rels/drawing1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/media/image1.png" pkg:contentType="image/png" pkg:compression="store">
        <pkg:binaryData>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
    SUVORK5CYII=
    </pkg:binaryData>
      </pkg:part>
      <pkg:part pkg:name="/xl/sharedStrings.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml">
        <pkg:xmlData>
          <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="103" uniqueCount="43">
            <si>
              <t>OWNED ASSET</t>
            </si>
            <si>
              <t></t>
            </si>
            <si>
              <t>DATE:</t>
            </si>
            <si>
              <t>1/0/00</t>
            </si>
            <si>
              <t xml:space="preserve">RFQ NAME:   </t>
            </si>
            <si>
              <t xml:space="preserve">RFQ NO.   </t>
            </si>
            <si>
              <t>EQUIPMENT</t>
            </si>
            <si>
              <t>MATERIAL</t>
            </si>
            <si>
              <t>LABOR</t>
            </si>
            <si>
              <t>TAX</t>
            </si>
            <si>
              <t>equipment_item</t>
            </si>
            <si>
              <t>equipment_material</t>
            </si>
            <si>
              <t>equipment_labor</t>
            </si>
            <si>
              <t>equipment_tax</t>
            </si>
            <si>
              <t>SWITCHGEAR / ATS /APS</t>
            </si>
            <si>
              <t>INSTALL / ASSESMBLY</t>
            </si>
            <si>
              <t>assembly_item</t>
            </si>
            <si>
              <t>assembly_material</t>
            </si>
            <si>
              <t>assembly_labor</t>
            </si>
            <si>
              <t>assembly_tax</t>
            </si>
            <si>
              <t>SWITCHGEAR / ATS /APS INSTALL</t>
            </si>
            <si>
              <t>GC, START UP, COMMISSIONING, PERMIT, FACTORY WITNESS TEST, FREIGHT</t>
            </si>
            <si>
              <t>GENERAL CONDITIONS</t>
            </si>
            <si>
              <t xml:space="preserve"> $                                                                             -  </t>
            </si>
            <si>
              <t>START UP</t>
            </si>
            <si>
              <t>COMMISSIONING</t>
            </si>
            <si>
              <t>AIR PERMITTING</t>
            </si>
            <si>
              <t xml:space="preserve"> $                                                           -  </t>
            </si>
            <si>
              <t>CONSTRUCTION PERMIT</t>
            </si>
            <si>
              <t>FREIGHT</t>
            </si>
            <si>
              <t>PM AND ENGINEERING</t>
            </si>
            <si>
              <t>PROJECT MANAGER</t>
            </si>
            <si>
              <t>CONSTRUCTION MANAGER</t>
            </si>
            <si>
              <t>SR PROJECT MANAGER</t>
            </si>
            <si>
              <t>ENGINEERING</t>
            </si>
            <si>
              <t>ALLIANCE PARTNER FEE</t>
            </si>
            <si>
              <t>WARRANTY</t>
            </si>
            <si>
              <t>5 - YEAR WARRANTY</t>
            </si>
            <si>
              <t>CONTINGENCY</t>
            </si>
            <si>
              <t>owerSecure-</t>
            </si>
            <si>
              <t>GRAND TOTAL COST WITH TAX</t>
            </si>
            <si>
              <t>total_materials</t>
            </si>
            <si>
              <t>total_labor</t>
            </si>
          </sst>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml">
        <pkg:xmlData>
          <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
            <numFmts count="1">
              <numFmt numFmtId="8" formatCode="&quot;$&quot;#,##0.00_);[Red]\(&quot;$&quot;#,##0.00\)" />
            </numFmts>
            <fonts count="7">
              <font>
                <sz val="11" />
                <color theme="1" />
                <name val="Calibri" />
                <family val="2" />
                <scheme val="minor" />
              </font>
              <font>
                <b />
                <sz val="12" />
                <color rgb="FF000000" />
                <name val="Calibri" />
                <family val="2" />
              </font>
              <font>
                <sz val="11" />
                <color rgb="FF000000" />
                <name val="Calibri" />
                <family val="2" />
              </font>
              <font>
                <b />
                <sz val="11" />
                <color rgb="FF000000" />
                <name val="Calibri" />
                <family val="2" />
              </font>
              <font>
                <sz val="11" />
                <name val="Calibri" />
                <family val="2" />
              </font>
              <font>
                <b />
                <sz val="11" />
                <name val="Calibri" />
                <family val="2" />
              </font>
              <font>
                <b />
                <sz val="12" />
                <name val="Calibri" />
                <family val="2" />
              </font>
            </fonts>
            <fills count="5">
              <fill>
                <patternFill patternType="none" />
              </fill>
              <fill>
                <patternFill patternType="gray125" />
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <fgColor rgb="FFBFBFBF" />
                  <bgColor rgb="FF000000" />
                </patternFill>
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <fgColor rgb="FF00B0F0" />
                  <bgColor rgb="FF000000" />
                </patternFill>
              </fill>
              <fill>
                <patternFill patternType="solid">
                  <fgColor rgb="FFC6E0B4" />
                  <bgColor rgb="FF000000" />
                </patternFill>
              </fill>
            </fills>
            <borders count="17">
              <border>
                <left />
                <right />
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right />
                <top style="double">
                  <color rgb="FF808080" />
                </top>
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FF808080" />
                </right>
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FFFFFFFF" />
                </right>
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FF000000" />
                </left>
                <right />
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FF000000" />
                </left>
                <right />
                <top style="double">
                  <color rgb="FF000000" />
                </top>
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right />
                <top style="double">
                  <color rgb="FF000000" />
                </top>
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FF000000" />
                </right>
                <top style="double">
                  <color rgb="FF000000" />
                </top>
                <bottom />
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FF000000" />
                </left>
                <right />
                <top style="double">
                  <color rgb="FF808080" />
                </top>
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FF000000" />
                </right>
                <top style="double">
                  <color rgb="FF808080" />
                </top>
                <bottom />
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FF000000" />
                </right>
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FFFFFFFF" />
                </left>
                <right style="double">
                  <color rgb="FF000000" />
                </right>
                <top />
                <bottom />
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FF000000" />
                </left>
                <right style="double">
                  <color rgb="FFFFFFFF" />
                </right>
                <top />
                <bottom style="double">
                  <color rgb="FF000000" />
                </bottom>
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FFFFFFFF" />
                </left>
                <right style="double">
                  <color rgb="FFFFFFFF" />
                </right>
                <top />
                <bottom style="double">
                  <color rgb="FF000000" />
                </bottom>
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FFFFFFFF" />
                </right>
                <top />
                <bottom style="double">
                  <color rgb="FF000000" />
                </bottom>
                <diagonal />
              </border>
              <border>
                <left />
                <right style="double">
                  <color rgb="FF000000" />
                </right>
                <top />
                <bottom style="double">
                  <color rgb="FF000000" />
                </bottom>
                <diagonal />
              </border>
              <border>
                <left style="double">
                  <color rgb="FFFFFFFF" />
                </left>
                <right />
                <top />
                <bottom />
                <diagonal />
              </border>
            </borders>
            <cellStyleXfs count="1">
              <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
            </cellStyleXfs>
            <cellXfs count="43">
              <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
              <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="4" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="2" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="2" fillId="4" borderId="3" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="6" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="6" fillId="3" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="4" fillId="4" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="3" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="1" fillId="2" borderId="5" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="2" borderId="6" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="2" borderId="6" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="2" borderId="7" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="0" borderId="8" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="0" borderId="9" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="0" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="0" borderId="10" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="3" fillId="3" borderId="10" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="4" fillId="4" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="4" fillId="4" borderId="10" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="4" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="0" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="0" borderId="10" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="4" fillId="4" borderId="11" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="5" fillId="0" borderId="11" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="4" fillId="0" borderId="11" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="5" fillId="3" borderId="10" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="4" fillId="4" borderId="10" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="5" fillId="0" borderId="10" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="5" fillId="4" borderId="11" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="0" borderId="12" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="0" borderId="13" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="0" borderId="14" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="5" fillId="0" borderId="15" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="6" fillId="3" borderId="10" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="6" fillId="3" borderId="4" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="4" fillId="4" borderId="11" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="8" fontId="2" fillId="4" borderId="16" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
              <xf numFmtId="0" fontId="2" fillId="4" borderId="16" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" />
            </cellXfs>
            <cellStyles count="1">
              <cellStyle name="Normal" xfId="0" builtinId="0" />
            </cellStyles>
            <dxfs count="0" />
            <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9" />
            <extLst>
              <ext uri="" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
                <x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1" />
              </ext>
              <ext uri="" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
                <x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1" />
              </ext>
            </extLst>
          </styleSheet>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/theme/theme1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.theme+xml">
        <pkg:xmlData>
          <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
            <a:themeElements>
              <a:clrScheme name="Office">
                <a:dk1>
                  <a:sysClr val="windowText" lastClr="000000" />
                </a:dk1>
                <a:lt1>
                  <a:sysClr val="window" lastClr="FFFFFF" />
                </a:lt1>
                <a:dk2>
                  <a:srgbClr val="44546A" />
                </a:dk2>
                <a:lt2>
                  <a:srgbClr val="E7E6E6" />
                </a:lt2>
                <a:accent1>
                  <a:srgbClr val="4472C4" />
                </a:accent1>
                <a:accent2>
                  <a:srgbClr val="ED7D31" />
                </a:accent2>
                <a:accent3>
                  <a:srgbClr val="A5A5A5" />
                </a:accent3>
                <a:accent4>
                  <a:srgbClr val="FFC000" />
                </a:accent4>
                <a:accent5>
                  <a:srgbClr val="5B9BD5" />
                </a:accent5>
                <a:accent6>
                  <a:srgbClr val="70AD47" />
                </a:accent6>
                <a:hlink>
                  <a:srgbClr val="0563C1" />
                </a:hlink>
                <a:folHlink>
                  <a:srgbClr val="954F72" />
                </a:folHlink>
              </a:clrScheme>
              <a:fontScheme name="Office">
                <a:majorFont>
                  <a:latin typeface="Calibri Light" panose="020F0302020204030204" />
                  <a:ea typeface="" />
                  <a:cs typeface="" />
                  <a:font script="Jpan" typeface="游ゴシック Light" />
                  <a:font script="Hang" typeface="맑은 고딕" />
                  <a:font script="Hans" typeface="等线 Light" />
                  <a:font script="Hant" typeface="新細明體" />
                  <a:font script="Arab" typeface="Times New Roman" />
                  <a:font script="Hebr" typeface="Times New Roman" />
                  <a:font script="Thai" typeface="Tahoma" />
                  <a:font script="Ethi" typeface="Nyala" />
                  <a:font script="Beng" typeface="Vrinda" />
                  <a:font script="Gujr" typeface="Shruti" />
                  <a:font script="Khmr" typeface="MoolBoran" />
                  <a:font script="Knda" typeface="Tunga" />
                  <a:font script="Guru" typeface="Raavi" />
                  <a:font script="Cans" typeface="Euphemia" />
                  <a:font script="Cher" typeface="Plantagenet Cherokee" />
                  <a:font script="Yiii" typeface="Microsoft Yi Baiti" />
                  <a:font script="Tibt" typeface="Microsoft Himalaya" />
                  <a:font script="Thaa" typeface="MV Boli" />
                  <a:font script="Deva" typeface="Mangal" />
                  <a:font script="Telu" typeface="Gautami" />
                  <a:font script="Taml" typeface="Latha" />
                  <a:font script="Syrc" typeface="Estrangelo Edessa" />
                  <a:font script="Orya" typeface="Kalinga" />
                  <a:font script="Mlym" typeface="Kartika" />
                  <a:font script="Laoo" typeface="DokChampa" />
                  <a:font script="Sinh" typeface="Iskoola Pota" />
                  <a:font script="Mong" typeface="Mongolian Baiti" />
                  <a:font script="Viet" typeface="Times New Roman" />
                  <a:font script="Uigh" typeface="Microsoft Uighur" />
                  <a:font script="Geor" typeface="Sylfaen" />
                  <a:font script="Armn" typeface="Arial" />
                  <a:font script="Bugi" typeface="Leelawadee UI" />
                  <a:font script="Bopo" typeface="Microsoft JhengHei" />
                  <a:font script="Java" typeface="Javanese Text" />
                  <a:font script="Lisu" typeface="Segoe UI" />
                  <a:font script="Mymr" typeface="Myanmar Text" />
                  <a:font script="Nkoo" typeface="Ebrima" />
                  <a:font script="Olck" typeface="Nirmala UI" />
                  <a:font script="Osma" typeface="Ebrima" />
                  <a:font script="Phag" typeface="Phagspa" />
                  <a:font script="Syrn" typeface="Estrangelo Edessa" />
                  <a:font script="Syrj" typeface="Estrangelo Edessa" />
                  <a:font script="Syre" typeface="Estrangelo Edessa" />
                  <a:font script="Sora" typeface="Nirmala UI" />
                  <a:font script="Tale" typeface="Microsoft Tai Le" />
                  <a:font script="Talu" typeface="Microsoft New Tai Lue" />
                  <a:font script="Tfng" typeface="Ebrima" />
                </a:majorFont>
                <a:minorFont>
                  <a:latin typeface="Calibri" panose="020F0502020204030204" />
                  <a:ea typeface="" />
                  <a:cs typeface="" />
                  <a:font script="Jpan" typeface="游ゴシック" />
                  <a:font script="Hang" typeface="맑은 고딕" />
                  <a:font script="Hans" typeface="等线" />
                  <a:font script="Hant" typeface="新細明體" />
                  <a:font script="Arab" typeface="Arial" />
                  <a:font script="Hebr" typeface="Arial" />
                  <a:font script="Thai" typeface="Tahoma" />
                  <a:font script="Ethi" typeface="Nyala" />
                  <a:font script="Beng" typeface="Vrinda" />
                  <a:font script="Gujr" typeface="Shruti" />
                  <a:font script="Khmr" typeface="DaunPenh" />
                  <a:font script="Knda" typeface="Tunga" />
                  <a:font script="Guru" typeface="Raavi" />
                  <a:font script="Cans" typeface="Euphemia" />
                  <a:font script="Cher" typeface="Plantagenet Cherokee" />
                  <a:font script="Yiii" typeface="Microsoft Yi Baiti" />
                  <a:font script="Tibt" typeface="Microsoft Himalaya" />
                  <a:font script="Thaa" typeface="MV Boli" />
                  <a:font script="Deva" typeface="Mangal" />
                  <a:font script="Telu" typeface="Gautami" />
                  <a:font script="Taml" typeface="Latha" />
                  <a:font script="Syrc" typeface="Estrangelo Edessa" />
                  <a:font script="Orya" typeface="Kalinga" />
                  <a:font script="Mlym" typeface="Kartika" />
                  <a:font script="Laoo" typeface="DokChampa" />
                  <a:font script="Sinh" typeface="Iskoola Pota" />
                  <a:font script="Mong" typeface="Mongolian Baiti" />
                  <a:font script="Viet" typeface="Arial" />
                  <a:font script="Uigh" typeface="Microsoft Uighur" />
                  <a:font script="Geor" typeface="Sylfaen" />
                  <a:font script="Armn" typeface="Arial" />
                  <a:font script="Bugi" typeface="Leelawadee UI" />
                  <a:font script="Bopo" typeface="Microsoft JhengHei" />
                  <a:font script="Java" typeface="Javanese Text" />
                  <a:font script="Lisu" typeface="Segoe UI" />
                  <a:font script="Mymr" typeface="Myanmar Text" />
                  <a:font script="Nkoo" typeface="Ebrima" />
                  <a:font script="Olck" typeface="Nirmala UI" />
                  <a:font script="Osma" typeface="Ebrima" />
                  <a:font script="Phag" typeface="Phagspa" />
                  <a:font script="Syrn" typeface="Estrangelo Edessa" />
                  <a:font script="Syrj" typeface="Estrangelo Edessa" />
                  <a:font script="Syre" typeface="Estrangelo Edessa" />
                  <a:font script="Sora" typeface="Nirmala UI" />
                  <a:font script="Tale" typeface="Microsoft Tai Le" />
                  <a:font script="Talu" typeface="Microsoft New Tai Lue" />
                  <a:font script="Tfng" typeface="Ebrima" />
                </a:minorFont>
              </a:fontScheme>
              <a:fmtScheme name="Office">
                <a:fillStyleLst>
                  <a:solidFill>
                    <a:schemeClr val="phClr" />
                  </a:solidFill>
                  <a:gradFill rotWithShape="1">
                    <a:gsLst>
                      <a:gs pos="0">
                        <a:schemeClr val="phClr">
                          <a:lumMod val="110000" />
                          <a:satMod val="105000" />
                          <a:tint val="67000" />
                        </a:schemeClr>
                      </a:gs>
                      <a:gs pos="50000">
                        <a:schemeClr val="phClr">
                          <a:lumMod val="105000" />
                          <a:satMod val="103000" />
                          <a:tint val="73000" />
                        </a:schemeClr>
                      </a:gs>
                      <a:gs pos="100000">
                        <a:schemeClr val="phClr">
                          <a:lumMod val="105000" />
                          <a:satMod val="109000" />
                          <a:tint val="81000" />
                        </a:schemeClr>
                      </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0" />
                  </a:gradFill>
                  <a:gradFill rotWithShape="1">
                    <a:gsLst>
                      <a:gs pos="0">
                        <a:schemeClr val="phClr">
                          <a:satMod val="103000" />
                          <a:lumMod val="102000" />
                          <a:tint val="94000" />
                        </a:schemeClr>
                      </a:gs>
                      <a:gs pos="50000">
                        <a:schemeClr val="phClr">
                          <a:satMod val="110000" />
                          <a:lumMod val="100000" />
                          <a:shade val="100000" />
                        </a:schemeClr>
                      </a:gs>
                      <a:gs pos="100000">
                        <a:schemeClr val="phClr">
                          <a:lumMod val="99000" />
                          <a:satMod val="120000" />
                          <a:shade val="78000" />
                        </a:schemeClr>
                      </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0" />
                  </a:gradFill>
                </a:fillStyleLst>
                <a:lnStyleLst>
                  <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                      <a:schemeClr val="phClr" />
                    </a:solidFill>
                    <a:prstDash val="solid" />
                    <a:miter lim="800000" />
                  </a:ln>
                  <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                      <a:schemeClr val="phClr" />
                    </a:solidFill>
                    <a:prstDash val="solid" />
                    <a:miter lim="800000" />
                  </a:ln>
                  <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                      <a:schemeClr val="phClr" />
                    </a:solidFill>
                    <a:prstDash val="solid" />
                    <a:miter lim="800000" />
                  </a:ln>
                </a:lnStyleLst>
                <a:effectStyleLst>
                  <a:effectStyle>
                    <a:effectLst />
                  </a:effectStyle>
                  <a:effectStyle>
                    <a:effectLst />
                  </a:effectStyle>
                  <a:effectStyle>
                    <a:effectLst>
                      <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                        <a:srgbClr val="000000">
                          <a:alpha val="63000" />
                        </a:srgbClr>
                      </a:outerShdw>
                    </a:effectLst>
                  </a:effectStyle>
                </a:effectStyleLst>
                <a:bgFillStyleLst>
                  <a:solidFill>
                    <a:schemeClr val="phClr" />
                  </a:solidFill>
                  <a:solidFill>
                    <a:schemeClr val="phClr">
                      <a:tint val="95000" />
                      <a:satMod val="170000" />
                    </a:schemeClr>
                  </a:solidFill>
                  <a:gradFill rotWithShape="1">
                    <a:gsLst>
                      <a:gs pos="0">
                        <a:schemeClr val="phClr">
                          <a:tint val="93000" />
                          <a:satMod val="150000" />
                          <a:shade val="98000" />
                          <a:lumMod val="102000" />
                        </a:schemeClr>
                      </a:gs>
                      <a:gs pos="50000">
                        <a:schemeClr val="phClr">
                          <a:tint val="98000" />
                          <a:satMod val="130000" />
                          <a:shade val="90000" />
                          <a:lumMod val="103000" />
                        </a:schemeClr>
                      </a:gs>
                      <a:gs pos="100000">
                        <a:schemeClr val="phClr">
                          <a:shade val="63000" />
                          <a:satMod val="120000" />
                        </a:schemeClr>
                      </a:gs>
                    </a:gsLst>
                    <a:lin ang="5400000" scaled="0" />
                  </a:gradFill>
                </a:bgFillStyleLst>
              </a:fmtScheme>
            </a:themeElements>
            <a:objectDefaults />
            <a:extraClrSchemeLst />
            <a:extLst>
              <a:ext uri="">
              <!--<a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">-->
                <!--<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" />-->
                <thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" />
              </a:ext>
            </a:extLst>
          </a:theme>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/workbook.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml">
        <pkg:xmlData>
          <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
            <fileVersion appName="xl" lastEdited="7" lowestEdited="4" rupBuild="23219" />
            <workbookPr defaultThemeVersion="166925" />
            <xr:revisionPtr revIDLastSave="143" documentId="11_E60897F41BE170836B02CE998F75CCDC64E183C8" xr6:coauthVersionLast="45" xr6:coauthVersionMax="45" xr10:uidLastSave="{DE4ED4BD-4207-4AB3-AB87-2ADB95C81E4B}" />
            <bookViews>
              <workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010" />
              <!--<workbookView xWindow="240" yWindow="105" windowWidth="14805" windowHeight="8010" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}" />-->
            </bookViews>
            <sheets>
              <sheet name="Sheet1" sheetId="1" r:id="rId1" />
            </sheets>
            <calcPr calcId="191028" calcCompleted="0" />
            <extLst>
              <ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">
                <xcalcf:calcFeatures>
                  <xcalcf:feature name="microsoft.com:RD" />
                  <xcalcf:feature name="microsoft.com:Single" />
                  <xcalcf:feature name="microsoft.com:FV" />
                  <xcalcf:feature name="microsoft.com:CNMTM" />
                </xcalcf:calcFeatures>
              </ext>
            </extLst>
          </workbook>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/worksheets/sheet1.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml">
        <pkg:xmlData>
          <!--<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{00000000-0001-0000-0000-000000000000}">-->
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3">
            <dimension ref="A5:D17" />
            <sheetViews>
              <sheetView tabSelected="1" topLeftCell="A5" workbookViewId="0">
                <selection activeCell="C18" sqref="C18" />
              </sheetView>
            </sheetViews>
            <sheetFormatPr defaultColWidth="9.140625" defaultRowHeight="15" />
            <cols>
              <col min="1" max="1" width="70.28515625" customWidth="1" />
              <col min="2" max="2" width="30.140625" bestFit="1" customWidth="1" />
              <col min="3" max="4" width="38.140625" bestFit="1" customWidth="1" />
            </cols>
            <sheetData>
              <xsl:call-template name="AddSheetData"/>
            </sheetData>

            <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
            <drawing r:id="rId1" />
          </worksheet> 

        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/worksheets/_rels/sheet1.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/xl/_rels/workbook.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
            <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
      <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
        <pkg:xmlData>
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
            <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
            <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
          </Relationships>
        </pkg:xmlData>
      </pkg:part>
    </pkg:package>
  </xsl:template>
</xsl:stylesheet>