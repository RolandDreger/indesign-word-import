<?xml version="1.0" encoding="UTF-8"?>

<!--    
        
    Microsoft Word Document -> HTML -> InDesign
    (InDesign Module)
    
    Created: 30. September 2021
    Modified: 2. February 2022
    
    Author: Roland Dreger, www.rolanddreger.net
    
     
    # Notes
    
    ## InDesign Import
    
    Use indent="no" in <xsl:output> for InDesign import and 
    deactivate option »Do Not Import Contents Of Whitespace-Only Elements«. 
    
    Otherwise, there may be problems with text wrap in cells 
    with multiple paragraphs. (&#x0d;)
    
    
    ## Document Ressources
    
    InDesign sometimes crashes with copy-of therefore the construct
    document($document-file-name) that always exits xsl:choose and 
    xsl:copy-of for global paramerters
    
-->

<xsl:transform 
    xmlns:xml="http://www.w3.org/XML/1998/namespace" 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
    xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
    xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
    xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
    xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
    xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
    xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
    xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:rel="http://schemas.openxmlformats.org/package/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    xmlns:dc="http://purl.org/dc/elements/1.1/" 
    xmlns:dcterms="http://purl.org/dc/terms/" 
    xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"
    xmlns:rd="http://www.rolanddreger.net"
    xmlns:aid="http://ns.adobe.com/AdobeInDesign/4.0/"
    xmlns:aid5="http://ns.adobe.com/AdobeInDesign/5.0/"
    exclude-result-prefixes="rd pkg wpc cx cx1 cx2 cx3 cx4 cx5 cx6 cx7 cx8 mc aink am3d o r rel m v wp14 wp w10 w w14 w15 w16cex w16cid w16 w16sdtdh w16se wpg wpi wne wps cp dc dcterms dcmitype dcmitype a pic xsi b"
    version="1.0"
>
    
    
    <xsl:import href="docx2html.xsl"/>


    <!-- ++++++++++++ -->
    <!-- + Settings + -->
    <!-- ++++++++++++ -->
    
    <xsl:param name="ns" select="''"/> <!-- Document Namespace -->
    <xsl:param name="directory-separator" select="'/'"/>
    <xsl:param name="language" select="'en'"/>
    <xsl:param name="max-bookmark-length" select="500"/>
    <xsl:param name="is-empty-paragraph-to-remove" select="false()"/>
    <xsl:param name="is-inline-style-to-remove-on-empty-text" select="false()"/>
    <xsl:param name="is-local-override-without-tag-to-apply" select="false()"/> <!-- Ignore all local overrides except: strong, i, em, u, superscript, subscript  -->
    <xsl:param name="is-comment-to-be-inserted" select="false()"/> <!-- Comments for Complex Fields, Tab, ... -->
    <xsl:param name="is-tab-to-be-preserved" select="true()"/>  <!-- Tab Character --> 
    <xsl:param name="is-special-local-override-to-apply" select="true()"/> <!-- Ignore all local overrides except: strong, i, em, u, superscript, subscript, small caps, caps, highlight, lang  -->
    
    <!-- Heading Style Map -->
    <xsl:param name="h1-paragraph-style-names" select="''"/> <!-- e.g. '»Custom_Name_1« »Custom_Name_1.1«' -->
    <xsl:param name="h2-paragraph-style-names" select="''"/> <!-- e.g. 'Custom_Name_2' -->
    <xsl:param name="h3-paragraph-style-names" select="''"/> 
    <xsl:param name="h4-paragraph-style-names" select="''"/> 
    <xsl:param name="h5-paragraph-style-names" select="''"/> 
    <xsl:param name="h6-paragraph-style-names" select="''"/> 
    
    <xsl:variable name="heading-marker" select="'-Heading-'"/>
    
    <!-- Case conversion -->
    <xsl:variable name="lowercase" select="'abcdefghijklmnopqrstuvwxyzàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿžšœ'" />
    <xsl:variable name="uppercase" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞŸŽŠŒ'" />
    
    
    <!-- +++++++++ -->
    <!-- + INPUT + -->
    <!-- +++++++++ -->
    
    <!-- Folder and File Paths -->
    <xsl:param name="package-base-uri" select="''"/> <!-- for Word-XML-Document an empty string -->
    <xsl:param name="document-file-name" select="'document.xml'"/> <!-- document.xml or name of Word-XML-Document -->
    <xsl:param name="image-folder-path" select="''"/> <!-- If image folder path is defined, all images get the path according to this pattern: $image-folder-path + '/' + $image-name  -->
    <xsl:param name="core-props-file-path" select="$document-file-name"/> <!-- ../docProps/core.xml -->
    <xsl:param name="styles-file-path" select="$document-file-name"/> <!-- styles.xml --> 
    <xsl:param name="numbering-file-path" select="$document-file-name"/> <!-- numbering.xml -->
    <xsl:param name="footnotes-file-path" select="$document-file-name"/> <!-- footnotes.xml --> 
    <xsl:param name="endnotes-file-path" select="$document-file-name"/> <!-- endnotes.xml --> 
    <xsl:param name="comments-file-path" select="$document-file-name"/> <!-- comments.xml --> 
    <xsl:param name="document-relationships-file-path" select="$document-file-name"/> <!-- _rels/document.xml.rels --> 
    <xsl:param name="footnotes-relationships-file-path" select="$document-file-name"/> <!-- _rels/footnotes.xml.rels --> 
    <xsl:param name="endnotes-relationships-file-path" select="$document-file-name"/> <!-- _rels/endnotes.xml.rels --> 
    
    <!-- Core Properties (Metadata) -->
    <xsl:variable name="core-props" select="document($core-props-file-path)/cp:coreProperties | /pkg:package/pkg:part[@pkg:name = '/docProps/core.xml']/pkg:xmlData/cp:coreProperties"/>
    
    <!-- Styles -->
    <xsl:variable name="styles" select="document($styles-file-path)/w:styles | /pkg:package/pkg:part[@pkg:name = '/word/styles.xml']/pkg:xmlData/w:styles"/>
    
    <!-- Numbering -->
    <xsl:variable name="numbering" select="document($numbering-file-path)/w:numbering | /pkg:package/pkg:part[@pkg:name = '/word/numbering.xml']/pkg:xmlData/w:numbering"/>
    
    <!-- Document Relationships (e.g. Hyperlinks) -->
    <xsl:variable name="document-relationships" select="document($document-relationships-file-path)/rel:Relationships | /pkg:package/pkg:part[@pkg:name = '/word/_rels/document.xml.rels']/pkg:xmlData/rel:Relationships"/>
    
    <!-- Footnotes -->
    <xsl:variable name="footnotes" select="document($footnotes-file-path)/w:footnotes | /pkg:package/pkg:part[@pkg:name = '/word/footnotes.xml']/pkg:xmlData/w:footnotes"/>
    
    <!-- Footnote Relationships -->
    <xsl:variable name="footnotes-relationships" select="document($footnotes-relationships-file-path)/rel:Relationships | /pkg:package/pkg:part[@pkg:name = '/word/_rels/footnotes.xml.rels']/pkg:xmlData/rel:Relationships"/>
        
    <!-- Endnotes -->
    <xsl:variable name="endnotes" select="document($endnotes-file-path)/w:endnotes | /pkg:package/pkg:part[@pkg:name = '/word/endnotes.xml']/pkg:xmlData/w:endnotes"/>
    
    <!-- Endnote Relationships -->
    <xsl:variable name="endnotes-relationships" select="document($endnotes-relationships-file-path)/rel:Relationships | /pkg:package/pkg:part[@pkg:name = '/word/_rels/endnotes.xml.rels']/pkg:xmlData/rel:Relationships"/>
    
    <!-- Comments -->
    <xsl:variable name="comments" select="document($comments-file-path)/w:comments | /pkg:package/pkg:part[@pkg:name = '/word/comments.xml']/pkg:xmlData/w:comments"/>
    
    <!-- Citations -->
    <xsl:variable name="citations-relationships">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($document-relationships-file-path)">
                <xsl:variable name="relationships-document" select="document($document-relationships-file-path)"/>
                <xsl:for-each select="$relationships-document/rel:Relationships/rel:Relationship">
                    <xsl:if test="contains(@Type, 'customXml')">
                        <xsl:variable name="custom-xml-file-path" select="concat($package-base-uri, $directory-separator, substring-after(@Target, '../'))"/>
                        <xsl:variable name="sources-element" select="document($custom-xml-file-path)/b:Sources"/>
                        <xsl:if test="$sources-element and namespace-uri($sources-element) = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'">
                            <xsl:apply-templates select="document($custom-xml-file-path)/b:Sources" mode="citation-sources"/>
                        </xsl:if>
                    </xsl:if>
                </xsl:for-each>  
            </xsl:when>
            <xsl:otherwise>
                <xsl:apply-templates select="/pkg:package/pkg:part[contains(@pkg:name, 'customXml')]/pkg:xmlData/b:Sources[namespace-uri() = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography']" mode="citation-sources"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Identity Transform for all Citation Sources -->
    <xsl:template match="@*|node()" mode="citation-sources">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()" mode="citation-sources"/>
        </xsl:copy>
    </xsl:template>
    
    
    <!-- ++++++++++ -->
    <!-- + OUTPUT + -->
    <!-- ++++++++++ -->
    
    <xsl:output method="xml" version="1.0" doctype-public="" doctype-system="" media-type="text/xml" omit-xml-declaration="no" indent="no"/>

    <!-- Output tag and attribute names -->
    <xsl:variable name="root-tag-name" select="'Document'"/>
    <xsl:variable name="paragraph-tag-name" select="'Absatz'"/>
    <xsl:variable name="paragraph-id-attribute-name" select="'ID'"/>
    <xsl:variable name="paragraph-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="list-item-level-attribute-name" select="'Ebene'"/>
    <xsl:variable name="list-id-attribute-name" select="'Liste'"/>
    <xsl:variable name="table-tag-name" select="'Tabelle'"/>
    <xsl:variable name="table-index-attribute-name" select="'Index'"/>
    <xsl:variable name="table-style-attribute-name" select="'tablestyle'"/>
    <xsl:variable name="table-marker-attribute-name" select="'table'"/>
    <xsl:variable name="table-rows-attribute-name" select="'trows'"/>
    <xsl:variable name="table-columns-attribute-name" select="'tcols'"/>
    <xsl:variable name="table-cell-tag-name" select="'Zelle'"/>
    <xsl:variable name="table-header-attribute-name" select="'theader'"/>
    <xsl:variable name="table-column-number-attribute-name" select="'ccols'"/>
    <xsl:variable name="table-row-number-attribute-name" select="'crows'"/>
    <xsl:variable name="footnote-tag-name" select="'Fussnote'"/>
    <xsl:variable name="footnote-reference-tag-name" select="'Fussnoten_Referenz'"/>
    <xsl:variable name="footnote-index-attribute-name" select="'Index'"/>
    <xsl:variable name="footnote-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="footnote-style-attribute-value" select="'Fussnote'"/>
    <xsl:variable name="endnote-tag-name" select="'Endnote'"/>
    <xsl:variable name="endnote-reference-tag-name" select="'Endnoten_Referenz'"/>
    <xsl:variable name="endnote-index-attribute-name" select="'Index'"/>
    <xsl:variable name="endnote-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="endnote-style-attribute-value" select="'Endnote'"/>
    <xsl:variable name="comment-tag-name" select="'Kommentar'"/>
    <xsl:variable name="comment-reference-tag-name" select="'Kommentar_Referenz'"/>
    <xsl:variable name="comment-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="comment-style-attribute-value" select="'comment'"/>
    <xsl:variable name="comment-index-attribute-name" select="'Index'"/>
    <xsl:variable name="comment-date-attribute-name" select="'Datum'"/>
    <xsl:variable name="comment-initials-attribute-name" select="'Initialien'"/>
    <xsl:variable name="comment-author-attribute-name" select="'Autor'"/>
    <xsl:variable name="citation-tag-name" select="'Zitat'"/>
    <xsl:variable name="citation-call-tag-name" select="'Zitataufruf'"/>
    <xsl:variable name="citation-source-tag-name" select="'Zitatquelle'"/>
    <xsl:variable name="citation-style-type-attribute-name" select="'Formattyp'"/>
    <xsl:variable name="citation-style-name-attribute-name" select="'Formatname'"/>
    <xsl:variable name="citation-version-attribute-name" select="'Version'"/>
    <xsl:variable name="citation-text-tag-name" select="'Text'"/>
    <xsl:variable name="citation-value-attribute-name" select="'Value'"/>
    <xsl:variable name="group-tag-name" select="'Gruppe'"/>
    <xsl:variable name="group-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="group-style-attribute-value" select="'Gruppe'"/>
    <xsl:variable name="image-tag-name" select="'Bild'"/>
    <xsl:variable name="image-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="image-style-attribute-value" select="'Bild_im_Textfluss'"/>
    <xsl:variable name="image-source-attribute-name" select="'Quelle'"/>
    <xsl:variable name="image-title-attribute-name" select="'Titel'"/>
    <xsl:variable name="image-alt-attribute-name" select="'Beschreibung'"/>
    <xsl:variable name="image-position-attribute-name" select="'Position'"/>
    <xsl:variable name="image-uri-attribute-name" select="'URI'"/>
    <xsl:variable name="textbox-tag-name" select="'Textbox'"/>
    <xsl:variable name="textbox-index-attribute-name" select="'Index'"/>
    <xsl:variable name="textbox-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="textbox-style-attribute-value" select="'Textbox'"/>
    <xsl:variable name="textbox-name-attribute-name" select="'Name'"/>
    <xsl:variable name="textbox-alt-attribute-name" select="'Beschreibung'"/>
    <xsl:variable name="shape-tag-name" select="'Vectorform'"/>
    <xsl:variable name="shape-index-attribute-name" select="'Index'"/>
    <xsl:variable name="shape-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="shape-style-attribute-value" select="'Vectorform'"/>
    <xsl:variable name="shape-name-attribute-name" select="'Name'"/>
    <xsl:variable name="shape-alt-attribute-name" select="'Beschreibung'"/>
    <xsl:variable name="embedded-object-tag-name" select="'Eingebettetes_Objekt'"/>
    <xsl:variable name="embedded-object-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="embedded-object-style-attribute-value" select="'Eingebettetes_Objekt'"/>
    <xsl:variable name="embedded-object-target-attribute-name" select="'Ziel'"/>
    <xsl:variable name="embedded-object-program-attribute-name" select="'Programm'"/>
    <xsl:variable name="bookmark-tag-name" select="'Lesezeichen'"/>
    <xsl:variable name="bookmark-index-attribute-name" select="'Index'"/>
    <xsl:variable name="bookmark-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="bookmark-style-attribute-value" select="'Lesezeichen'"/>
    <xsl:variable name="bookmark-id-attribute-name" select="'ID'"/>
    <xsl:variable name="bookmark-content-attribute-name" select="'Inhalt'"/>
    <xsl:variable name="indexmark-tag-name" select="'Indexmarke'"/>
    <xsl:variable name="indexmark-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="indexmark-style-attribute-value" select="'Indexmarke'"/>
    <xsl:variable name="indexmark-type-attribute-name" select="'Typ'"/>
    <xsl:variable name="indexmark-format-attribute-name" select="'Format'"/>
    <xsl:variable name="indexmark-entry-attribute-name" select="'Inhalt'"/>
    <xsl:variable name="indexmark-target-attribute-name" select="'Ziel'"/>
    <xsl:variable name="complex-field-tag-name" select="'Komplexes_Feld'"/>
    <xsl:variable name="complex-field-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="complex-field-style-attribute-value" select="'Komplexes_Feld'"/>
    <xsl:variable name="complex-field-content-attribute-name" select="'Inhalt'"/>
    <xsl:variable name="complex-field-data-attribute-name" select="'Daten'"/>
    <xsl:variable name="hyperlink-tag-name" select="'Hyperlink'"/>
    <xsl:variable name="hyperlink-uri-attribute-name" select="'URI'"/>
    <xsl:variable name="hyperlink-title-attribute-name" select="'Titel'"/>
    <xsl:variable name="cross-reference-tag-name" select="'Querverweis'"/>
    <xsl:variable name="cross-reference-uri-attribute-name" select="'URI'"/>
    <xsl:variable name="cross-reference-type-attribute-name" select="'Typ'"/>
    <xsl:variable name="cross-reference-format-attribute-name" select="'Format'"/>
    <xsl:variable name="inline-style-tag-name" select="'Zeichenformat'"/>
    <xsl:variable name="inline-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="page-break-tag-name" select="'Seitenumbruch'"/>
    <xsl:variable name="column-break-tag-name" select="'Spaltenumbruch'"/>
    <xsl:variable name="forced-line-break-tag-name" select="'Harter_Zeilenumbruch'"/>
    <xsl:variable name="carriage-return-tag-name" select="'Harter_Zeilenumbruch'"/>
    <xsl:variable name="section-break-tag-name" select="'Abschnittwechsel'"/>
    <xsl:variable name="element-type-attribute-name" select="'Typ'"/>
    <xsl:variable name="inserted-text-tag-name" select="'Eingefügter_Text'"/>
    <xsl:variable name="deleted-text-tag-name" select="'Gelöschter_Text'"/>
    <xsl:variable name="moved-to-text-tag-name" select="'Verschobener_Text'"/>
    <xsl:variable name="moved-from-text-tag-name" select="'Gelöschter_Text'"/>
    <xsl:variable name="track-change-author-attribute-name" select="'Autor'"/>
    <xsl:variable name="track-change-date-attribute-name" select="'Datum'"/>
    <xsl:variable name="subdocument-tag-name" select="'Teildokument'"/>
    <xsl:variable name="subdocument-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="subdocument-style-attribute-value" select="'Teildokument'"/>
    <xsl:variable name="subdocument-uri-attribute-name" select="'URI'"/>
    <xsl:variable name="equation-tag-name" select="'Mathematische_Gleichung'"/>
    <xsl:variable name="equation-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="equation-style-attribute-value" select="'Mathematische_Gleichung'"/>
    <xsl:variable name="time-tag-name" select="'Zeitangabe'"/>
    <xsl:variable name="time-format-attribute-name" select="'Format'"/>
    <xsl:variable name="time-type-attribute-name" select="'Typ'"/>
    <xsl:variable name="data-tag-name" select="'Daten'"/>
    <xsl:variable name="data-value-attribute-name" select="'Wert'"/>
    <xsl:variable name="symbol-tag-name" select="'Symbol'"/>
    <xsl:variable name="symbol-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="symbol-style-attribute-value" select="'Symbol'"/>
    <xsl:variable name="symbol-font-attribute-name" select="'Schrift'"/>
    <xsl:variable name="symbol-code-attribute-name" select="'Unicode'"/>
    <xsl:variable name="tab-tag-name" select="'Tabulator'"/>
    <xsl:variable name="tab-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="tab-style-attribute-value" select="'Tabulator'"/>
    
    <!-- Style Names -->
    <xsl:variable name="default-paragraph-style-name" select="'Standard'"/>
    
    
    <!-- Spaces -->
    <xsl:preserve-space elements="w:t"/>
    <xsl:strip-space elements="pkg:package pkg:part pkg:xmlData w:document w:body w:p w:pPr w:rPr w:r w:sectPr"/>
    


    <!-- +++++++++++++ -->
    <!-- + Templates + -->
    <!-- +++++++++++++ -->

    <!-- Root -->
    <xsl:template match="/">
        <xsl:element name="{$root-tag-name}" namespace="{$ns}">
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>


    <!-- Body -->
    <xsl:template match="w:body">
        <xsl:apply-templates/>
    </xsl:template>




    <!-- Paragraph -->
    <xsl:template match="w:p">
        <!-- Check: Remove empty paragraph? -->
        <xsl:if test="not($is-empty-paragraph-to-remove) or ($is-empty-paragraph-to-remove and boolean(normalize-space(.)))">
            <xsl:element name="{$paragraph-tag-name}" namespace="{$ns}">
                <xsl:call-template name="insert-paragraph-attributes"/>
                <!-- Structure text runs (e.g. complex fields) -->
                <xsl:call-template name="structure-text-runs">
                    <xsl:with-param name="target-elements" select="*"/>
                </xsl:call-template>
            </xsl:element>
            <xsl:if test="(following-sibling::w:p or following-sibling::w:tbl) and not(parent::w:footnote or parent::w:endnote or parent::w:comment)">
                <xsl:text>&#x0d;</xsl:text> <!-- Return einfuegen -->
            </xsl:if>
        </xsl:if>
    </xsl:template>

    <!-- Attributes for Paragraph -->
    <xsl:template name="insert-paragraph-attributes">
        <xsl:variable name="p-style-id" select="w:pPr/w:pStyle/@w:val"/>
        <xsl:variable name="list-id" select="w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="list-item-level" select="w:pPr/w:numPr/w:ilvl/@w:val"/>
        <!-- ID -->
        <xsl:attribute name="{$paragraph-id-attribute-name}">
            <xsl:value-of select="@w14:paraId"/>
        </xsl:attribute>
        <!-- Type -->
        <xsl:attribute name="{$element-type-attribute-name}">
            <xsl:choose>
                <xsl:when test="parent::w:footnote or parent::w:endnote or parent::w:comment">
                    <xsl:value-of select="'inline'"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="'block'"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
        <!-- Paragraph Style -->
        <xsl:variable name="style-attribute-name">
            <xsl:choose>
                <xsl:when test="parent::w:footnote or parent::w:endnote or parent::w:comment">
                    <xsl:value-of select="$paragraph-style-attribute-name"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="concat('aid',':', $paragraph-style-attribute-name)"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="{$style-attribute-name}">
            <xsl:choose>
                <xsl:when test="boolean($p-style-id)">
                    <xsl:choose>
                        <!-- Registered Style Name + List Item Level -->
                        <xsl:when test="boolean(number($list-item-level))">
                            <xsl:value-of select="concat($p-style-id, '-', $list-item-level)"/>
                        </xsl:when>
                        <!-- Registered Style Name -->
                        <xsl:otherwise>
                            <xsl:value-of select="$p-style-id"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>
                <!-- Default Style Name -->
                <xsl:otherwise>
                    <xsl:value-of select="$default-paragraph-style-name"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
        <!-- List ID -->
        <xsl:if test="boolean($list-id)">
            <xsl:attribute name="{$list-id-attribute-name}">
                <xsl:value-of select="$list-id"/>
            </xsl:attribute>
        </xsl:if>
        <!-- List Item Level -->
        <xsl:if test="boolean($list-item-level)">
            <xsl:attribute name="{$list-item-level-attribute-name}">
                <xsl:value-of select="$list-item-level"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>


    <!-- Table -->
    <xsl:template match="w:tbl">
        <xsl:element name="{$table-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-table-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
        <xsl:if test="(following-sibling::w:p or following-sibling::w:tbl) and not(parent::w:footnote or parent::w:endnote)">
            <xsl:text>&#x0d;</xsl:text> <!-- Return einfuegen -->
        </xsl:if>
    </xsl:template>

    <!-- Table Attributes -->
    <xsl:template name="insert-table-attributes">
        <xsl:attribute name="{$table-index-attribute-name}">
            <xsl:value-of select="count(preceding-sibling::w:tbl | ancestor::w:tbl) + 1"/>
        </xsl:attribute>
        <xsl:attribute name="{concat('aid',':',$table-marker-attribute-name)}">
            <xsl:value-of select="'table'"/>
        </xsl:attribute>
        <xsl:attribute name="{concat('aid',':',$table-rows-attribute-name)}">
            <xsl:value-of select="count(w:tr)"/>
        </xsl:attribute>
        <xsl:attribute name="{concat('aid',':',$table-columns-attribute-name)}">
            <xsl:value-of select="count(w:tblGrid/w:gridCol)"/>
        </xsl:attribute>
        <xsl:attribute name="{concat('aid5',':',$table-style-attribute-name)}">
            <xsl:value-of select="w:tblPr/w:tblStyle/@w:val"/>
        </xsl:attribute>
    </xsl:template>

    <!-- Table Column Group -->
    <xsl:template match="w:tblGrid">
        <xsl:apply-templates/>
    </xsl:template>

    <!-- Table Column -->
    <xsl:template match="w:gridCol">
        <xsl:apply-templates/>
    </xsl:template>

    <!-- Table Row -->
    <xsl:template match="w:tr">
        <xsl:apply-templates/>
    </xsl:template>

    <!-- Table Cell -->
    <xsl:template match="w:tc">
        <xsl:element name="{$table-cell-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-cell-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

    <!-- Merged Table Cell -->
    <xsl:template match="w:tc[w:tcPr/w:vMerge[not(@w:val)]]">
        <!-- Skip merged cell --> 
    </xsl:template>

    <!-- Table Cell Attributes -->
    <xsl:template name="insert-cell-attributes">
        <!-- Type -->
        <xsl:attribute name="{concat('aid',':',$table-marker-attribute-name)}">
            <xsl:value-of select="'cell'"/>
        </xsl:attribute>
        <!-- Header Marker -->
        <xsl:if test="parent::w:tr/w:trPr/w:tblHeader">
            <xsl:attribute name="{concat('aid',':',$table-header-attribute-name)}">
                <xsl:value-of select="''"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Column Span -->
        <xsl:choose>
            <xsl:when test="w:tcPr/w:gridSpan">
                <xsl:attribute name="{concat('aid',':',$table-column-number-attribute-name)}">
                    <xsl:value-of select="w:tcPr/w:gridSpan/@w:val"/>
                </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
                <xsl:attribute name="{concat('aid',':',$table-column-number-attribute-name)}">
                    <xsl:value-of select="1"/>
                </xsl:attribute>
            </xsl:otherwise>
        </xsl:choose>
        <!-- Row Span -->
        <xsl:choose>
            <xsl:when test="w:tcPr/w:vMerge[@w:val='restart']">
                <xsl:variable name="column-num">
                    <xsl:call-template name="get-column-number"/>
                </xsl:variable>
                <xsl:attribute name="{concat('aid',':',$table-row-number-attribute-name)}">
                    <xsl:call-template name="get-rowspan-value">
                        <xsl:with-param name="following-row-elements" select="parent::w:tr/following-sibling::w:tr"/>
                        <xsl:with-param name="target-column-num" select="$column-num"/>
                        <xsl:with-param name="num-of-rowspans" select="1"/>
                    </xsl:call-template>
                </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
                <xsl:attribute name="{concat('aid',':',$table-row-number-attribute-name)}">
                    <xsl:value-of select="1"/>
                </xsl:attribute>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>




    <!-- InDesign Character Style -->
    <xsl:template name="assign-inline-styles">
        <xsl:param name="inline-style-elements"/>
        <xsl:param name="class-names" select="''"/>
        <xsl:choose>
            <xsl:when test="$inline-style-elements">
                <xsl:variable name="target-style-element" select="$inline-style-elements[1]"/>
                <xsl:choose>
                    <!-- Ignored Character Style -->
                    <xsl:when test="
                        w:footnoteReference or
                        w:footnoteRef or
                        w:endnoteReference or
                        w:endnoteRef or
                        w:commentReference or
                        w:annotationRef or
                        w:instrText or 
                        $target-style-element[name() = 'w:noProof'] or 
                        ($is-inline-style-to-remove-on-empty-text and w:t and normalize-space(w:t) = '')
                        ">
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="$class-names"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!-- Referenced Character Style -->
                    <xsl:when test="$target-style-element[name() = 'w:rStyle']">
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="concat($target-style-element/@w:val, ' ', $class-names)"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!-- Attribute: class -->
                    <xsl:when test="
                        $target-style-element[name() = 'w:b'] or 
                        $target-style-element[name() = 'w:i'] or
                        $target-style-element[name() = 'w:u'] or
                        $target-style-element[name() = 'w:em'] or
                        $target-style-element[name() = 'w:vertAlign'] or
                        $target-style-element[name() = 'w:smallCaps'] or 
                        $target-style-element[name() = 'w:caps'] or 
                        $target-style-element[name() = 'w:highlight'] or  
                        $target-style-element[name() = 'w:lang'] or 
                        ($is-special-local-override-to-apply and 
                            (
                                $target-style-element[name() = 'w:rFonts'] or 
                                $target-style-element[name() = 'w:strike'] or 
                                $target-style-element[name() = 'w:dstrike'] or 
                                $target-style-element[name() = 'w:outline'] or 
                                $target-style-element[name() = 'w:shadow'] or 
                                $target-style-element[name() = 'w:emboss'] or 
                                $target-style-element[name() = 'w:imprint'] or 
                                $target-style-element[name() = 'w:noProof'] or 
                                $target-style-element[name() = 'w:snapToGrid'] or 
                                $target-style-element[name() = 'w:vanish'] or 
                                $target-style-element[name() = 'w:webHidden'] or 
                                $target-style-element[name() = 'w:color'] or 
                                $target-style-element[name() = 'w:spacing'] or 
                                $target-style-element[name() = 'w:w'] or 
                                $target-style-element[name() = 'w:kern'] or 
                                $target-style-element[name() = 'w:position'] or 
                                $target-style-element[name() = 'w:sz'] or 
                                $target-style-element[name() = 'w:szCs'] or 
                                $target-style-element[name() = 'w:effect'] or 
                                $target-style-element[name() = 'w:bdr'] or 
                                $target-style-element[name() = 'w:shd'] or 
                                $target-style-element[name() = 'w:fitText'] or 
                                $target-style-element[name() = 'w:rtl'] or 
                                $target-style-element[name() = 'w:cs'] or 
                                $target-style-element[name() = 'w:eastAsianLayout'] or 
                                $target-style-element[name() = 'w:specVanish'] or 
                                $target-style-element[name() = 'w:oMath']
                            )
                        )
                    ">
                        <xsl:variable name="class-name">
                            <xsl:call-template name="join-attribute-names-and-values">
                                <xsl:with-param name="local-tag-name" select="local-name($target-style-element)"/>
                                <xsl:with-param name="attributes" select="$target-style-element/@*"/>
                            </xsl:call-template>
                        </xsl:variable>
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="concat($class-name, ' ', $class-names)"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!-- No Inline Style Match -->
                    <xsl:otherwise>
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="$class-names"/>
                        </xsl:call-template>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
                <xsl:choose>
                    <!-- span Element + class Attribute  -->
                    <xsl:when test="$class-names">
                        <xsl:element name="{$inline-style-tag-name}" namespace="{$ns}">
                            <!-- Type -->
                            <xsl:attribute name="{$element-type-attribute-name}">
                                <xsl:value-of select="'inline'"/>
                            </xsl:attribute>
                            <!-- Character Style -->
                            <xsl:attribute name="{concat('aid',':', $inline-style-attribute-name)}">
                                <xsl:value-of select="normalize-space($class-names)"/>
                            </xsl:attribute>
                            <xsl:apply-templates/>
                        </xsl:element>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:apply-templates/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
   
   
    <!-- Mathematical Equation -->
    <xsl:template match="m:oMath">
        <xsl:element name="{$equation-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-equation-attributes"/>
            <xsl:copy-of select="."/>
        </xsl:element>
    </xsl:template>
    
    
    <!-- Break -->
    <xsl:template match="w:br">
        <xsl:choose>
            <xsl:when test="boolean(@w:type='page')">
                <!-- Seitenumbruch einfuegen -->
                <xsl:element name="{$page-break-tag-name}" namespace="{$ns}">
                    <!-- SpecialCharacters.PAGE_BREAK --> 
                </xsl:element>
            </xsl:when>
            <xsl:when test="boolean(@w:type='column')">
                <!-- Spaltenumbruch einfuegen -->
                <xsl:element name="{$column-break-tag-name}" namespace="{$ns}">
                    <!-- SpecialCharacters.COLUMN_BREAK --> 
                </xsl:element>
            </xsl:when>
            <xsl:otherwise>
                <xsl:element name="{$forced-line-break-tag-name}" namespace="{$ns}">
                    <!-- SpecialCharacters.FORCED_LINE_BREAK --> 
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>


    <!-- Carriage Return -->
    <xsl:template match="w:cr">
        <xsl:element name="{$carriage-return-tag-name}" namespace="{$ns}">
            <!-- SpecialCharacters.FORCED_LINE_BREAK --> 
        </xsl:element>
    </xsl:template>


    <!-- Position of Last Calculated Page Break -->
    <xsl:template match="w:lastRenderedPageBreak">
        <xsl:element name="{$section-break-tag-name}" namespace="{$ns}">
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

</xsl:transform>