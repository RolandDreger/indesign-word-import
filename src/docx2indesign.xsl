<?xml version="1.0" encoding="UTF-8"?>

<!--    
        
    Microsoft Word Document -> HTML -> InDesign
    (InDesign Module)
    
    Created: September 30, 2021
    Modified: April 1, 2023
    
    Author: Roland Dreger, www.rolanddreger.net
    
     
    # Notes
    
    ## InDesign Import
    
    For InDesign import use indent="no" in <xsl:output> in this stylesheet and 
    deactivate option »Do Not Import Contents Of Whitespace-Only Elements« 
    in InDesign XML import settings. 
    
    Otherwise, there may be problems with text wrap in cells 
    with multiple paragraphs. (&#x0d;)
    
    
    ## Document Resources
    
    InDesign sometimes crashes with copy-of therefore the construct
    document($document-file-name) that always exits instead of 
    xsl:choose and xsl:copy-of for global paramerters
    
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
    xmlns:extp="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
    xmlns:cusp="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" 
    xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
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
    exclude-result-prefixes="rd pkg wpc cx cx1 cx2 cx3 cx4 cx5 cx6 cx7 cx8 mc aink am3d o r rel m v wp14 wp w10 w w14 w15 w16cex w16cid w16 w16sdtdh w16se wpg wpi wne wps cp extp cusp vt dc dcterms dcmitype dcmitype a pic xsi b"
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
    <xsl:param name="is-empty-paragraph-removed" select="false()"/>
    <xsl:param name="is-inline-style-on-empty-text-removed" select="false()"/>
    <xsl:param name="is-comment-inserted" select="false()"/> <!-- Comments for Complex Fields, Tab, ... -->
    <xsl:param name="is-tab-preserved" select="true()"/>  <!-- Tab Character --> 
    <xsl:param name="style-mode" select="'extended'"/> <!-- Values: 'extended' or 'minimized'. If 'minimized', ignore all local overrides except: strong, i, em, u, superscript, subscript, small caps, caps, highlight  -->
    <xsl:param name="table-mode" select="'table'"/> <!-- Values: 'table' or 'tabbedlist'. If 'tabbedlist', import table structure as tab separated text. -->
    <xsl:param name="fallback-paragraph-style-name" select="'Standard'"/>
    
    
    <!-- +++++++++ -->
    <!-- + INPUT + -->
    <!-- +++++++++ -->
    
    <!-- Folder and File Paths -->
    <xsl:param name="package-base-uri" select="''"/> <!-- for Word-XML-Document an empty string -->
    <xsl:param name="document-file-name" select="'document.xml'"/> <!-- document.xml or name of Word-XML-Document -->
    <xsl:param name="image-folder-path" select="''"/> <!-- If image folder path is defined, all images get the path according to this pattern: $image-folder-path + '/' + $image-name  -->
    <xsl:param name="app-props-file-path" select="$document-file-name"/> <!-- ../docProps/app.xml -->
    <xsl:param name="core-props-file-path" select="$document-file-name"/> <!-- ../docProps/core.xml -->
    <xsl:param name="custom-props-file-path" select="$document-file-name"/> <!-- ../docProps/custom.xml -->
    <xsl:param name="styles-file-path" select="$document-file-name"/> <!-- styles.xml --> 
    <xsl:param name="numbering-file-path" select="$document-file-name"/> <!-- numbering.xml -->
    <xsl:param name="footnotes-file-path" select="$document-file-name"/> <!-- footnotes.xml --> 
    <xsl:param name="endnotes-file-path" select="$document-file-name"/> <!-- endnotes.xml --> 
    <xsl:param name="comments-file-path" select="$document-file-name"/> <!-- comments.xml --> 
    <xsl:param name="document-relationships-file-path" select="$document-file-name"/> <!-- _rels/document.xml.rels --> 
    <xsl:param name="footnotes-relationships-file-path" select="$document-file-name"/> <!-- _rels/footnotes.xml.rels --> 
    <xsl:param name="endnotes-relationships-file-path" select="$document-file-name"/> <!-- _rels/endnotes.xml.rels --> 
    
    <!-- App Properties (Metadata) -->
    <xsl:variable name="app-props" select="document($app-props-file-path)/extp:Properties | /pkg:package/pkg:part[@pkg:name = '/docProps/app.xml']/pkg:xmlData/extp:Properties"/>
    
    <!-- Core Properties (Metadata) -->
    <xsl:variable name="core-props" select="document($core-props-file-path)/cp:coreProperties | /pkg:package/pkg:part[@pkg:name = '/docProps/core.xml']/pkg:xmlData/cp:coreProperties"/>
    
    <!-- Custom Properties (Metadata) -->
    <xsl:variable name="custom-props" select="document($custom-props-file-path)/cusp:Properties | /pkg:package/pkg:part[@pkg:name = '/docProps/custom.xml']/pkg:xmlData/cusp:Properties"/>
    
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
                        <xsl:variable name="custom-xml-relative-file-path" select="substring-after(@Target, '../')"/>
                        <xsl:if test="not($custom-xml-relative-file-path = '') and starts-with($custom-xml-relative-file-path, 'customXml')">
                            <xsl:variable name="custom-xml-file-path" select="concat($package-base-uri, $directory-separator, $custom-xml-relative-file-path)"/>
                            <xsl:variable name="sources-element" select="document($custom-xml-file-path)/b:Sources"/>
                            <xsl:if test="$sources-element and namespace-uri($sources-element) = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'">
                                <xsl:apply-templates select="document($custom-xml-file-path)/b:Sources" mode="citation-sources"/>
                            </xsl:if>
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
    
    <xsl:output 
        method="xml" 
        version="1.0" 
        doctype-public="" 
        doctype-system="" 
        media-type="text/xml" 
        omit-xml-declaration="no" 
        indent="no"
    />

    <!-- Output tag and attribute names -->
    <xsl:variable name="root-tag-name" select="'document'"/>
    <xsl:variable name="paragraph-tag-name" select="'paragraph'"/>
    <xsl:variable name="paragraph-id-attribute-name" select="'ID'"/>
    <xsl:variable name="paragraph-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="list-id-attribute-name" select="'list-id'"/>
    <xsl:variable name="list-item-level-attribute-name" select="'level'"/>
    <xsl:variable name="list-format-attribute-name" select="'list-format'"/>
    <xsl:variable name="list-start-attribute-name" select="'list-start'"/>
    <xsl:variable name="table-tag-name" select="'table'"/>
    <xsl:variable name="table-index-attribute-name" select="'index'"/>
    <xsl:variable name="table-style-attribute-name" select="'tablestyle'"/>
    <xsl:variable name="table-marker-attribute-name" select="'table'"/>
    <xsl:variable name="table-rows-attribute-name" select="'trows'"/>
    <xsl:variable name="table-columns-attribute-name" select="'tcols'"/>
    <xsl:variable name="table-cell-tag-name" select="'cell'"/>
    <xsl:variable name="table-header-attribute-name" select="'theader'"/>
    <xsl:variable name="table-row-number-attribute-name" select="'crows'"/>
    <xsl:variable name="table-column-number-attribute-name" select="'ccols'"/>
    <xsl:variable name="table-column-width-attribute-name" select="'ccolwidth'"/>
    <xsl:variable name="tabbed-list-tab-type-attribute-name" select="'type'"/>
    <xsl:variable name="tabbed-list-tab-type-attribute-value" select="'table'"/>
    <xsl:variable name="tabbed-list-tab-inline-style-attribute-name" select="'Tabbed_List'"/>
    <xsl:variable name="footnote-tag-name" select="'footnote'"/>
    <xsl:variable name="footnote-reference-tag-name" select="'footnoteref'"/>
    <xsl:variable name="footnote-index-attribute-name" select="'index'"/>
    <xsl:variable name="footnote-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="footnote-style-attribute-value" select="'Footnote'"/>
    <xsl:variable name="endnote-tag-name" select="'endnote'"/>
    <xsl:variable name="endnote-reference-tag-name" select="'endnoteref'"/>
    <xsl:variable name="endnote-index-attribute-name" select="'index'"/>
    <xsl:variable name="endnote-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="endnote-style-attribute-value" select="'Endnote'"/>
    <xsl:variable name="comment-tag-name" select="'comment'"/>
    <xsl:variable name="comment-reference-tag-name" select="'commentref'"/>
    <xsl:variable name="comment-style-attribute-name" select="'pstyle'"/>
    <xsl:variable name="comment-style-attribute-value" select="'Comment'"/>
    <xsl:variable name="comment-index-attribute-name" select="'index'"/>
    <xsl:variable name="comment-date-attribute-name" select="'date'"/>
    <xsl:variable name="comment-initials-attribute-name" select="'initials'"/>
    <xsl:variable name="comment-author-attribute-name" select="'author'"/>
    <xsl:variable name="citation-tag-name" select="'Zitat'"/>
    <xsl:variable name="citation-call-tag-name" select="'Zitataufruf'"/>
    <xsl:variable name="citation-source-tag-name" select="'Zitatquelle'"/>
    <xsl:variable name="citation-style-attribute-name" select="'class'"/>
    <xsl:variable name="citation-style-attribute-value" select="'Citation'"/>
    <xsl:variable name="citation-style-type-attribute-name" select="'Formattyp'"/>
    <xsl:variable name="citation-style-name-attribute-name" select="'Formatname'"/>
    <xsl:variable name="citation-version-attribute-name" select="'Version'"/>
    <xsl:variable name="citation-text-tag-name" select="'Text'"/>
    <xsl:variable name="citation-value-attribute-name" select="'Value'"/>
    <xsl:variable name="group-tag-name" select="'Gruppe'"/>
    <xsl:variable name="group-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="group-style-attribute-value" select="'Gruppe'"/>
    <xsl:variable name="image-tag-name" select="'image'"/>
    <xsl:variable name="image-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="image-style-anchored-attribute-value" select="'Anchored_Image'"/>
    <xsl:variable name="image-style-inline-attribute-value" select="'Inline_Image'"/>
    <xsl:variable name="image-source-attribute-name" select="'source'"/>
    <xsl:variable name="image-title-attribute-name" select="'title'"/>
    <xsl:variable name="image-alt-attribute-name" select="'description'"/>
    <xsl:variable name="image-position-attribute-name" select="'position'"/>
    <xsl:variable name="image-uri-attribute-name" select="'uri'"/>
    <xsl:variable name="textbox-tag-name" select="'textbox'"/>
    <xsl:variable name="textbox-index-attribute-name" select="'index'"/>
    <xsl:variable name="textbox-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="textbox-style-attribute-value" select="'Textbox'"/>
    <xsl:variable name="textbox-name-attribute-name" select="'name'"/>
    <xsl:variable name="textbox-alt-attribute-name" select="'alt'"/>
    <xsl:variable name="shape-tag-name" select="'shape'"/>
    <xsl:variable name="shape-index-attribute-name" select="'Index'"/>
    <xsl:variable name="shape-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="shape-style-attribute-value" select="'Shape'"/>
    <xsl:variable name="shape-name-attribute-name" select="'name'"/>
    <xsl:variable name="shape-alt-attribute-name" select="'description'"/>
    <xsl:variable name="embedded-object-tag-name" select="'embedded-object'"/>
    <xsl:variable name="embedded-object-style-attribute-name" select="'ostyle'"/>
    <xsl:variable name="embedded-object-style-attribute-value" select="'Embedded_Object'"/>
    <xsl:variable name="embedded-object-target-attribute-name" select="'target'"/>
    <xsl:variable name="embedded-object-program-attribute-name" select="'application'"/>
    <xsl:variable name="bookmark-tag-name" select="'bookmark'"/>
    <xsl:variable name="bookmark-index-attribute-name" select="'index'"/>
    <xsl:variable name="bookmark-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="bookmark-style-attribute-value" select="'Bookmark'"/>
    <xsl:variable name="bookmark-id-attribute-name" select="'id'"/>
    <xsl:variable name="bookmark-content-attribute-name" select="'content'"/>
    <xsl:variable name="indexmark-tag-name" select="'indexmark'"/>
    <xsl:variable name="indexmark-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="indexmark-style-attribute-value" select="'Indexmark'"/>
    <xsl:variable name="indexmark-type-attribute-name" select="'type'"/>
    <xsl:variable name="indexmark-format-attribute-name" select="'format'"/>
    <xsl:variable name="indexmark-entry-attribute-name" select="'entry'"/>
    <xsl:variable name="indexmark-target-attribute-name" select="'target'"/>
    <xsl:variable name="complex-field-tag-name" select="'complex-field'"/>
    <xsl:variable name="complex-field-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="complex-field-style-attribute-value" select="'Complex_Field'"/>
    <xsl:variable name="complex-field-content-attribute-name" select="'content'"/>
    <xsl:variable name="complex-field-data-attribute-name" select="'data'"/>
    <xsl:variable name="hyperlink-tag-name" select="'hyperlink'"/>
    <xsl:variable name="hyperlink-uri-attribute-name" select="'uri'"/>
    <xsl:variable name="hyperlink-title-attribute-name" select="'title'"/>
    <xsl:variable name="cross-reference-tag-name" select="'cross-reference'"/>
    <xsl:variable name="cross-reference-uri-attribute-name" select="'uri'"/>
    <xsl:variable name="cross-reference-type-attribute-name" select="'type'"/>
    <xsl:variable name="cross-reference-format-attribute-name" select="'format'"/>
    <xsl:variable name="inline-style-tag-name" select="'character-style'"/>
    <xsl:variable name="inline-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="page-break-tag-name" select="'pagebreak'"/>
    <xsl:variable name="column-break-tag-name" select="'columnbreak'"/>
    <xsl:variable name="forced-line-break-tag-name" select="'forcedlinebreak'"/>
    <xsl:variable name="carriage-return-tag-name" select="'carriagereturn'"/>
    <xsl:variable name="section-break-tag-name" select="'sectionbreak'"/>
    <xsl:variable name="element-type-attribute-name" select="'Typ'"/>
    <xsl:variable name="inserted-text-tag-name" select="'insertedtext'"/>
    <xsl:variable name="deleted-text-tag-name" select="'deletedtext'"/>
    <xsl:variable name="moved-to-text-tag-name" select="'movedtext'"/>
    <xsl:variable name="moved-from-text-tag-name" select="'deletedtext'"/>
    <xsl:variable name="track-change-author-attribute-name" select="'author'"/>
    <xsl:variable name="track-change-date-attribute-name" select="'date'"/>
    <xsl:variable name="subdocument-tag-name" select="'subdocument'"/>
    <xsl:variable name="subdocument-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="subdocument-style-attribute-value" select="'Subdocument'"/>
    <xsl:variable name="subdocument-uri-attribute-name" select="'uri'"/>
    <xsl:variable name="equation-tag-name" select="'equation'"/>
    <xsl:variable name="equation-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="equation-style-attribute-value" select="'Equation'"/>
    <xsl:variable name="time-tag-name" select="'time'"/>
    <xsl:variable name="time-format-attribute-name" select="'format'"/>
    <xsl:variable name="time-type-attribute-name" select="'type'"/>
    <xsl:variable name="data-tag-name" select="'data'"/>
    <xsl:variable name="data-value-attribute-name" select="'value'"/>
    <xsl:variable name="symbol-tag-name" select="'symbol'"/>
    <xsl:variable name="symbol-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="symbol-style-attribute-value" select="'Symbol'"/>
    <xsl:variable name="symbol-font-attribute-name" select="'font'"/>
    <xsl:variable name="symbol-code-attribute-name" select="'unicode'"/>
    <xsl:variable name="tab-tag-name" select="'tab'"/>
    <xsl:variable name="tab-style-attribute-name" select="'cstyle'"/>
    <xsl:variable name="tab-style-attribute-value" select="'Tabulator'"/>
    
    
    <!-- Spaces -->
    <xsl:preserve-space elements="w:t"/>
    <xsl:strip-space elements="pkg:package pkg:part pkg:xmlData w:document w:body w:p w:pPr w:rPr w:r w:sectPr"/>
    


    <!-- +++++++++++++ -->
    <!-- + Templates + -->
    <!-- +++++++++++++ -->

    <!-- Root -->
    <xsl:template match="/">
        <xsl:element name="{$root-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-metadata-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributs for Root Element -->
    <xsl:template name="insert-metadata-attributes">
        <xsl:attribute name="title">
            <xsl:value-of select="$core-props/dc:title"/>
        </xsl:attribute>
        <xsl:attribute name="subject">
            <xsl:value-of select="$core-props/dc:subject"/>
        </xsl:attribute>
        <xsl:attribute name="author">
            <xsl:value-of select="$core-props/dc:creator"/>
        </xsl:attribute>
        <xsl:attribute name="keywords">
            <xsl:value-of select="$core-props/cp:keywords"/>
        </xsl:attribute>
        <xsl:attribute name="category">
            <xsl:value-of select="$core-props/cp:category"/>
        </xsl:attribute>
        <xsl:attribute name="description">
            <xsl:value-of select="$core-props/dc:description"/>
        </xsl:attribute>
        <xsl:attribute name="created">
            <xsl:value-of select="$core-props/dcterms:created"/>
        </xsl:attribute>
        <xsl:attribute name="modified">
            <xsl:value-of select="$core-props/dcterms:modified"/>
        </xsl:attribute>
        <xsl:attribute name="application">
            <xsl:value-of select="$app-props/extp:Application"/>
        </xsl:attribute>
        <xsl:attribute name="app-version">
            <xsl:value-of select="$app-props/extp:AppVersion"/>
        </xsl:attribute>
        <xsl:if test="$custom-props">
            <xsl:for-each select="$custom-props/cusp:property">
                <xsl:if test="@name">
                    <xsl:attribute name="{@name}">
                        <xsl:value-of select="vt:lpwstr | vt:i4 | vt:bool | vt:filetime"/>
                    </xsl:attribute>
                </xsl:if>
            </xsl:for-each>
        </xsl:if>
    </xsl:template>




    <!-- Body -->
    <xsl:template match="w:body">
        <xsl:apply-templates/>
    </xsl:template>




    <!-- Paragraph -->
    <xsl:template match="w:p">
        <!-- Check: Remove empty paragraph? -->
        <xsl:if test="not($is-empty-paragraph-removed) or ($is-empty-paragraph-removed and boolean(normalize-space(.)))">
            <xsl:element name="{$paragraph-tag-name}" namespace="{$ns}">
                <xsl:call-template name="insert-paragraph-attributes"/>
                <!-- Structure text runs (e.g. complex fields) -->
                <xsl:call-template name="structure-text-runs">
                    <xsl:with-param name="target-elements" select="*"/>
                </xsl:call-template>
            </xsl:element>
            <xsl:if test="(following-sibling::w:p or following-sibling::w:tbl) and not(parent::w:comment)">
                <xsl:text>&#x0d;</xsl:text> <!-- Carriage Return -->
            </xsl:if>
        </xsl:if>
    </xsl:template>

    <!-- Attributes for Paragraph -->
    <xsl:template name="insert-paragraph-attributes">
        <!-- ID -->
        <xsl:attribute name="{$paragraph-id-attribute-name}">
            <xsl:value-of select="@w14:paraId"/>
        </xsl:attribute>
        <!-- Type -->
        <xsl:attribute name="{$element-type-attribute-name}">
            <xsl:choose>
                <xsl:when test="parent::w:footnote or parent::w:endnote or parent::w:comment or parent::w:txbxContent">
                    <xsl:value-of select="'inline'"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="'block'"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
        <!-- List Attributes -->
        <xsl:variable name="list-id" select="w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="list-item-level" select="w:pPr/w:numPr/w:ilvl/@w:val"/>
        <xsl:if test="boolean(w:pPr/w:numPr)">
            <xsl:call-template name="insert-list-attributes">
                <xsl:with-param name="list-id" select="$list-id"/>
                <xsl:with-param name="list-item-level" select="$list-item-level"/>
            </xsl:call-template>
        </xsl:if>
        <!-- Paragraph Style -->
        <xsl:variable name="paragraph-style-name">
            <xsl:call-template name="get-style-name">
                <xsl:with-param name="style-id" select="w:pPr/w:pStyle/@w:val"/>
                <xsl:with-param name="style-type" select="'paragraph'"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="indesign-paragraph-style-attribute-name">
            <xsl:choose>
                <xsl:when test="parent::w:footnote or parent::w:endnote or parent::w:comment or parent::w:txbxContent">
                    <xsl:value-of select="$paragraph-style-attribute-name"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="concat('aid',':', $paragraph-style-attribute-name)"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="paragraph-style-attribute-value">
            <xsl:choose>
                <xsl:when test="$paragraph-style-name != ''">
                    <xsl:choose>
                        <!-- Registered Style Name + List Item Level -->
                        <xsl:when test="boolean(number($list-item-level))">
                            <xsl:value-of select="concat($paragraph-style-name, '-', $list-item-level)"/>
                        </xsl:when>
                        <!-- Registered Style Name -->
                        <xsl:otherwise>
                            <xsl:value-of select="$paragraph-style-name"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>
                <!-- Default Style Name -->
                <xsl:otherwise>
                    <xsl:call-template name="get-style-name">
                        <xsl:with-param name="style-id" select="$fallback-paragraph-style-name"/>
                        <xsl:with-param name="style-type" select="'paragraph'"/>
                    </xsl:call-template>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="translated-paragraph-style-attribute-value" select="translate($paragraph-style-attribute-value,'[]','')"/>
        <xsl:attribute name="{$indesign-paragraph-style-attribute-name}">
            <xsl:value-of select="$translated-paragraph-style-attribute-value"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Insert List Attributes -->
    <xsl:template name="insert-list-attributes">
        
        <xsl:param name="list-id" select="''"/>
        <xsl:param name="list-item-level" select="''"/>
        
        <xsl:variable name="list-abstract-num-id" select="$numbering/w:num[@w:numId=$list-id]/w:abstractNumId/@w:val"/>
        <xsl:variable name="numbering-style" select="$numbering/w:abstractNum[@w:abstractNumId=$list-abstract-num-id]"/>

        <!-- List ID -->
        <xsl:attribute name="{$list-id-attribute-name}">
            <xsl:value-of select="$list-id"/>
        </xsl:attribute>
        
        <!-- List Item Level -->
        <xsl:attribute name="{$list-item-level-attribute-name}">
            <xsl:value-of select="$list-item-level"/>
        </xsl:attribute>
        
        <!-- List Format -->
        <xsl:attribute name="{$list-format-attribute-name}">
            <xsl:call-template name="get-list-format">
                <xsl:with-param name="list-item-level" select="$list-item-level"/>
                <xsl:with-param name="numbering-style" select="$numbering-style"/>
            </xsl:call-template>
        </xsl:attribute>
        
        <!-- List Start Value -->
        <xsl:attribute name="{$list-start-attribute-name}">
            <xsl:call-template name="get-list-start">
                <xsl:with-param name="list-item-level" select="$list-item-level"/>
                <xsl:with-param name="numbering-style" select="$numbering-style"/>
            </xsl:call-template>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Get Format of List -->
    <xsl:template name="get-list-format">
        <xsl:param name="list-item-level"/>
        <xsl:param name="numbering-style"/>
        <xsl:choose>
            <!-- Linked Numbering Style -->
            <xsl:when test="boolean($numbering-style/w:numStyleLink)">
                <xsl:variable name="num-style-link" select="$numbering-style/w:numStyleLink/@w:val"/>
                <xsl:call-template name="get-list-format">
                    <xsl:with-param name="list-item-level" select="$list-item-level"/>
                    <xsl:with-param name="numbering-style" select="$numbering/w:abstractNum[w:styleLink[@w:val=$num-style-link]]"/>
                </xsl:call-template>
            </xsl:when>
            <!-- Custom Numbering Style -->
            <xsl:when test="boolean($numbering-style/w:lvl[@w:ilvl=$list-item-level]/mc:AlternateContent)">
                <xsl:variable name="custom-numbering-style" select="$numbering-style/w:lvl[@w:ilvl=$list-item-level]/mc:AlternateContent/mc:Choice/w:numFmt"/>
                <xsl:variable name="custom-numbering-value" select="$custom-numbering-style/@w:val"/>
                <xsl:variable name="custom-numbering-format" select="$custom-numbering-style/@w:format"/>
                <xsl:value-of select="concat($custom-numbering-value,'; ',$custom-numbering-format)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$numbering-style/w:lvl[@w:ilvl=$list-item-level]/w:numFmt/@w:val"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Get Start Value of List -->
    <xsl:template name="get-list-start">
        <xsl:param name="list-item-level"/>
        <xsl:param name="numbering-style"/>
        <xsl:choose>
            <!-- Linked Numbering Style -->
            <xsl:when test="boolean($numbering-style/w:numStyleLink)">
                <xsl:variable name="num-style-link" select="$numbering-style/w:numStyleLink/@w:val"/>
                <xsl:call-template name="get-list-start">
                    <xsl:with-param name="list-item-level" select="$list-item-level"/>
                    <xsl:with-param name="numbering-style" select="$numbering/w:abstractNum[w:styleLink[@w:val=$num-style-link]]"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$numbering-style/w:lvl[@w:ilvl=$list-item-level]/w:start/@w:val"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>


    <!-- Table -->
    <xsl:template match="w:tbl">
        <xsl:variable name="is-table-ok">
            <xsl:call-template name="check-table"/>
        </xsl:variable>
        <xsl:choose>
            <!-- Tabbed List -->
            <xsl:when test="$is-table-ok = 'false' or $table-mode = 'tabbedlist'">
                <xsl:apply-templates mode="tabbed-list"/>
            </xsl:when>
            <!-- Normal Table -->
            <xsl:otherwise>
                <xsl:element name="{$table-tag-name}" namespace="{$ns}">
                    <xsl:call-template name="insert-table-attributes"/>
                    <xsl:apply-templates/>
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
        <xsl:if test="(following-sibling::w:p or following-sibling::w:tbl) and not(parent::w:footnote or parent::w:endnote)">
            <xsl:text>&#x0d;</xsl:text> <!-- Carriage Return -->
        </xsl:if>
    </xsl:template>
    
    <!-- 
        Check Table
        The number of cells in a row must be equal to the number of columns. 
        (Taking into account the merged cells.)
    -->
    <xsl:template name="check-table">
        <xsl:variable name="num-of-table-columns" select="count(w:tblGrid/w:gridCol)"/>
        <xsl:variable name="faulty-table-rows" select="w:tr[$num-of-table-columns != (count(w:tc) + sum(w:tc/w:tcPr/w:gridSpan/@w:val) - count(w:tc/w:tcPr/w:gridSpan))]"/>
        <xsl:choose>
            <xsl:when test="count($faulty-table-rows) > 0">
                <xsl:value-of select="'false'"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="'true'"/>
            </xsl:otherwise>
        </xsl:choose>
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

    <!-- Table Columns -->
    <xsl:template match="w:gridCol">
        <xsl:apply-templates/>
    </xsl:template>

    <!-- Table Row -->
    <xsl:template match="w:tr">
        <xsl:apply-templates/>
    </xsl:template>
    
    <!-- Table Row - Tabbed List -->
    <xsl:template match="w:tr" mode="tabbed-list">
        <xsl:apply-templates mode="tabbed-list"/>
        <xsl:if test="position() != last()">
            <xsl:text>&#x0d;</xsl:text> <!-- Carriage Return -->
        </xsl:if>
    </xsl:template>

    <!-- Table Cell -->
    <xsl:template match="w:tc">
        <xsl:element name="{$table-cell-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-cell-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

    <!-- Table Cell - Merged -->
    <xsl:template match="w:tc[w:tcPr/w:vMerge[not(@w:val)]]">
        <!-- Skip merged cell --> 
    </xsl:template>
    
    <!-- Table Cell - Tabbed List -->
    <xsl:template match="w:tc" mode="tabbed-list">
        <xsl:apply-templates />
        <xsl:if test="position() != last()">
            <xsl:element name="{$tab-tag-name}" namespace="{$ns}">
                <xsl:call-template name="insert-tabbed-list-tab-attributes"/>
                <xsl:text>&#x9;</xsl:text> <!-- Tab -->
            </xsl:element>
        </xsl:if>
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
        <!-- Column Width -->
        <!-- dxa (twentieths of a point), pct (fiftieths of a percent, not available in InDesign) -->
        <xsl:choose>
            <!-- Defined width -->
            <xsl:when test="w:tcPr/w:tcW/@w:type = 'dxa' or w:tcPr/w:tcW/@w:type = 'pct'">
                <xsl:variable name="column-width" select="number(w:tcPr/w:tcW/@w:w) div 20"/>
                <xsl:attribute name="{concat('aid',':',$table-column-width-attribute-name)}">
                    <xsl:value-of select="$column-width"/>
                </xsl:attribute>
            </xsl:when>
            <!-- Automatically width (Attribute is removed when importing to InDesign but inserted here for consistency.) -->
            <xsl:when test="w:tcPr/w:tcW/@w:type = 'auto'">
                <xsl:attribute name="{$table-column-width-attribute-name}">
                    <xsl:value-of select="'auto'"/>
                </xsl:attribute>
            </xsl:when>
        </xsl:choose>
    </xsl:template>
    
    <!-- Tabbed List Tab Attributes -->
    <xsl:template name="insert-tabbed-list-tab-attributes">
        <xsl:attribute name="{$tabbed-list-tab-type-attribute-name}">
            <xsl:value-of select="$tabbed-list-tab-type-attribute-value"/>
        </xsl:attribute>
        <xsl:attribute name="{concat('aid',':', $inline-style-attribute-name)}">
            <xsl:value-of select="$tabbed-list-tab-inline-style-attribute-name"/>
        </xsl:attribute>
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
                        ($is-inline-style-on-empty-text-removed and w:t and normalize-space(w:t) = '')
                        ">
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="$class-names"/>
                        </xsl:call-template>
                    </xsl:when>
                    <!-- Referenced Character Style -->
                    <xsl:when test="$target-style-element[name() = 'w:rStyle']">
                        <xsl:variable name="character-style-name">
                            <xsl:call-template name="get-style-name">
                                <xsl:with-param name="style-id" select="$target-style-element/@w:val"/>
                                <xsl:with-param name="style-type" select="'character'"/>
                            </xsl:call-template>
                        </xsl:variable>
                        <xsl:variable name="translated-character-style-name" select="translate($character-style-name, '[]', '')"/>
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="concat($translated-character-style-name, ' ', $class-names)"/>
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
                        ($style-mode = 'extended' and 
                            (
                                $target-style-element[name() = 'w:lang'] or
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
            <xsl:apply-templates/>
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


    <!-- Section Break -->
    <xsl:template match="w:sectPr">
        <xsl:element name="{$section-break-tag-name}" namespace="{$ns}">
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

</xsl:transform>