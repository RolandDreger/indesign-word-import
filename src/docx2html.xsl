<?xml version="1.0" encoding="UTF-8"?>

<!--    
        
    Microsoft Word Document -> HTML
    
    Created: September 30, 2021
    Modified: April 8, 2022
    
    Author: Roland Dreger, www.rolanddreger.net
    
    
    # Implemented:
    
    - Paragraph styles (ID, Class)
    - Character styles (Class)
    - Local style overrides (Class)
    - Tables (Header, Merged Cells, Table Styles)
    - Lists (ID, Level for nested Lists)
    - Index (inline)
    - Footnotes (inline)
    - Endnotes (inline)
    - Comments (Notes)
    - Bookmarks (data-content Attribute, empty string for a text mark)
    - Hyperlinks
    - Cross-references
    - Textboxes
    - Images 
    - Track changes
    
    
    # Partially implemented:
    
    - Subdocument Location (only anchor, but not content of the subdocument)
    - Office Math (m:oMath, ... only enclosing element) ToDo: Convert to MathML (standard for HTML5)
    - Textvariablen (Date, Time, ...)
    
    
    # Not implemented:
    
    - Headers: header1.xml, ... (w:hdr with block-level content – w:p, w:tbl, ...)
    - Footers: footer1.xml, ... (w:ftr with block-level content – w:p, w:tbl, ...)
    - IF fields 
    - Calculations
    - Structured Document Tags (w:sdt, ...)
    - Anchor for Imported External Content (w:altChunk)
    - Custom XML (w:customXml, ...)
    - Section Properties (w:sectPr)
    - Table of Authorities (Glosary)
    
    
    
    # Heading Style Map
     
      Default:  
    
      docx Name           HTML Element
     ================= => ==================
      H1 or h1            h1
      Warning-Heading-1   h1 class="warning"
    
      The following parameters define paragraph names that are transformed into H1, H2, ... elements.
      Multiple entries: Names must be enclosed by »«.
      
      
    # Heading Font Size
      
      The headings can also be defined by the font sizes in Microsoft Word document.
      Value 0, if the heading should not be defined by the font size.
      
      e.g.: 
      Font size is 28 + <xsl:param name="h1-font-size" select="28"/>  =>  <h1>-Element
      Font size is 24 + <xsl:param name="h2-font-size" select="24"/>  =>  <h2>-Element
      
      ...
      
      otherwise <p>- or <li>-Element
    
-->

<xsl:transform 
    xmlns=""
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
    exclude-result-prefixes="rd pkg wpc cx cx1 cx2 cx3 cx4 cx5 cx6 cx7 cx8 mc aink am3d o r rel m v wp14 wp w10 w w14 w15 w16cex w16cid w16 w16sdtdh w16se wpg wpi wne wps cp dc dcterms dcmitype dcmitype a pic xsi b"
    version="1.0"
>
    
    
    <!-- ++++++++++++ -->
    <!-- + Settings + -->
    <!-- ++++++++++++ -->
    
    <xsl:param name="ns" select="''"/> <!-- Document Namespace -->
    <xsl:param name="directory-separator" select="'/'"/>
    <xsl:param name="language" select="'en'"/>
    <xsl:param name="max-bookmark-length" select="500"/>
    <xsl:param name="is-empty-paragraph-removed" select="false()"/>
    <xsl:param name="is-inline-style-on-empty-text-removed" select="false()"/>
    <xsl:param name="is-local-override-without-tag-applied" select="false()"/> <!-- Ignore all local overrides except: strong, i, em, u, superscript, subscript  -->
    <xsl:param name="is-comment-inserted" select="false()"/> <!-- Comments for Complex Fields, Tab, ... -->
    <xsl:param name="is-tab-preserved" select="true()"/>  <!-- Tab Character --> 
    
    <!-- Heading Style Map -->
    <xsl:param name="h1-paragraph-style-names" select="''"/> <!-- e.g. '»Custom_Name_1« »Custom_Name_1.1«' -->
    <xsl:param name="h2-paragraph-style-names" select="''"/> <!-- e.g. 'Custom_Name_2' -->
    <xsl:param name="h3-paragraph-style-names" select="''"/> 
    <xsl:param name="h4-paragraph-style-names" select="''"/> 
    <xsl:param name="h5-paragraph-style-names" select="''"/> 
    <xsl:param name="h6-paragraph-style-names" select="''"/> 
    
    <!-- Heading Marker -->
    <xsl:variable name="heading-marker" select="'-Heading-'"/>
    
    <!-- Heading Font Sizes -->
    <xsl:param name="h1-font-size" select="0"/> <!-- e.g. 28 or 28.5 or 0   -->
    <xsl:param name="h2-font-size" select="0"/>
    <xsl:param name="h3-font-size" select="0"/>
    <xsl:param name="h4-font-size" select="0"/>
    <xsl:param name="h5-font-size" select="0"/>
    <xsl:param name="h6-font-size" select="0"/>
    
    <!-- Case conversion -->
    <xsl:variable name="lowercase" select="'abcdefghijklmnopqrstuvwxyzàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿžšœ'" />
    <xsl:variable name="uppercase" select="'ABCDEFGHIJKLMNOPQRSTUVWXYZÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞŸŽŠŒ'" />
    
    
    <!-- +++++++++ -->
    <!-- + INPUT + -->
    <!-- +++++++++ -->
    
    <!-- Folder and File Paths -->
    <xsl:param name="package-base-uri" select="''"/> <!-- an empty string is passed if input is a Word-XML-Document. -->
    <xsl:param name="document-file-name" select="'document.xml'"/> <!-- document.xml or name of Word-XML-Document -->
    <xsl:param name="image-folder-path" select="''"/> <!-- If image folder path is defined, all images get the path according to this pattern: $image-folder-path + '/' + $image-name  -->
    <xsl:param name="core-props-file-path" select="''"/> <!-- docProps/core.xml -->
    <xsl:param name="styles-file-path" select="''"/> <!-- word/styles.xml --> 
    <xsl:param name="numbering-file-path" select="''"/> <!-- word/numbering.xml -->
    <xsl:param name="footnotes-file-path" select="''"/> <!-- word/footnotes.xml --> 
    <xsl:param name="endnotes-file-path" select="''"/> <!-- word/endnotes.xml --> 
    <xsl:param name="comments-file-path" select="''"/> <!-- word/comments.xml --> 
    <xsl:param name="document-relationships-file-path" select="''"/> <!-- word/_rels/document.xml.rels --> 
    <xsl:param name="footnotes-relationships-file-path" select="''"/> <!-- word/_rels/footnotes.xml.rels --> 
    <xsl:param name="endnotes-relationships-file-path" select="''"/> <!-- word/_rels/endnotes.xml.rels --> 
    
    
    <!-- Core Properties (Metadata) -->
    <xsl:variable name="core-props">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($core-props-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $core-props-file-path))/cp:coreProperties/*" />
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/docProps/core.xml']/pkg:xmlData/cp:coreProperties/*"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Styles -->
    <xsl:variable name="styles">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($styles-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $styles-file-path))/w:styles/w:style" />
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/styles.xml']/pkg:xmlData/w:styles/w:style"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Numbering -->
    <xsl:variable name="numbering">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($numbering-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $numbering-file-path))/w:numbering/*" />
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/numbering.xml']/pkg:xmlData/w:numbering/*"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Document Relationships (e.g. Hyperlinks) -->
    <xsl:variable name="document-relationships">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($document-relationships-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $document-relationships-file-path))/rel:Relationships/rel:Relationship"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/_rels/document.xml.rels']/pkg:xmlData/rel:Relationships/rel:Relationship"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Footnotes -->
    <xsl:variable name="footnotes">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($footnotes-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $footnotes-file-path))/w:footnotes/w:footnote" />
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/footnotes.xml']/pkg:xmlData/w:footnotes/w:footnote"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Footnote Relationships -->
    <xsl:variable name="footnotes-relationships">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($footnotes-relationships-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $footnotes-relationships-file-path))/rel:Relationships/rel:Relationship"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/_rels/footnotes.xml.rels']/pkg:xmlData/rel:Relationships/rel:Relationship"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Endnotes -->
    <xsl:variable name="endnotes">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($endnotes-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $endnotes-file-path))/w:endnotes/w:endnote" />
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/endnotes.xml']/pkg:xmlData/w:endnotes/w:endnote"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Endnote Relationships -->
    <xsl:variable name="endnotes-relationships">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($endnotes-relationships-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $endnotes-relationships-file-path))/rel:Relationships/rel:Relationship"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/_rels/endnotes.xml.rels']/pkg:xmlData/rel:Relationships/rel:Relationship"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Comments -->
    <xsl:variable name="comments">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($comments-file-path)">
                <xsl:copy-of select="document(concat($package-base-uri, $directory-separator, $comments-file-path))/w:comments/w:comment" />
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[@pkg:name = '/word/comments.xml']/pkg:xmlData/w:comments/w:comment"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    <!-- Citations -->
    <xsl:variable name="citations-relationships">
        <xsl:choose>
            <xsl:when test="boolean($package-base-uri) and boolean($document-relationships-file-path)">
                <xsl:variable name="relationships-document" select="document(concat($package-base-uri, $directory-separator, $document-relationships-file-path))"/>
                <xsl:for-each select="$relationships-document/rel:Relationships/rel:Relationship">
                    <xsl:if test="contains(@Type, 'customXml')">
                        <xsl:variable name="custom-xml-file-path" select="concat($package-base-uri, $directory-separator, substring-after(@Target, '../'))"/>
                        <xsl:variable name="sources-element" select="document($custom-xml-file-path)/b:Sources"/>
                        <xsl:if test="$sources-element and namespace-uri($sources-element) = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'">
                            <xsl:copy-of select="document($custom-xml-file-path)/b:Sources"/>
                        </xsl:if>
                    </xsl:if>
                </xsl:for-each>  
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="/pkg:package/pkg:part[contains(@pkg:name, 'customXml')]/pkg:xmlData/b:Sources[namespace-uri() = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography']"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:variable>
    
    
    
    <!-- ++++++++++ -->
    <!-- + OUTPUT + -->
    <!-- ++++++++++ -->
    
    <xsl:output method="html" version="5.0" media-type="text/html" indent="yes"/>
    
    <!-- Output tag and attribute names -->
    <xsl:variable name="heading-tag-name" select="'h'"/>
    <xsl:variable name="paragraph-tag-name" select="'p'"/>
    <xsl:variable name="paragraph-id-attribute-name" select="'id'"/>
    <xsl:variable name="paragraph-style-attribute-name" select="'class'"/>
    <xsl:variable name="list-item-tag-name" select="'li'"/>
    <xsl:variable name="list-item-level-attribute-name" select="'data-item-level'"/>
    <xsl:variable name="list-id-attribute-name" select="'data-list-id'"/>
    <xsl:variable name="table-tag-name" select="'table'"/>
    <xsl:variable name="table-style-attribute-name" select="'class'"/>
    <xsl:variable name="table-column-group-tag-name" select="'colgroup'"/>
    <xsl:variable name="table-column-tag-name" select="'col'"/>
    <xsl:variable name="table-row-tag-name" select="'tr'"/>
    <xsl:variable name="table-cell-tag-name" select="'td'"/>
    <xsl:variable name="table-cell-column-span-attribute-name" select="'colspan'"/>
    <xsl:variable name="table-cell-row-span-attribute-name" select="'rowspan'"/>
    <xsl:variable name="table-header-cell-tag-name" select="'th'"/>
    <xsl:variable name="table-header-cell-scope-attribute-name" select="'scope'"/>
    <xsl:variable name="footnote-tag-name" select="'foot-note'"/>
    <xsl:variable name="footnote-reference-tag-name" select="'foot-note-ref'"/>
    <xsl:variable name="footnote-index-attribute-name" select="'index'"/>
    <xsl:variable name="footnote-style-attribute-name" select="'class'"/>
    <xsl:variable name="footnote-style-attribute-value" select="'footnote'"/>
    <xsl:variable name="endnote-tag-name" select="'end-note'"/>
    <xsl:variable name="endnote-reference-tag-name" select="'end-note-ref'"/>
    <xsl:variable name="endnote-index-attribute-name" select="'index'"/>
    <xsl:variable name="endnote-style-attribute-name" select="'class'"/>
    <xsl:variable name="endnote-style-attribute-value" select="'endnote'"/>
    <xsl:variable name="comment-tag-name" select="'side-note'"/>
    <xsl:variable name="comment-reference-tag-name" select="'side-note-ref'"/>
    <xsl:variable name="comment-style-attribute-name" select="'class'"/>
    <xsl:variable name="comment-style-attribute-value" select="'comment'"/>
    <xsl:variable name="comment-index-attribute-name" select="'index'"/>
    <xsl:variable name="comment-date-attribute-name" select="'data-date'"/>
    <xsl:variable name="comment-initials-attribute-name" select="'data-initials'"/>
    <xsl:variable name="comment-author-attribute-name" select="'data-author'"/>
    <xsl:variable name="citation-tag-name" select="'citation-note'"/>
    <xsl:variable name="citation-call-tag-name" select="'citation-call'"/>
    <xsl:variable name="citation-source-tag-name" select="'citation-source'"/>
    <xsl:variable name="citation-style-attribute-name" select="'class'"/>
    <xsl:variable name="citation-style-attribute-value" select="'citation'"/>
    <xsl:variable name="citation-style-type-attribute-name" select="'data-style-type'"/>
    <xsl:variable name="citation-style-name-attribute-name" select="'data-style-name'"/>
    <xsl:variable name="citation-version-attribute-name" select="'data-version'"/>
    <xsl:variable name="citation-text-tag-name" select="'text'"/>
    <xsl:variable name="citation-value-attribute-name" select="'value'"/>
    <xsl:variable name="group-tag-name" select="'shape-group'"/>
    <xsl:variable name="group-style-attribute-name" select="'class'"/>
    <xsl:variable name="group-style-attribute-value" select="'shapegroup'"/>
    <xsl:variable name="image-tag-name" select="'img'"/>
    <xsl:variable name="image-style-attribute-name" select="'class'"/>
    <xsl:variable name="image-style-anchored-attribute-value" select="'anchored-image'"/>
    <xsl:variable name="image-style-inline-attribute-value" select="'inline-image'"/>
    <xsl:variable name="image-source-attribute-name" select="'src'"/>
    <xsl:variable name="image-title-attribute-name" select="'title'"/>
    <xsl:variable name="image-alt-attribute-name" select="'alt'"/>
    <xsl:variable name="image-position-attribute-name" select="'data-position'"/>
    <xsl:variable name="image-uri-attribute-name" select="'data-uri'"/>
    <xsl:variable name="textbox-tag-name" select="'text-box'"/>
    <xsl:variable name="textbox-index-attribute-name" select="'index'"/>
    <xsl:variable name="textbox-style-attribute-name" select="'class'"/>
    <xsl:variable name="textbox-style-attribute-value" select="'textbox'"/>
    <xsl:variable name="textbox-name-attribute-name" select="'name'"/>
    <xsl:variable name="textbox-alt-attribute-name" select="'alt'"/>
    <xsl:variable name="shape-tag-name" select="'vector-shape'"/>
    <xsl:variable name="shape-index-attribute-name" select="'index'"/>
    <xsl:variable name="shape-style-attribute-name" select="'class'"/>
    <xsl:variable name="shape-style-attribute-value" select="'vectorshape'"/>
    <xsl:variable name="shape-name-attribute-name" select="'name'"/>
    <xsl:variable name="shape-alt-attribute-name" select="'alt'"/>
    <xsl:variable name="embedded-object-tag-name" select="'embedded-object'"/>
    <xsl:variable name="embedded-object-style-attribute-name" select="'class'"/>
    <xsl:variable name="embedded-object-style-attribute-value" select="'embedded-object'"/>
    <xsl:variable name="embedded-object-target-attribute-name" select="'data-target'"/>
    <xsl:variable name="embedded-object-program-attribute-name" select="'data-program'"/>
    <xsl:variable name="bookmark-tag-name" select="'book-mark'"/>
    <xsl:variable name="bookmark-index-attribute-name" select="'index'"/>
    <xsl:variable name="bookmark-style-attribute-name" select="'class'"/>
    <xsl:variable name="bookmark-style-attribute-value" select="'bookmark'"/>
    <xsl:variable name="bookmark-id-attribute-name" select="'id'"/>
    <xsl:variable name="bookmark-content-attribute-name" select="'data-content'"/>
    <xsl:variable name="indexmark-tag-name" select="'index-mark'"/>
    <xsl:variable name="indexmark-style-attribute-name" select="'class'"/>
    <xsl:variable name="indexmark-style-attribute-value" select="'indexmark'"/>
    <xsl:variable name="indexmark-type-attribute-name" select="'data-type'"/>
    <xsl:variable name="indexmark-format-attribute-name" select="'data-format'"/>
    <xsl:variable name="indexmark-entry-attribute-name" select="'data-entry'"/>
    <xsl:variable name="indexmark-target-attribute-name" select="'data-target'"/>
    <xsl:variable name="complex-field-tag-name" select="'complex-field'"/>
    <xsl:variable name="complex-field-style-attribute-name" select="'class'"/>
    <xsl:variable name="complex-field-style-attribute-value" select="'complex-field'"/>
    <xsl:variable name="complex-field-content-attribute-name" select="'data-content'"/>
    <xsl:variable name="complex-field-data-attribute-name" select="'data-data'"/>
    <xsl:variable name="hyperlink-tag-name" select="'a'"/>
    <xsl:variable name="hyperlink-uri-attribute-name" select="'href'"/>
    <xsl:variable name="hyperlink-title-attribute-name" select="'title'"/>
    <xsl:variable name="cross-reference-tag-name" select="'a'"/>
    <xsl:variable name="cross-reference-uri-attribute-name" select="'href'"/>
    <xsl:variable name="cross-reference-type-attribute-name" select="'data-type'"/>
    <xsl:variable name="cross-reference-format-attribute-name" select="'data-format'"/>
    <xsl:variable name="inline-style-tag-name" select="'span'"/>
    <xsl:variable name="inline-style-attribute-name" select="'class'"/>
    <xsl:variable name="bold-tag-name" select="'strong'"/>
    <xsl:variable name="italics-tag-name" select="'i'"/>
    <xsl:variable name="emphasis-mark-tag-name" select="'em'"/>
    <xsl:variable name="underline-tag-name" select="'u'"/>
    <xsl:variable name="subscript-tag-name" select="'sub'"/>
    <xsl:variable name="superscript-tag-name" select="'sup'"/>
    <xsl:variable name="section-break-tag-name" select="'span'"/>
    <xsl:variable name="section-break-type-attribute-name" select="'data-wrap-type'"/>
    <xsl:variable name="page-break-tag-name" select="'br'"/>
    <xsl:variable name="column-break-tag-name" select="'br'"/>
    <xsl:variable name="forced-line-break-tag-name" select="'br'"/>
    <xsl:variable name="carriage-return-tag-name" select="'br'"/>
    <xsl:variable name="inserted-text-tag-name" select="'ins'"/>
    <xsl:variable name="deleted-text-tag-name" select="'del'"/>
    <xsl:variable name="moved-to-text-tag-name" select="'ins'"/>
    <xsl:variable name="moved-from-text-tag-name" select="'del'"/>
    <xsl:variable name="track-change-author-attribute-name" select="'data-author'"/>
    <xsl:variable name="track-change-date-attribute-name" select="'datetime'"/>
    <xsl:variable name="subdocument-tag-name" select="'sub-document'"/>
    <xsl:variable name="subdocument-style-attribute-name" select="'class'"/>
    <xsl:variable name="subdocument-style-attribute-value" select="'sub-document'"/>
    <xsl:variable name="subdocument-uri-attribute-name" select="'data-uri'"/>
    <xsl:variable name="equation-tag-name" select="'math-equation'"/>
    <xsl:variable name="equation-style-attribute-name" select="'class'"/>
    <xsl:variable name="equation-style-attribute-value" select="'math-equation'"/>
    <xsl:variable name="time-tag-name" select="'time'"/>
    <xsl:variable name="time-format-attribute-name" select="'data-format'"/>
    <xsl:variable name="time-type-attribute-name" select="'data-type'"/>
    <xsl:variable name="data-tag-name" select="'data'"/>
    <xsl:variable name="data-value-attribute-name" select="'value'"/>
    <xsl:variable name="symbol-tag-name" select="'symbol-character'"/>
    <xsl:variable name="symbol-style-attribute-name" select="'class'"/>
    <xsl:variable name="symbol-style-attribute-value" select="'symbol-character'"/>
    <xsl:variable name="symbol-font-attribute-name" select="'data-font'"/>
    <xsl:variable name="symbol-code-attribute-name" select="'data-unicode'"/>
    <xsl:variable name="tab-tag-name" select="'span'"/>
    <xsl:variable name="tab-style-attribute-name" select="'class'"/>
    <xsl:variable name="tab-style-attribute-value" select="'tab'"/>
    
    
    <!-- Spaces -->
    <xsl:preserve-space elements="w:t"/>
    <xsl:strip-space elements="pkg:package pkg:part pkg:xmlData w:document w:body w:p w:pPr w:rPr w:r w:sectPr"/>
    


    <!-- +++++++++++++ -->
    <!-- + Templates + -->
    <!-- +++++++++++++ -->
    
    <xsl:template match="/">
        <xsl:element name="html" namespace="{$ns}">
            <xsl:attribute name="xml:lang">
                <xsl:value-of select="$language"/>
            </xsl:attribute>
            <xsl:attribute name="lang">
                <xsl:value-of select="$language"/>
            </xsl:attribute>
            <!-- Head -->
            <xsl:call-template name="create-head-section" />
            <!-- Body -->
            <xsl:call-template name="create-body-section" />
        </xsl:element>
    </xsl:template>
    
    
    <!-- Head -->
    <xsl:template name="create-head-section">
        <xsl:element name="head" namespace="{$ns}">
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="charset">UTF-8</xsl:attribute>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="name">
                    <xsl:value-of select="'viewport'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="'width=device-width, initial-scale=1'"/>
                </xsl:attribute>
            </xsl:element>
            <xsl:element name="title" namespace="{$ns}">
                <xsl:value-of select="$core-props/dc:title"/>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="name">
                    <xsl:value-of select="'author'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="$core-props/dc:creator"/>
                </xsl:attribute>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="name">
                    <xsl:value-of select="'keywords'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="$core-props/cp:keywords"/>
                </xsl:attribute>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="name">
                    <xsl:value-of select="'description'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="$core-props/dc:description"/>
                </xsl:attribute>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="name">
                    <xsl:value-of select="'generator'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="'Stylesheet'"/>
                </xsl:attribute>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="property">
                    <xsl:value-of select="'created'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="$core-props/dcterms:created"/>
                </xsl:attribute>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="property">
                    <xsl:value-of select="'modified'"/>
                </xsl:attribute>
                <xsl:attribute name="content">
                    <xsl:value-of select="$core-props/dcterms:modified"/>
                </xsl:attribute>
            </xsl:element>
        </xsl:element>
    </xsl:template>
    
    
    <!-- Body -->
    <xsl:template name="create-body-section">
        <xsl:element name="body" namespace="{$ns}">
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <xsl:template match="pkg:package">
        <xsl:apply-templates select="pkg:part[@pkg:name='/word/document.xml']/pkg:xmlData/w:document" /> <!-- for Word XML Document (single XML file) -->
    </xsl:template>

    <xsl:template match="w:document">
        <xsl:apply-templates />
    </xsl:template>
    
    <xsl:template match="w:body">
        <xsl:call-template name="structure-paragraphs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    
    
    
    <!-- Structure paragraphs (e.g. lists) -->
    <xsl:template name="structure-paragraphs">
        <xsl:param name="target-elements"/>
        <xsl:choose>
            <!-- Check: Are there any (list) elements? -->
            <xsl:when test="not($target-elements/w:pPr/w:numPr)">
                <xsl:apply-templates select="$target-elements"/>
            </xsl:when>
            <!-- Check: List elements only? -->
            <xsl:when test="not($target-elements[not(w:pPr/w:numPr)])">
                <xsl:call-template name="create-html-list">
                    <xsl:with-param name="list-elements" select="$target-elements"/>
                    <xsl:with-param name="prev-level" select="-1"/>
                </xsl:call-template>
            </xsl:when>
            <!-- Check: First element is list element? -->
            <xsl:when test="$target-elements[1]/w:pPr/w:numPr">
                <xsl:variable name="num-of-following-elements" select="count($target-elements[not(w:pPr/w:numPr)][1]/following-sibling::*) + 1"/>
                <xsl:variable name="num-of-list-elements" select="count($target-elements) - $num-of-following-elements"/>
                <xsl:call-template name="create-html-list">
                    <xsl:with-param name="list-elements" select="$target-elements[position() &lt;= $num-of-list-elements]"/>
                    <xsl:with-param name="prev-level" select="-1"/>
                </xsl:call-template>
                <xsl:call-template name="structure-paragraphs">
                    <xsl:with-param name="target-elements" select="$target-elements[position() &gt; $num-of-list-elements]"/>
                </xsl:call-template>
            </xsl:when>
            <!-- The first element is not a list element, but there are! -->
            <xsl:otherwise>
                <xsl:variable name="num-of-following-elements" select="count($target-elements[w:pPr/w:numPr][1]/following-sibling::*) + 1"/>
                <xsl:variable name="num-of-non-list-elements" select="count($target-elements) - $num-of-following-elements"/>
                <xsl:apply-templates select="$target-elements[position() &lt;= $num-of-non-list-elements]"/>
                <xsl:call-template name="structure-paragraphs">
                    <xsl:with-param name="target-elements" select="$target-elements[position() &gt; $num-of-non-list-elements]"/>
                </xsl:call-template>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- HTML list -->
    <!-- 
        Level    (@w:ilvl = 0)
        . Level  (@w:ilvl = 1)
        .. Level (@w:ilvl = 2)
        ...
    -->
    <xsl:template name="create-html-list">
        <xsl:param name="list-elements"/>
        <xsl:param name="prev-level"/>
        
        <xsl:variable name="first-element" select="$list-elements[1]"/>
        <xsl:variable name="cur-level" select="$first-element/w:pPr/w:numPr/w:ilvl/@w:val"/>
        
        <xsl:variable name="first-lower-level-element" select="$list-elements[w:pPr/w:numPr/w:ilvl/@w:val &gt; $cur-level][1]"/>
        <xsl:variable name="lower-level" select="$first-lower-level-element/w:pPr/w:numPr/w:ilvl/@w:val"/>
        
        <xsl:variable name="first-diff-level-element" select="$list-elements[w:pPr/w:numPr/w:ilvl/@w:val != $cur-level][1]"/>
        <xsl:variable name="first-following-cur-level-element" select="$first-lower-level-element/following-sibling::*[w:pPr/w:numPr/w:ilvl/@w:val = $cur-level][1]"/>
        
        <xsl:variable name="num-of-first-cur-level-elements" select="count($first-diff-level-element/preceding-sibling::*) - count($first-element/preceding-sibling::*)"/>
        <xsl:variable name="num-of-first-lower-level-elements" select="count($first-following-cur-level-element/preceding-sibling::*) - count($first-lower-level-element/preceding-sibling::*)"/>
        
        <!-- Element name for current level -->
        <xsl:variable name="first-num-id" select="$first-element/w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="first-abstract-num-id" select="$numbering/w:num[@w:numId = $first-num-id]/w:abstractNumId/@w:val"/>
        <xsl:variable name="first-num-format" select="$numbering/w:abstractNum[@w:abstractNumId = $first-abstract-num-id]/w:lvl[@w:ilvl = $cur-level]/w:numFmt/@w:val"/>
        <xsl:variable name="cur-level-list-tag-name">
            <xsl:call-template name="get-list-tag-name">
                <xsl:with-param name="num-format" select="$first-num-format"/>
            </xsl:call-template>
        </xsl:variable>
        
        <!-- Element name for lower level -->
        <xsl:variable name="first-lower-level-num-id" select="$first-lower-level-element/w:pPr/w:numPr/w:numId/@w:val"/>
        <xsl:variable name="first-lower-level-abstract-num-id" select="$numbering/w:num[@w:numId = $first-lower-level-num-id]/w:abstractNumId/@w:val"/>
        <xsl:variable name="first-lower-level-num-format" select="$numbering/w:abstractNum[@w:abstractNumId = $first-lower-level-abstract-num-id]/w:lvl[@w:ilvl = $lower-level]/w:numFmt/@w:val"/>
        <xsl:variable name="lower-level-list-tag-name">
            <xsl:call-template name="get-list-tag-name">
                <xsl:with-param name="num-format" select="$first-lower-level-num-format"/>
            </xsl:call-template>
        </xsl:variable>
        
        <xsl:choose>
            <!-- Check: Any different levels at first run? -->
            <xsl:when test="not($first-diff-level-element) and $prev-level = -1">
                <xsl:element name="{$cur-level-list-tag-name}" namespace="{$ns}">
                    <!-- List Attributes -->
                    <xsl:call-template name="insert-list-attributes">
                        <xsl:with-param name="id" select="$first-num-id"/>
                    </xsl:call-template>
                    <!-- List Items -->
                    <xsl:for-each select="$list-elements">
                        <xsl:element name="{$list-item-tag-name}" namespace="{$ns}">
                            <!-- List Item Attributes -->
                            <xsl:call-template name="insert-list-item-attributes">
                                <xsl:with-param name="level" select="$cur-level"/>
                            </xsl:call-template>
                            <!-- Liste Item Content -->
                            <xsl:apply-templates select="."/> <!-- omit paragraph tag: mode="content-only" -->
                        </xsl:element>
                    </xsl:for-each>
                </xsl:element>
            </xsl:when>
            <!-- Check: First run with different levels?  -->
            <xsl:when test="$prev-level = -1">
                <xsl:element name="{$cur-level-list-tag-name}" namespace="{$ns}">
                    <!-- List Attributes -->
                    <xsl:call-template name="insert-list-attributes">
                        <xsl:with-param name="id" select="$first-num-id"/>
                    </xsl:call-template>
                    <!-- List Content -->
                    <xsl:call-template name="create-html-list">
                        <xsl:with-param name="list-elements" select="$list-elements"/>
                        <xsl:with-param name="prev-level" select="$cur-level"/>
                    </xsl:call-template>
                </xsl:element>
            </xsl:when>
            <!-- Check: Any lower level elements with following current level elements? -->
            <xsl:when test="$first-lower-level-element and $first-following-cur-level-element">
                <!-- List Items -->
                <xsl:for-each select="$list-elements[position() &lt; $num-of-first-cur-level-elements]">
                    <xsl:element name="{$list-item-tag-name}" namespace="{$ns}">
                        <!-- List Item Attributes -->
                        <xsl:call-template name="insert-list-item-attributes">
                            <xsl:with-param name="level" select="$cur-level"/>
                        </xsl:call-template>
                        <!-- Liste Item Content -->
                        <xsl:apply-templates select="."/> <!-- omit paragraph tag: mode="content-only" -->
                    </xsl:element>
                </xsl:for-each>
                <!-- List Item -->
                <xsl:element name="{$list-item-tag-name}" namespace="{$ns}">
                    <!-- List Item Attributes -->
                    <xsl:call-template name="insert-list-item-attributes">
                        <xsl:with-param name="level" select="$cur-level"/>
                    </xsl:call-template>
                    <!-- List Item Content -->
                    <xsl:apply-templates select="$list-elements[position() = $num-of-first-cur-level-elements]"/> <!-- omit paragraph tag: mode="content-only" -->
                    <!-- Nested List Element (lower Level) -->
                    <xsl:element name="{$lower-level-list-tag-name}" namespace="{$ns}">
                        <!-- List Attributes -->
                        <xsl:call-template name="insert-list-attributes">
                            <xsl:with-param name="id" select="$first-lower-level-num-id"/>
                        </xsl:call-template>
                        <!-- List Content -->
                        <xsl:call-template name="create-html-list">
                            <xsl:with-param name="list-elements" select="$list-elements[position() &gt; $num-of-first-cur-level-elements and position() &lt;= ($num-of-first-cur-level-elements + $num-of-first-lower-level-elements)]"/>
                            <xsl:with-param name="prev-level" select="$cur-level"/>
                        </xsl:call-template>
                    </xsl:element>
                </xsl:element>
                <!-- Remaining Elements (same Level) -->
                <xsl:call-template name="create-html-list">
                    <xsl:with-param name="list-elements" select="$list-elements[position() &gt; ($num-of-first-cur-level-elements + $num-of-first-lower-level-elements)]"/>
                    <xsl:with-param name="prev-level" select="$cur-level"/>
                </xsl:call-template>
            </xsl:when>
            <!-- Check: Any lower level elements with NO following current level elements? -->
            <xsl:when test="$first-lower-level-element and not($first-following-cur-level-element)">
                <!-- List Items -->
                <xsl:for-each select="$list-elements[position() &lt; $num-of-first-cur-level-elements]">
                    <xsl:element name="{$list-item-tag-name}" namespace="{$ns}">
                        <!-- List Item Attributes -->
                        <xsl:call-template name="insert-list-item-attributes">
                            <xsl:with-param name="level" select="$cur-level"/>
                        </xsl:call-template>
                        <!-- List Item Content -->
                        <xsl:apply-templates select="."/> <!-- omit paragraph tag: mode="content-only" -->
                    </xsl:element>
                </xsl:for-each>
                <!-- List Item -->
                <xsl:element name="{$list-item-tag-name}" namespace="{$ns}">
                    <!-- List Item Attributes -->
                    <xsl:call-template name="insert-list-item-attributes">
                        <xsl:with-param name="level" select="$cur-level"/>
                    </xsl:call-template>
                    <!-- List Item Content -->
                    <xsl:apply-templates select="$list-elements[position() = $num-of-first-cur-level-elements]"/> <!-- omit paragraph tag: mode="content-only" -->
                    <!-- Nested List Element (lower Level) -->
                    <xsl:element name="{$lower-level-list-tag-name}" namespace="{$ns}">
                        <!-- List Attributes -->
                        <xsl:call-template name="insert-list-attributes">
                            <xsl:with-param name="id" select="$first-lower-level-num-id"/>
                        </xsl:call-template>
                        <!-- List Content -->
                        <xsl:call-template name="create-html-list">
                            <xsl:with-param name="list-elements" select="$list-elements[position() &gt; $num-of-first-cur-level-elements]"/>
                            <xsl:with-param name="prev-level" select="$cur-level"/>
                        </xsl:call-template>
                    </xsl:element>
                </xsl:element>
            </xsl:when>
            <!-- The boring rest -->
            <xsl:otherwise>
                <!-- List Items -->
                <xsl:for-each select="$list-elements">
                    <xsl:element name="{$list-item-tag-name}" namespace="{$ns}">
                        <!-- List Item Attributes -->
                        <xsl:call-template name="insert-list-item-attributes">
                            <xsl:with-param name="level" select="$cur-level"/>
                        </xsl:call-template>
                        <!-- List Item Content -->
                        <xsl:apply-templates select="."/> <!-- omit paragraph tag: mode="content-only" -->
                    </xsl:element>
                </xsl:for-each>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Tag name for List Element -->
    <xsl:template name="get-list-tag-name">
        <xsl:param name="num-format" select="''"/>
        <xsl:choose>
            <xsl:when test="not($num-format) or $num-format = 'bullet' or $num-format ='none'">
                <xsl:text>ul</xsl:text>
            </xsl:when>
            <xsl:otherwise>
                <xsl:text>ol</xsl:text>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Attributes for List Element (ol, ul) -->
    <xsl:template name="insert-list-attributes">
        <xsl:param name="id" select="''"/>
        <!-- ID -->
        <xsl:attribute name="{$list-id-attribute-name}">
            <xsl:value-of select="$id"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Attribute for List Item Element -->
    <xsl:template name="insert-list-item-attributes">
        <xsl:param name="level" select="''"/>
        <!-- Level -->
        <xsl:attribute name="{$list-item-level-attribute-name}">
            <xsl:value-of select="$level"/>
        </xsl:attribute>
    </xsl:template>
    
    
    <!-- Paragraph -->
    <xsl:template match="w:p">
        <xsl:choose>
            <!-- Check: Remove empty paragraphs? -->
            <xsl:when test="$is-empty-paragraph-removed and normalize-space(.) = '' and count(*) = 0">
                <!-- Skip paragraph -->
            </xsl:when>
            <xsl:otherwise>
                <xsl:variable name="paragraph-tag-name">
                    <xsl:call-template name="get-paragraph-tag-name"/>
                </xsl:variable>
                <xsl:element name="{$paragraph-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-paragraph-attributes">
                        <xsl:with-param name="tag-name" select="$paragraph-tag-name"/>
                    </xsl:call-template>
                    <!-- Structure text runs (e.g. complex fields) -->
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="*"/>
                    </xsl:call-template>
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Paragraph content (without container element) -->
    <!--
    <xsl:template match="w:p" mode="content-only">
        <xsl:call-template name="structure-text-runs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    -->
    
    <!-- Tag Name for Paragraph (h1, h2, ..., li, p) -->
    <xsl:template name="get-paragraph-tag-name">
        <xsl:param name="target-element" select="."/>
        <xsl:variable name="p-style-id" select="$target-element/w:pPr/w:pStyle/@w:val"/>
        <xsl:variable name="p-style-name" select="$styles/w:style[@w:type='paragraph' and @w:styleId=$p-style-id]/w:name/@w:val"/>
        <xsl:variable name="heading-level">
            <xsl:choose>
                <!-- Heading: Default WordML Name -->
                <xsl:when test="starts-with($p-style-name, 'heading')">
                    <xsl:variable name="heading-level-string" select="substring-after($p-style-name, 'heading')"/>
                    <xsl:variable name="heading-level-number" select="number($heading-level-string)"/>
                    <xsl:choose>
                        <xsl:when test="$heading-level-number">
                            <xsl:value-of select="$heading-level-number"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="0"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>
                <!-- Heading: User Paragraph Name (H1, H2, H3, H4, H5, H6) -->
                <xsl:when test="
                    translate($p-style-name, 'H ', 'h') = 'h1' or 
                    translate($p-style-name, 'H ', 'h') = 'h2' or
                    translate($p-style-name, 'H ', 'h') = 'h3' or
                    translate($p-style-name, 'H ', 'h') = 'h4' or
                    translate($p-style-name, 'H ', 'h') = 'h5' or
                    translate($p-style-name, 'H ', 'h') = 'h6'
                ">
                    <xsl:value-of select="number(substring-after(translate($p-style-name, 'H','h'), 'h'))"/>
                </xsl:when>
                <!-- Heading: User Paragraph Name with $heading-marker -->
                <xsl:when test="contains($p-style-name, $heading-marker)">
                    <xsl:value-of select="number(substring-after(translate($p-style-name, 'h','H'), $heading-marker))"/>
                </xsl:when>
                <!-- Heading: User Paragraph Name in $h1-paragraph-style-names -->
                <xsl:when test="boolean($h1-paragraph-style-names)">
                    <xsl:choose>
                        <xsl:when test="($h1-paragraph-style-names = $p-style-name or (contains($h1-paragraph-style-names, concat('»', $p-style-name, '«'))))">
                            <xsl:value-of select="1"/>
                        </xsl:when>
                        <xsl:when test="($h2-paragraph-style-names = $p-style-name or (contains($h2-paragraph-style-names, concat('»', $p-style-name, '«'))))">
                            <xsl:value-of select="2"/>
                        </xsl:when>
                        <xsl:when test="($h3-paragraph-style-names = $p-style-name or (contains($h3-paragraph-style-names, concat('»', $p-style-name, '«'))))">
                            <xsl:value-of select="3"/>
                        </xsl:when>
                        <xsl:when test="($h4-paragraph-style-names = $p-style-name or (contains($h4-paragraph-style-names, concat('»', $p-style-name, '«'))))">
                            <xsl:value-of select="4"/>
                        </xsl:when>
                        <xsl:when test="($h5-paragraph-style-names = $p-style-name or (contains($h5-paragraph-style-names, concat('»', $p-style-name, '«'))))">
                            <xsl:value-of select="5"/>
                        </xsl:when>
                        <xsl:when test="($h6-paragraph-style-names = $p-style-name or (contains($h6-paragraph-style-names, concat('»', $p-style-name, '«'))))">
                            <xsl:value-of select="6"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="0"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>
                <!-- Heading: Defined Font Size (last in chain) -->
                <xsl:when test="$h1-font-size or $h2-font-size or $h3-font-size or $h4-font-size or $h5-font-size or $h6-font-size">
                    <xsl:variable name="sz-attribute-value">
                        <xsl:choose>
                            <xsl:when test="$target-element/w:pPr/w:rPr/w:sz">
                                <xsl:value-of select="$target-element/w:pPr/w:rPr/w:sz/@w:val"/>
                            </xsl:when>
                            <xsl:when test="$target-element/w:pPr/w:szCs">
                                <xsl:value-of select="$target-element/w:pPr/w:rPr/w:szCs/@w:val"/>
                            </xsl:when>
                           <xsl:when test="$styles/w:style[@w:type='paragraph' and @w:styleId=$p-style-id]/w:rPr/w:sz">
                                <xsl:value-of select="$styles/w:style[@w:type='paragraph' and @w:styleId=$p-style-id]/w:rPr/w:sz/@w:val"/>
                            </xsl:when>
                            <xsl:when test="$styles/w:style[@w:type='paragraph' and @w:styleId=$p-style-id]/w:rPr/w:szCs">
                                <xsl:value-of select="$styles/w:style[@w:type='paragraph' and @w:styleId=$p-style-id]/w:rPr/w:szCs/@w:val"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="0"/>
                            </xsl:otherwise>
                        </xsl:choose>   
                    </xsl:variable>
                    <xsl:variable name="p-font-size" select="number($sz-attribute-value) div 2"/>
                    <xsl:choose>
                        <xsl:when test="$p-font-size and $p-font-size = $h1-font-size">
                            <xsl:value-of select="1"/>
                        </xsl:when>
                        <xsl:when test="$p-font-size and $p-font-size = $h2-font-size">
                            <xsl:value-of select="2"/>
                        </xsl:when>
                        <xsl:when test="$p-font-size and $p-font-size = $h3-font-size">
                            <xsl:value-of select="3"/>
                        </xsl:when>
                        <xsl:when test="$p-font-size and $p-font-size = $h4-font-size">
                            <xsl:value-of select="4"/>
                        </xsl:when>
                        <xsl:when test="$p-font-size and $p-font-size = $h5-font-size">
                            <xsl:value-of select="5"/>
                        </xsl:when>
                        <xsl:when test="$p-font-size and $p-font-size = $h6-font-size">
                            <xsl:value-of select="6"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="0"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="0"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:choose>
            <!-- Heading -->
            <xsl:when test="boolean($heading-level) and $heading-level &gt; 0 and $heading-level &lt;= 6">
                <xsl:value-of select="concat($heading-tag-name,$heading-level)"/>
            </xsl:when>
            <!-- Normal Paragraph -->
            <xsl:otherwise>
                <xsl:value-of select="$paragraph-tag-name"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Attribute for Headline and Paragraph -->
    <xsl:template name="insert-paragraph-attributes">
        <xsl:param name="tag-name" select="'p'"/>
        <!-- ID -->
        <xsl:attribute name="{$paragraph-id-attribute-name}">
            <xsl:choose>
                <xsl:when test="@w14:paraId">
                    <xsl:value-of select="@w14:paraId"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="generate-id()"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
        <!-- Paragraph Style -->
        <xsl:if test="w:pPr/w:pStyle">
            <xsl:variable name="p-style-id" select="w:pPr/w:pStyle/@w:val"/>
            <xsl:attribute name="{$paragraph-style-attribute-name}">
                <xsl:variable name="class">
                    <xsl:choose>
                        <xsl:when test="contains($p-style-id, $heading-marker)">
                            <xsl:value-of select="substring-before($p-style-id, $heading-marker)"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="$p-style-id"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:variable>
                <xsl:value-of select="translate(translate($class, $uppercase, $lowercase), ' ', '-')"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Table -->
    <xsl:template match="w:tbl">
        <xsl:element name="{$table-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-table-attributes"/>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Table -->
    <xsl:template name="insert-table-attributes">
        <xsl:attribute name="{$table-style-attribute-name}">
            <xsl:value-of select="w:tblPr/w:tblStyle/@w:val"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Table Column Group -->
    <xsl:template match="w:tblGrid">
        <xsl:element name="{$table-column-group-tag-name}" namespace="{$ns}">
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Table Column -->
    <xsl:template match="w:gridCol">
        <xsl:element name="{$table-column-tag-name}" namespace="{$ns}">
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Table Row -->
    <xsl:template match="w:tr">
        <xsl:element name="{$table-row-tag-name}" namespace="{$ns}">
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Table Cell -->
    <xsl:template match="w:tc">
        <xsl:variable name="cell-tag-name">
            <xsl:call-template name="get-cell-tag-name"/>
        </xsl:variable>
        <xsl:element name="{$cell-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-cell-attributes"/>
            <xsl:call-template name="structure-paragraphs">
                <xsl:with-param name="target-elements" select="*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Merged Table Cell -->
    <xsl:template match="w:tc[w:tcPr/w:vMerge[not(@w:val)]]">
        <!-- Skip merged cell -->
    </xsl:template>
    
    <!-- Tag Name for Cell (th, td) -->
    <xsl:template name="get-cell-tag-name">
        <xsl:choose>
            <xsl:when test="parent::w:tr/w:trPr/w:tblHeader">
                <xsl:value-of select="$table-header-cell-tag-name"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$table-cell-tag-name"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Attributes for Table Cell -->
    <xsl:template name="insert-cell-attributes">
        <!-- colspan (for horizontal merged cells) -->
        <xsl:if test="w:tcPr/w:gridSpan">
            <xsl:attribute name="{$table-cell-column-span-attribute-name}">
                <xsl:value-of select="w:tcPr/w:gridSpan/@w:val"/>
            </xsl:attribute>
        </xsl:if>
        <!-- rowspan (for vertical merged cells) -->
        <xsl:if test="w:tcPr/w:vMerge[@w:val='restart']">
            <xsl:variable name="column-num">
                <xsl:call-template name="get-column-number"/>
            </xsl:variable>
            <xsl:attribute name="{$table-cell-row-span-attribute-name}">
                <xsl:call-template name="get-rowspan-value">
                    <xsl:with-param name="following-row-elements" select="parent::w:tr/following-sibling::w:tr"/>
                    <xsl:with-param name="target-column-num" select="$column-num"/>
                    <xsl:with-param name="num-of-rowspans" select="1"/>
                </xsl:call-template>
            </xsl:attribute>
        </xsl:if>
        <!-- scope (for header cells) -->
        <xsl:if test="parent::w:tr/w:trPr/w:tblHeader">
            <xsl:attribute name="{$table-header-cell-scope-attribute-name}">
                <xsl:value-of select="'col'"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Rowspan value for merged cells (w:vMerge) -->
    <xsl:template name="get-rowspan-value">
        <xsl:param name="following-row-elements"/>
        <xsl:param name="target-column-num"/>
        <xsl:param name="num-of-rowspans" select="1"/>
        <!-- First row -->
        <xsl:variable name="target-row-element" select="$following-row-elements[1]"/>
        <!-- Check: Is a merged cell in the examined column? -->
        <xsl:variable name="is-merged-cell">
            <xsl:for-each select="$target-row-element/w:tc">
                <xsl:variable name="cur-column-num">
                    <xsl:call-template name="get-column-number"/>
                </xsl:variable>
                <xsl:if test="$cur-column-num = $target-column-num">
                    <xsl:choose>
                        <xsl:when test="w:tcPr/w:vMerge[not(@w:val='restart')]">
                            <xsl:value-of select="'true'"/>
                        </xsl:when>
                        <xsl:otherwise>
                            <xsl:value-of select="'false'"/>
                        </xsl:otherwise>
                    </xsl:choose>
                </xsl:if>
            </xsl:for-each>
        </xsl:variable>
        <xsl:choose>
            <xsl:when test="$target-row-element and $is-merged-cell = 'true'">
                <xsl:call-template name="get-rowspan-value">
                    <xsl:with-param name="following-row-elements" select="$following-row-elements[position() &gt; 1]"/>
                    <xsl:with-param name="target-column-num" select="$target-column-num"/>
                    <xsl:with-param name="num-of-rowspans" select="$num-of-rowspans + 1"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$num-of-rowspans"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Column number (considering merged cells with w:gridSpan) -->
    <xsl:template name="get-column-number">
        <xsl:variable name="absolute-column-num" select="count(preceding-sibling::w:tc) + 1"/>
        <xsl:variable name="offset" select="sum(preceding-sibling::w:tc/w:tcPr/w:gridSpan/@w:val) - count(preceding-sibling::w:tc/w:tcPr/w:gridSpan[@w:val])"/>
        <xsl:value-of select="number($absolute-column-num) + number($offset)"/>
    </xsl:template>
    
    
    <!-- Structured Document Tags -->
    <xsl:template match="w:sdt">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Structured Document Block Content -->
    <xsl:template match="w:sdtContent[child::w:p]">
        <xsl:call-template name="structure-paragraphs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- Structured Document Run Content -->
    <xsl:template match="w:sdtContent[child::w:r]">
        <xsl:call-template name="structure-text-runs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    
    <!-- Block-Level Custom XML Element -->
    <xsl:template match="w:customXml[w:p]">
        <xsl:call-template name="structure-paragraphs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- Inline-Level Custom XML Element -->
    <xsl:template match="w:customXml[w:r]">
        <xsl:call-template name="structure-text-runs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    
    <!-- Track Changes -->
    <!-- Inserted Run Content -->
    <xsl:template match="w:ins[w:r]">
        <xsl:element name="{$inserted-text-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-track-change-attributes"/>
            <xsl:call-template name="structure-text-runs">
                <xsl:with-param name="target-elements" select="*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Inserted Run Content (ancestor::w:numPr or ancestor::w:rPr or ancestor::w:trPr or child::w:rPr) -->
    <xsl:template match="w:ins">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Deleted Run Content -->
    <xsl:template match="w:del[w:r]">
        <xsl:element name="{$deleted-text-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-track-change-attributes"/>
            <xsl:call-template name="structure-text-runs">
                <xsl:with-param name="target-elements" select="*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Deleted Run Content (ancestor::w:numPr or ancestor::w:rPr or ancestor::w:trPr or child::w:rPr) -->
    <xsl:template match="w:del">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Deleted Text -->
    <xsl:template match="w:delText">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Move Source Run Content -->
    <xsl:template match="w:moveFrom[w:r]">
        <xsl:element name="{$moved-from-text-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-track-change-attributes"/>
            <xsl:call-template name="structure-text-runs">
                <xsl:with-param name="target-elements" select="*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Move Source Run Content (ancestor::w:numPr or ancestor::w:rPr or ancestor::w:trPr or child::w:rPr) -->
    <xsl:template match="w:moveFrom">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Move Destination Run Content -->
    <xsl:template match="w:moveTo[w:r]">
        <xsl:element name="{$moved-to-text-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-track-change-attributes"/>
            <xsl:call-template name="structure-text-runs">
                <xsl:with-param name="target-elements" select="*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Move Destination Run Content (ancestor::w:numPr or ancestor::w:rPr or ancestor::w:trPr or child::w:rPr) -->
    <xsl:template match="w:moveTo">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Attributes for Track Changes -->
    <xsl:template name="insert-track-change-attributes">
        <xsl:attribute name="{$track-change-author-attribute-name}">
            <xsl:value-of select="@w:author"/>
        </xsl:attribute>
        <xsl:attribute name="{$track-change-date-attribute-name}">
            <xsl:value-of select="@w:date"/>
        </xsl:attribute>
    </xsl:template>
    
    
    
    
    <!-- Structure text runs (e.g. complex fields) -->
    <xsl:template name="structure-text-runs">
        <xsl:param name="target-elements"/>
        <xsl:param name="parent-field" select="''"/>
        <xsl:choose>
            <!-- Check: Are there any complex field starts? -->
            <xsl:when test="$target-elements/w:fldChar[@w:fldCharType='begin']">
                <xsl:variable name="field-begin-position">
                    <xsl:call-template name="get-field-begin-position">
                        <xsl:with-param name="target-elements" select="$target-elements"/>
                    </xsl:call-template>
                </xsl:variable>
                <xsl:variable name="field-end-position">
                    <xsl:call-template name="get-field-end-position">
                        <xsl:with-param name="target-elements" select="$target-elements"/>
                    </xsl:call-template>
                </xsl:variable>
                <!-- Elements before complex field (first level) -->
                <xsl:apply-templates select="$target-elements[position() &lt; $field-begin-position]"/>
                <!-- Elements of complex field (nested fields possible) -->
                <xsl:call-template name="transform-complex-field">
                    <xsl:with-param name="field-begin-element" select="$target-elements[position() = $field-begin-position]"/>
                    <xsl:with-param name="field-content-elements" select="$target-elements[position() &gt; $field-begin-position and position() &lt; $field-end-position]"/>
                    <xsl:with-param name="field-end-element" select="$target-elements[position() = $field-end-position]"/>
                    <xsl:with-param name="parent-field" select="$parent-field"/>
                </xsl:call-template>
                <!-- Elements after (first) complex field (first level) -->
                <xsl:call-template name="structure-text-runs">
                    <xsl:with-param name="target-elements" select="$target-elements[position() &gt; $field-end-position]"/>
                    <xsl:with-param name="parent-field" select="$parent-field"/>
                </xsl:call-template>
            </xsl:when>
            <!-- No complex field starts -->
            <xsl:otherwise>
                <xsl:apply-templates select="$target-elements"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Begin position of complex field -->
    <xsl:template name="get-field-begin-position">
        <xsl:param name="target-elements"/>
        <xsl:param name="position" select="0"/>
        <xsl:variable name="first-element" select="$target-elements[position() = 1]"/>
        <xsl:choose>
            <xsl:when test="not($first-element) or $first-element[w:fldChar/@w:fldCharType='begin']">
                <xsl:value-of select="number($position) + 1"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:call-template name="get-field-begin-position">
                    <xsl:with-param name="target-elements" select="$target-elements[position() &gt; 1]"/>
                    <xsl:with-param name="position" select="number($position) + 1"/>
                </xsl:call-template>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- End position of complex field -->
    <xsl:template name="get-field-end-position">
        <xsl:param name="target-elements"/>
        <xsl:param name="position" select="0"/>
        <xsl:param name="level" select="0"/>
        <xsl:variable name="first-element" select="$target-elements[position() = 1]"/>
        <xsl:choose>
            <xsl:when test="not($first-element) or ($first-element/w:fldChar[@w:fldCharType='end'] and $level = 1)">
                <xsl:value-of select="number($position) + 1"/> 
            </xsl:when>
            <xsl:otherwise>
                <xsl:call-template name="get-field-end-position">
                    <xsl:with-param name="target-elements" select="$target-elements[position() &gt; 1]"/>
                    <xsl:with-param name="position" select="number($position) + 1"/>
                    <xsl:with-param name="level">
                        <xsl:choose>
                            <xsl:when test="$first-element/w:fldChar[@w:fldCharType='begin']">
                                <xsl:value-of select="number($level) + 1"/>
                            </xsl:when>
                            <xsl:when test="$first-element/w:fldChar[@w:fldCharType='end']">
                                <xsl:value-of select="number($level) - 1"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="$level"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:with-param>
                </xsl:call-template>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Complex field transformation -->
    <xsl:template name="transform-complex-field">
        <xsl:param name="field-begin-element"/>
        <xsl:param name="field-content-elements"/>
        <xsl:param name="field-end-element"/>
        <xsl:param name="parent-field" select="''"/>
        <!-- Text content -->
        <xsl:variable name="complex-field-content">
            <xsl:call-template name="get-complex-field-content">
                <xsl:with-param name="following-sibling-elements" select="$field-begin-element/following-sibling::w:r"/>
            </xsl:call-template>
        </xsl:variable>
        <!-- Check: Type of complex field? -->
        <xsl:choose>
            <!-- Ignored field e.g. nested hyperlinks, Table of Content field -->
            <xsl:when test="
                starts-with($complex-field-content, 'TOC') or
                $parent-field = 'HYPERLINK' and (starts-with($complex-field-content, 'HYPERLINK') or starts-with($complex-field-content, 'REF') or starts-with($complex-field-content, 'PAGEREF') or starts-with($complex-field-content, 'NOTEREF'))
            ">
                <!-- Field elements -->
                <xsl:apply-templates select="$field-begin-element"/>
                <xsl:call-template name="structure-text-runs">
                    <xsl:with-param name="target-elements" select="$field-content-elements"/>
                    <xsl:with-param name="parent-field" select="$parent-field"/>
                </xsl:call-template>
                <xsl:apply-templates select="$field-end-element"/>
            </xsl:when>
            <!-- Index Mark -->
            <xsl:when test="starts-with($complex-field-content, 'XE')">
                <xsl:element name="{$indexmark-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-indexmark-attributes">
                        <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
                    </xsl:call-template> 
                </xsl:element>
                <!-- Field elements -->
                <xsl:apply-templates select="$field-begin-element"/>
                <xsl:call-template name="structure-text-runs">
                    <xsl:with-param name="target-elements" select="$field-content-elements"/>
                    <xsl:with-param name="parent-field" select="$parent-field"/>
                </xsl:call-template>
                <xsl:apply-templates select="$field-end-element"/>
            </xsl:when>
            <!-- Hyperlink (across several paragraphs) -->
            <xsl:when test="starts-with($complex-field-content, 'HYPERLINK')">
                <xsl:element name="{$hyperlink-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-complex-field-hyperlink-attributes">
                        <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
                    </xsl:call-template>
                    <!-- Field elements -->
                    <xsl:apply-templates select="$field-begin-element"/>
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="$field-content-elements"/>
                        <xsl:with-param name="parent-field" select="'HYPERLINK'"/>
                    </xsl:call-template>
                    <xsl:apply-templates select="$field-end-element"/>
                </xsl:element>
            </xsl:when>
            <!-- Cross References -->
            <xsl:when test="starts-with($complex-field-content, 'REF') or starts-with($complex-field-content, 'PAGEREF') or starts-with($complex-field-content, 'NOTEREF')">
                <xsl:element name="{$cross-reference-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-cross-reference-attributes">
                        <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
                    </xsl:call-template>
                    <!-- Field elements -->
                    <xsl:apply-templates select="$field-begin-element"/>
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="$field-content-elements"/>
                        <xsl:with-param name="parent-field" select="'HYPERLINK'"/>
                    </xsl:call-template>
                    <xsl:apply-templates select="$field-end-element"/>
                </xsl:element>
            </xsl:when>
            <!-- Date/Time -->
            <xsl:when test="
                starts-with($complex-field-content, 'TIME') or 
                starts-with($complex-field-content, 'DATE') or
                starts-with($complex-field-content, 'CREATEDATE') or
                starts-with($complex-field-content, 'PRINTDATE') or
                starts-with($complex-field-content, 'SAVEDATE')
            ">
                <xsl:element name="{$time-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-time-attributes">
                        <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
                    </xsl:call-template>
                    <!-- Field elements -->
                    <xsl:apply-templates select="$field-begin-element"/>
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="$field-content-elements"/>
                        <xsl:with-param name="parent-field" select="$parent-field"/>
                    </xsl:call-template>
                    <xsl:apply-templates select="$field-end-element"/>
                </xsl:element>
            </xsl:when>
            <!-- Citations -->
            <xsl:when test="starts-with($complex-field-content, 'CITATION')">
                <xsl:variable name="citation-id" select="substring-before(normalize-space(substring-after($complex-field-content, 'CITATION')), ' ')"/>
                <xsl:element name="{$citation-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-citation-attributes">
                        <xsl:with-param name="id" select="$citation-id"/>
                    </xsl:call-template>
                    <!-- Field elements (Call) -->
                    <xsl:element name="{$citation-call-tag-name}" namespace="{$ns}">
                        <xsl:apply-templates select="$field-begin-element"/>
                        <xsl:call-template name="structure-text-runs">
                            <xsl:with-param name="target-elements" select="$field-content-elements"/>
                            <xsl:with-param name="parent-field" select="$parent-field"/>
                        </xsl:call-template>
                        <xsl:apply-templates select="$field-end-element"/>
                    </xsl:element>
                    <!-- Source elements -->
                    <xsl:call-template name="insert-citation-source">
                        <xsl:with-param name="id" select="$citation-id"/>
                    </xsl:call-template>
                </xsl:element>
            </xsl:when>
            <!-- Other Complex Field -->
            <xsl:otherwise>
                <xsl:element name="{$complex-field-tag-name}" namespace="{$ns}">
                    <!-- Attributes -->
                    <xsl:call-template name="insert-general-complex-field-attributes">
                        <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
                        <xsl:with-param name="field-begin-element" select="$field-begin-element"/>
                    </xsl:call-template>
                    <!-- Field elements -->
                    <xsl:apply-templates select="$field-begin-element"/>
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="$field-content-elements"/>
                        <xsl:with-param name="parent-field" select="$parent-field"/>
                    </xsl:call-template>
                    <xsl:apply-templates select="$field-end-element"/>
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Content of complex field -->
    <xsl:template name="get-complex-field-content">
        <xsl:param name="following-sibling-elements"/>
        <xsl:param name="complex-field-text" select="''"/>
        <xsl:variable name="target-element" select="$following-sibling-elements[1]"/>
        <xsl:variable name="target-element-text">
            <xsl:value-of select="$target-element/w:instrText"/>
        </xsl:variable>
        <xsl:choose>
            <xsl:when test="
                $target-element and 
                not($target-element/w:fldChar[@w:fldCharType = 'end']) and 
                not($target-element/w:fldChar[@w:fldCharType = 'begin'])
                ">
                <xsl:call-template name="get-complex-field-content">
                    <xsl:with-param name="following-sibling-elements" select="$following-sibling-elements[position() &gt; 1]" />
                    <xsl:with-param name="complex-field-text" select="concat($complex-field-text, $target-element-text)"/>                    
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="normalize-space(concat($complex-field-text, $target-element-text))"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Attributes for indexmark element -->
    <xsl:template name="insert-indexmark-attributes">
        <xsl:param name="complex-field-content" select="''"/>
        <xsl:variable name="index-flags">
            <xsl:call-template name="get-complex-field-flags">
                <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
            </xsl:call-template>
        </xsl:variable>
        <!-- Style -->
        <xsl:attribute name="{$indexmark-style-attribute-name}">
            <xsl:value-of select="$indexmark-style-attribute-value"/>
        </xsl:attribute>
        <!-- Type -->
        <xsl:choose>
            <xsl:when test="contains($index-flags,'r')">
                <xsl:attribute name="{$indexmark-type-attribute-name}">
                    <xsl:value-of select="'r'"/>
                </xsl:attribute>
            </xsl:when>
            <xsl:when test="contains($index-flags,'t')">
                <xsl:attribute name="{$indexmark-type-attribute-name}">
                    <xsl:value-of select="'t'"/>
                </xsl:attribute>
            </xsl:when>
            <xsl:when test="not(contains($index-flags,'r') or contains($index-flags,'t'))">
                <xsl:attribute name="{$indexmark-type-attribute-name}">
                    <xsl:value-of select="'x'"/>
                </xsl:attribute>
            </xsl:when>
        </xsl:choose>
        <!-- Styling -->
        <xsl:choose>
            <xsl:when test="contains($index-flags,'b') and contains($index-flags,'i')">
                <xsl:attribute name="{$indexmark-format-attribute-name}">
                    <xsl:value-of select="'b i'"/>
                </xsl:attribute>
            </xsl:when>
            <xsl:when test="contains($index-flags,'b')">
                <xsl:attribute name="{$indexmark-format-attribute-name}">
                    <xsl:value-of select="'b'"/>
                </xsl:attribute>
            </xsl:when>
            <xsl:when test="contains($index-flags,'i')">
                <xsl:attribute name="{$indexmark-format-attribute-name}">
                    <xsl:value-of select="'i'"/>
                </xsl:attribute>
            </xsl:when>
        </xsl:choose>
        <!-- Entry (e.g. »Animal:Cat« for entry with subentry) -->
        <xsl:attribute name="{$indexmark-entry-attribute-name}">
            <xsl:value-of select="substring-before(substring-after($complex-field-content,'&quot;'), '&quot;')"/>
        </xsl:attribute>
        <!-- Target (page range or see here entry, e.g. name of the textmark/bookmark or see here entry) -->
        <xsl:attribute name="{$indexmark-target-attribute-name}">
            <xsl:value-of select="substring-before(substring-after(substring-after(substring-after($complex-field-content,'&quot;'), '&quot;'), '&quot;'), '&quot;')"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Attributes for complex field hyperlink -->
    <xsl:template name="insert-complex-field-hyperlink-attributes">
        <xsl:param name="complex-field-content" select="''"/>
        <!-- URI -->
        <xsl:attribute name="{$hyperlink-uri-attribute-name}">
            <xsl:call-template name="get-complex-field-hyperlink-uri">
                <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
            </xsl:call-template>
        </xsl:attribute>
        <!-- Title -->
        <xsl:variable name="title">     
            <xsl:value-of select="substring-before(substring-after($complex-field-content,'\o &quot;'), '&quot;')"/>      
        </xsl:variable>
        <xsl:if test="not($title = '')">
            <xsl:attribute name="{$hyperlink-title-attribute-name}">
                <xsl:value-of select="$title"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Value of hyperlink href attribute of complex field -->
    <xsl:template name="get-complex-field-hyperlink-uri">
        <xsl:param name="complex-field-content" select="''"/>
        <xsl:variable name="uri">
            <xsl:value-of select="substring-before(substring-after($complex-field-content,'&quot;'), '&quot;')"/>
        </xsl:variable>
        <xsl:variable name="anchor">
            <xsl:value-of select="substring-before(substring-after($complex-field-content,'\l &quot;'), '&quot;')"/>
        </xsl:variable>
        <xsl:choose>
            <xsl:when test="not($anchor = '')">
                <xsl:value-of select="concat($uri, '#', $anchor)"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$uri"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Attributes for cross referenc element -->
    <xsl:template name="insert-cross-reference-attributes">
        <xsl:param name="complex-field-content" select="''"/>
        <!-- URI -->
        <xsl:attribute name="{$cross-reference-uri-attribute-name}">
            <xsl:text>#</xsl:text>
            <xsl:value-of select="substring-before(substring-after($complex-field-content,'REF '), ' \')"/>
        </xsl:attribute>
        <!-- Type -->
        <xsl:attribute name="{$cross-reference-type-attribute-name}">
            <xsl:value-of select="substring-before($complex-field-content,' ')"/>
        </xsl:attribute>
        <!-- Format -->
        <xsl:attribute name="{$cross-reference-format-attribute-name}">
            <xsl:call-template name="get-complex-field-flags">
                <xsl:with-param name="complex-field-content" select="$complex-field-content"/>
            </xsl:call-template>
        </xsl:attribute>
    </xsl:template>

    <!-- Attributes for complex field time element -->
    <xsl:template name="insert-time-attributes">
        <xsl:param name="complex-field-content" select="''"/>
        <!-- Type -->
        <xsl:attribute name="{$time-type-attribute-name}">
            <xsl:value-of select="normalize-space(translate(substring-before($complex-field-content,' '), $uppercase, $lowercase))"/>
        </xsl:attribute>
        <!-- Format -->
        <xsl:attribute name="{$time-format-attribute-name}">
            <xsl:value-of select="substring-before(substring-after($complex-field-content,'&quot;'), '&quot;')"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Attributes for citation element -->
    <xsl:template name="insert-citation-attributes">
        <xsl:param name="id" select="''"/>
        <!-- Style Attribute -->
        <xsl:attribute name="{$citation-style-attribute-name}">
            <xsl:value-of select="$citation-style-attribute-value"/>
        </xsl:attribute>
        <!-- Type -->
        <xsl:attribute name="{$citation-style-type-attribute-name}">
            <xsl:value-of select="$citations-relationships/b:Sources[b:Source/b:Tag = $id]/@SelectedStyle"/>
        </xsl:attribute>
        <!-- Name -->
        <xsl:attribute name="{$citation-style-name-attribute-name}">
            <xsl:value-of select="$citations-relationships/b:Sources[b:Source/b:Tag = $id]/@StyleName"/>
        </xsl:attribute>
        <!-- Version -->
        <xsl:attribute name="{$citation-version-attribute-name}">
            <xsl:value-of select="$citations-relationships/b:Sources[b:Source/b:Tag = $id]/@Version"/>
        </xsl:attribute> 
    </xsl:template>
    
    <!-- Citation source element -->
    <xsl:template name="insert-citation-source">
        <xsl:param name="id" select="''"/>
        <xsl:element name="{$citation-source-tag-name}" namespace="{$ns}">
            <xsl:apply-templates select="$citations-relationships/b:Sources/b:Source[b:Tag = $id]/*" mode="citation-source"/>
        </xsl:element>
    </xsl:template>
       
    <xsl:template match="*" mode="citation-source">
        <xsl:element name="{local-name()}" namespace="{$ns}">
            <xsl:apply-templates select="@*|*|text()" mode="citation-source"/>
        </xsl:element>
    </xsl:template>
    
    <xsl:template match="@*" mode="citation-source">
        <xsl:attribute name="{local-name()}">
            <xsl:value-of select="."/>
        </xsl:attribute>
    </xsl:template>
    
    <xsl:template match="text()" mode="citation-source">
        <xsl:if test="normalize-space()">
            <xsl:element name="{$citation-text-tag-name}" namespace="{$ns}">
                <xsl:attribute name="{$citation-value-attribute-name}">
                    <xsl:value-of select="."/>
                </xsl:attribute>
            </xsl:element>
        </xsl:if>
    </xsl:template>
    
    <!-- Attributes for general complex field -->
    <xsl:template name="insert-general-complex-field-attributes">
        <xsl:param name="complex-field-content" select="''"/>
        <xsl:param name="field-begin-element"/>
        <!-- Style -->
        <xsl:attribute name="{$complex-field-style-attribute-name}">
            <xsl:value-of select="$complex-field-style-attribute-value"/>
        </xsl:attribute>
        <!-- Content -->
        <xsl:attribute name="{$complex-field-content-attribute-name}">
            <xsl:value-of select="$complex-field-content"/>
        </xsl:attribute>
        <!-- Data -->
        <xsl:if test="$field-begin-element/w:fldChar/w:fldData">
            <xsl:attribute name="{$complex-field-data-attribute-name}">
                <xsl:value-of select="$field-begin-element/w:fldChar/w:fldData"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Value of type attribute for index mark -->
    <!-- 
        »b« for bold style of entry 
        »i« for italic style of entry 
        »r« for page range with textmark/bookmark
        »t« for see here entry
        empty string for normal entry without style 
        a colon separates the entry levels (e.g. Animal:Cat)
    -->
    <xsl:template name="get-complex-field-flags">
        <xsl:param name="complex-field-content" select="''"/>
        <xsl:param name="index-flags" select="''"/>
        <xsl:choose>
            <xsl:when test="contains($complex-field-content, '\')">
                <xsl:variable name="target-flags">
                    <xsl:value-of select="substring(substring-after($complex-field-content,'\'), 1, 1)"/>
                </xsl:variable>
                <xsl:call-template name="get-complex-field-flags">
                    <xsl:with-param name="complex-field-content" select="substring-after($complex-field-content, '\')"/>
                    <xsl:with-param name="index-flags" select="concat($index-flags, ' ', $target-flags)"/>
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="normalize-space($index-flags)"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Complex Field (handeled in template »transform-complex-field«) -->
    <xsl:template match="w:fldChar">
        <xsl:apply-templates />
    </xsl:template>
    
    <xsl:template match="w:fldChar[@w:fldCharType='begin']">
        <xsl:if test="$is-comment-inserted">
            <xsl:comment>
                <xsl:value-of select="'complex field begin'"/>
            </xsl:comment>
        </xsl:if>
        <xsl:apply-templates />
    </xsl:template>
    
    <xsl:template match="w:fldChar[@w:fldCharType='separate']">
        <xsl:if test="$is-comment-inserted">
            <xsl:comment>
                <xsl:value-of select="'complex field separate'"/>
            </xsl:comment>
        </xsl:if>
        <xsl:apply-templates />
    </xsl:template>
    
    <xsl:template match="w:fldChar[@w:fldCharType='end']">
        <xsl:apply-templates />
        <xsl:if test="$is-comment-inserted">
            <xsl:comment>
                <xsl:value-of select="'complex field end'"/>
            </xsl:comment>
        </xsl:if>
    </xsl:template>
    
    
    <!-- Custom Field Data -->
    <xsl:template match="w:fldData">
        <!-- Content as attribute (handeled in template »transform-complex-field«) -->
    </xsl:template>
    
    <!-- Form Field Properties -->
    <xsl:template match="w:ffData">
        <!-- Empty content (handeled in template »transform-complex-field«) -->
    </xsl:template>
    
    <!-- Field Code -->
    <xsl:template match="w:instrText">
        <!-- Content as attribute (handeled in template »transform-complex-field«) -->
    </xsl:template>
    
    <!-- Deleted Field Code -->
    <xsl:template match="w:delInstrText">
        <!-- Skip content -->
    </xsl:template>
    



    <!-- Text Run -->
    <xsl:template match="w:r">
        <xsl:call-template name="assign-inline-styles">
            <xsl:with-param name="inline-style-elements" select="w:rPr/*"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- Recursive nesting of inline styles at text run level -->
    <!-- 
        Example:
        from WordML Structur
        
            <w:r>
                <w:rPr><w:b/><w:i/></w:rPr>
                <w:t>Text</w:t>
            </w:r>
            
        to XHTML Structur
        
            <strong><i>Text</i></strong>
    -->
    <xsl:template name="assign-inline-styles">
        <xsl:param name="inline-style-elements" />
        <xsl:param name="class-names" select="''"/>
        
        <xsl:choose>
            <xsl:when test="$inline-style-elements">
                <xsl:variable name="target-style-element" select="$inline-style-elements[1]"/>
                <xsl:choose>
                    <!-- Ignored Inline Styles -->
                    <xsl:when test=" 
                        w:footnoteReference or
                        w:footnoteRef or
                        w:endnoteReference or
                        w:endnoteRef or
                        w:commentReference or
                        w:annotationRef or
                        w:instrText or 
                        $target-style-element[name() = 'w:noProof'] or 
                        $target-style-element[name() = 'w:rFonts'] or
                        $target-style-element[name() = 'w:rPrChange'] or 
                        ($is-inline-style-on-empty-text-removed and w:t and normalize-space(w:t) = '')
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
                    <!-- Element: Bold | Complex Script Bold -->
                    <xsl:when test="$target-style-element[name() = 'w:b']"> <!-- w:bCs Complex Script Bold ignored -->
                        <xsl:element name="{$bold-tag-name}" namespace="{$ns}">
                            <xsl:call-template name="assign-inline-styles">
                                <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                <xsl:with-param name="class-names" select="$class-names"/>
                            </xsl:call-template>
                        </xsl:element>
                    </xsl:when>
                    <!-- Element: Italics | Complex Script Italics -->
                    <xsl:when test="$target-style-element[name() = 'w:i']"> <!-- w:iCs Complex Script Italics ignored -->
                        <xsl:element name="{$italics-tag-name}" namespace="{$ns}">
                            <xsl:call-template name="assign-inline-styles">
                                <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                <xsl:with-param name="class-names" select="$class-names"/>
                            </xsl:call-template>
                        </xsl:element>
                    </xsl:when>
                    <!-- Element: Underline -->
                    <xsl:when test="$target-style-element[name() = 'w:u']">
                        <xsl:element name="{$underline-tag-name}" namespace="{$ns}">
                            <xsl:if test="count($target-style-element/@*)">
                                <xsl:attribute name="{$inline-style-attribute-name}">
                                    <xsl:call-template name="join-attribute-names-and-values">
                                        <xsl:with-param name="local-tag-name" select="local-name($target-style-element)"/>
                                        <xsl:with-param name="attributes" select="$target-style-element/@*"/>
                                    </xsl:call-template>
                                </xsl:attribute>
                            </xsl:if>
                            <xsl:call-template name="assign-inline-styles">
                                <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                <xsl:with-param name="class-names" select="$class-names"/>
                            </xsl:call-template>
                        </xsl:element>
                    </xsl:when>
                    <!-- Element: Emphasis Mark -->
                    <xsl:when test="$target-style-element[name() = 'w:em']">
                        <xsl:element name="{$emphasis-mark-tag-name}" namespace="{$ns}">
                            <xsl:call-template name="assign-inline-styles">
                                <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                <xsl:with-param name="class-names" select="$class-names"/>
                            </xsl:call-template>
                        </xsl:element>
                    </xsl:when>
                    <!-- Element: Subscript/Superscript Text -->
                    <xsl:when test="$target-style-element[name() = 'w:vertAlign']">
                        <xsl:choose>
                            <xsl:when test="$inline-style-elements/@w:val = 'superscript'">
                                <xsl:element name="{$superscript-tag-name}" namespace="{$ns}">
                                    <xsl:call-template name="assign-inline-styles">
                                        <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                        <xsl:with-param name="class-names" select="$class-names"/>
                                    </xsl:call-template>
                                </xsl:element>
                            </xsl:when>
                            <xsl:when test="$inline-style-elements/@w:val = 'subscript'">
                                <xsl:element name="{$subscript-tag-name}" namespace="{$ns}">
                                    <xsl:call-template name="assign-inline-styles">
                                        <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                        <xsl:with-param name="class-names" select="$class-names"/>
                                    </xsl:call-template>
                                </xsl:element>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:call-template name="assign-inline-styles">
                                    <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                                    <xsl:with-param name="class-names" select="$class-names"/>
                                </xsl:call-template>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                    <!-- Attribute: class -->
                    <xsl:when test="
                        $is-local-override-without-tag-applied and 
                        (
                            $target-style-element[name() = 'w:caps'] or
                            $target-style-element[name() = 'w:smallCaps'] or
                            $target-style-element[name() = 'w:strike'] or
                            $target-style-element[name() = 'w:dstrike'] or
                            $target-style-element[name() = 'w:outline'] or
                            $target-style-element[name() = 'w:shadow'] or
                            $target-style-element[name() = 'w:emboss'] or
                            $target-style-element[name() = 'w:imprint'] or
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
                            $target-style-element[name() = 'w:highlight'] or
                            $target-style-element[name() = 'w:effect'] or
                            $target-style-element[name() = 'w:bdr'] or
                            $target-style-element[name() = 'w:shd'] or
                            $target-style-element[name() = 'w:fitText'] or
                            $target-style-element[name() = 'w:rtl'] or
                            $target-style-element[name() = 'w:cs'] or
                            $target-style-element[name() = 'w:lang'] or
                            $target-style-element[name() = 'w:eastAsianLayout'] or
                            $target-style-element[name() = 'w:specVanish'] or
                            $target-style-element[name() = 'w:oMath']
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
                            <xsl:attribute name="{$inline-style-attribute-name}">
                                <xsl:value-of select="normalize-space($class-names)"/>
                            </xsl:attribute>
                            <xsl:apply-templates/> <!-- w:r -->
                        </xsl:element>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:apply-templates/> <!-- w:r -->
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Join local attribute names and values -->
    <xsl:template name="join-attribute-names-and-values">
        <xsl:param name="local-tag-name" select="''"/>
        <xsl:param name="attributes"/>
        <xsl:if test="$local-tag-name">
            <xsl:value-of select="$local-tag-name"/>
            <xsl:if test="$attributes">
                <xsl:text>-</xsl:text>
            </xsl:if>
        </xsl:if>
        <xsl:for-each select="$attributes">
            <xsl:value-of select="local-name()"/>
            <xsl:text>_</xsl:text>
            <xsl:value-of select="."/>
            <xsl:if test="not(position() = last())">
                <xsl:text>-</xsl:text>
            </xsl:if>
        </xsl:for-each>
    </xsl:template>
    
    
    <!-- Hyperlink -->
    <xsl:template match="w:hyperlink">
        <xsl:element name="{$hyperlink-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-hyperlink-attribute"/>
            <xsl:call-template name="structure-text-runs">
                <xsl:with-param name="target-elements" select="*"/>
                <xsl:with-param name="parent-field" select="'HYPERLINK'"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for hyperlink -->
    <xsl:template name="insert-hyperlink-attribute">
        <!-- URI Attribute -->
        <xsl:variable name="uri">
            <xsl:choose>
                <xsl:when test="@r:id">
                    <xsl:call-template name="get-hyperlink-uri"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:text></xsl:text>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="anchor">
            <xsl:choose>
                <xsl:when test="@w:anchor">
                    <xsl:text>#</xsl:text>
                    <xsl:value-of select="@w:anchor"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:text></xsl:text>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="{$hyperlink-uri-attribute-name}">
            <xsl:choose>
                <xsl:when test="$uri = '' and $anchor = ''">
                    <xsl:value-of select="w:r/w:t"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="concat($uri, $anchor)"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
        <!-- Title Attribute -->
        <xsl:if test="boolean(@w:tooltip)">
            <xsl:attribute name="{$hyperlink-title-attribute-name}">
                <xsl:value-of select="@w:tooltip" />
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Value for hyperlink href attribute of hyperlink element -->
    <xsl:template name="get-hyperlink-uri">
        <xsl:variable name="id" select="@r:id"/>
        <xsl:choose>
            <xsl:when test="ancestor::w:footnote">
                <xsl:value-of select="$footnotes-relationships/rel:Relationship[@Id = $id]/@Target"/>
            </xsl:when>
            <xsl:when test="ancestor::w:endnote">
                <xsl:value-of select="$endnotes-relationships/rel:Relationship[@Id = $id]/@Target"/>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$document-relationships/rel:Relationship[@Id = $id]/@Target"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    
    <!-- Subdocument Anchor -->
    <xsl:template match="w:subDoc">
        <xsl:element name="{$subdocument-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-subdocument-attributes"/>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for subdocument -->
    <xsl:template name="insert-subdocument-attributes">
        <xsl:variable name="id" select="@r:id"/>
        <!-- Style -->
        <xsl:attribute name="{$subdocument-style-attribute-name}">
            <xsl:value-of select="$subdocument-style-attribute-value"/>
        </xsl:attribute>
        <!-- URI -->
        <xsl:attribute name="{$subdocument-uri-attribute-name}">
            <xsl:value-of select="$document-relationships/rel:Relationship[@Id = $id]/@Target"/>
        </xsl:attribute>
    </xsl:template>
    
    
    <!-- Smart Tag (Inline-Level) -->
    <xsl:template match="w:smartTag[w:r]">
        <xsl:call-template name="structure-text-runs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    
    <!-- Simple Field -->
    <xsl:template match="w:fldSimple[w:r]">
        <!-- Text content -->
        <xsl:variable name="simple-field-content" select="normalize-space(@w:instr)"/>
        <!-- Check: Type of complex field? -->
        <xsl:choose>
            <!-- Date/Time -->
            <xsl:when test="
                starts-with($simple-field-content, 'DATE') or
                starts-with($simple-field-content, 'CREATEDATE') or 
                starts-with($simple-field-content, 'PRINTDATE') or 
                starts-with($simple-field-content, 'SAVEDATE') or 
                starts-with($simple-field-content, 'TIME') 
            ">
                <xsl:element name="{$time-tag-name}" namespace="{$ns}">
                    <!-- Date/Time Type -->
                    <xsl:attribute name="{$time-type-attribute-name}">
                        <xsl:call-template name="get-simple-field-time-type">
                            <xsl:with-param name="simple-field-content" select="$simple-field-content"/>
                        </xsl:call-template>
                    </xsl:attribute>
                    <!-- Date/Time Format -->
                    <xsl:attribute name="{$time-format-attribute-name}">
                        <xsl:call-template name="get-simple-field-time-format">
                            <xsl:with-param name="simple-field-content" select="$simple-field-content"/>
                        </xsl:call-template>
                    </xsl:attribute>
                    <!-- Field elements -->
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="*"/>
                    </xsl:call-template>
                </xsl:element>
            </xsl:when>
            <!-- Other Simple Field -->
            <xsl:otherwise>
                <xsl:element name="{$data-tag-name}" namespace="{$ns}">
                    <!-- Value -->
                    <xsl:attribute name="{$data-value-attribute-name}">
                        <xsl:value-of select="$simple-field-content"/>
                    </xsl:attribute>
                    <!-- Field elements -->
                    <xsl:call-template name="structure-text-runs">
                        <xsl:with-param name="target-elements" select="*"/>
                    </xsl:call-template>
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose> 
    </xsl:template>
    
    <!-- Value of time format attribute of simple field -->
    <xsl:template name="get-simple-field-time-type">
        <xsl:param name="simple-field-content" select="''"/>
        <xsl:value-of select="normalize-space(translate(substring-before($simple-field-content,' '), $uppercase, $lowercase))"/>
    </xsl:template>
    
    <!-- Value of time format attribute of simple field -->
    <xsl:template name="get-simple-field-time-format">
        <xsl:param name="simple-field-content" select="''"/>
        <xsl:value-of select="normalize-space(substring-after($simple-field-content,'\*'))"/>
    </xsl:template>
    
    
    <!-- DayLong -->
    <xsl:template match="w:dayLong">
        <xsl:element name="{$time-tag-name}" namespace="{$ns}">
            <!-- Type -->
            <xsl:attribute name="{$time-type-attribute-name}">
                <xsl:value-of select="'day'"/>
            </xsl:attribute>
            <!-- Format -->
            <xsl:attribute name="{$time-format-attribute-name}">
                <xsl:value-of select="'long'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- DayShort -->
    <xsl:template match="w:dayShort">
        <xsl:element name="{$time-tag-name}" namespace="{$ns}">
            <!-- Type -->
            <xsl:attribute name="{$time-type-attribute-name}">
                <xsl:value-of select="'day'"/>
            </xsl:attribute>
            <!-- Format -->
            <xsl:attribute name="{$time-format-attribute-name}">
                <xsl:value-of select="'short'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- MonthLong -->
    <xsl:template match="w:monthLong">
        <xsl:element name="{$time-tag-name}" namespace="{$ns}">
            <!-- Type -->
            <xsl:attribute name="{$time-type-attribute-name}">
                <xsl:value-of select="'month'"/>
            </xsl:attribute>
            <!-- Format -->
            <xsl:attribute name="{$time-format-attribute-name}">
                <xsl:value-of select="'long'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- MonthShort -->
    <xsl:template match="w:monthShort">
        <xsl:element name="{$time-tag-name}" namespace="{$ns}">
            <!-- Type -->
            <xsl:attribute name="{$time-type-attribute-name}">
                <xsl:value-of select="'month'"/>
            </xsl:attribute>
            <!-- Format -->
            <xsl:attribute name="{$time-format-attribute-name}">
                <xsl:value-of select="'short'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Datumsblock – Langes Jahresformat - YearLong -->
    <xsl:template match="w:yearlong">
        <xsl:element name="{$time-tag-name}" namespace="{$ns}">
            <!-- Type -->
            <xsl:attribute name="{$time-type-attribute-name}">
                <xsl:value-of select="'year'"/>
            </xsl:attribute>
            <!-- Format -->
            <xsl:attribute name="{$time-format-attribute-name}">
                <xsl:value-of select="'long'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Datumsblock – Kurzes Jahresformat - YearShort -->
    <xsl:template match="w:yearShort">
        <xsl:element name="{$time-tag-name}" namespace="{$ns}">
            <!-- Type -->
            <xsl:attribute name="{$time-type-attribute-name}">
                <xsl:value-of select="'year'"/>
            </xsl:attribute>
            <!-- Format -->
            <xsl:attribute name="{$time-format-attribute-name}">
                <xsl:value-of select="'short'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    

    <!-- Mathematical Equation -->
    <xsl:template match="m:oMath">
        <xsl:element name="{$equation-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-equation-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>

    <!-- Attributes for Mathematical Equation -->
    <xsl:template name="insert-equation-attributes">
        <!-- Style -->
        <xsl:attribute name="{$equation-style-attribute-name}">
            <xsl:value-of select="$equation-style-attribute-value"/>
        </xsl:attribute>
    </xsl:template>




    <!-- Text Element -->
    <xsl:template match="w:t">
        <xsl:apply-templates />
    </xsl:template>
    
    
    <!-- Group -->
    <xsl:template match="wpg:wgp">
        <xsl:element name="{$group-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-group-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Group -->
    <xsl:template name="insert-group-attributes">
        <!-- Style -->
        <xsl:attribute name="{$group-style-attribute-name}">
            <xsl:value-of select="$group-style-attribute-value"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Drawing -->
    <xsl:template match="w:drawing">
        <xsl:apply-templates /> <!-- includes Graphics, ... -->
    </xsl:template>
    
    
    <!-- Graphic -->
    <xsl:template match="a:graphic">
        <xsl:apply-templates /> <!-- includes Groups, Shapes, Images, Textboxes ... -->
    </xsl:template>
    
    <!-- Graphic horizontal position -->
    <xsl:template match="wp:positionH">
        <!-- Skip position value -->
    </xsl:template>
    
    <!-- Graphic vertical position -->
    <xsl:template match="wp:positionV">
        <!-- Skip position value -->
    </xsl:template>
    
    <!-- Graphic relative width horizontal-->
    <xsl:template match="wp14:sizeRelH">
        <!-- Skip width value -->
    </xsl:template>
    
    <!-- Graphic relative width vertical -->
    <xsl:template match="wp14:sizeRelV">
        <!-- Skip width value -->
    </xsl:template>
    
    
    <!-- Image -->
    <xsl:template match="pic:pic">
        <xsl:element name="{$image-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-image-attributes"/>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Image -->
    <xsl:template name="insert-image-attributes">
        <xsl:variable name="image-id" select="pic:blipFill/a:blip/@r:embed | pic:blipFill/a:blip/@r:link"/>
        <xsl:variable name="image-rel" select="$document-relationships/rel:Relationship[@Id = $image-id]"/>
        <xsl:variable name="image-target" select="$image-rel/@Target"/>
        <xsl:variable name="target-mode" select="$image-rel/@TargetMode"/>
        <xsl:variable name="image-name">
            <xsl:call-template name="substring-after-last">
                <xsl:with-param name="target-string" select="$image-target"/>
                <xsl:with-param name="delimiter" select="'/'"/>
            </xsl:call-template>
        </xsl:variable>
        <!-- Style -->
        <xsl:attribute name="{$image-style-attribute-name}">
            <xsl:choose>
                <xsl:when test="ancestor::wp:inline">
                    <xsl:value-of select="$image-style-inline-attribute-value"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="$image-style-anchored-attribute-value"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
        <!-- Source -->
        <xsl:variable name="source">
            <xsl:choose>
                <xsl:when test="$image-folder-path">
                    <xsl:value-of select="concat($image-folder-path, $directory-separator, $image-name)"/>
                </xsl:when>
                <xsl:when test="$target-mode = 'External' and contains($image-target, 'file:///')">
                    <xsl:value-of select="substring-after($image-target, 'file:///')"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="$image-target"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="{$image-source-attribute-name}">
            <xsl:value-of select="$source"/>
        </xsl:attribute>
        <!-- Title -->
        <xsl:variable name="title" select="normalize-space(pic:nvPicPr/pic:cNvPr/@title)"/>
        <xsl:if test="$title">
            <xsl:attribute name="{$image-title-attribute-name}">
                <xsl:value-of select="$title"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Description (alt text) -->
        <xsl:variable name="descr" select="normalize-space(pic:nvPicPr/pic:cNvPr/@descr)"/>
        <xsl:if test="$descr">
            <xsl:attribute name="{$image-alt-attribute-name}">
                <xsl:value-of select="$descr"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Position (inline or anchored) -->
        <xsl:variable name="position">
            <xsl:choose>
                <xsl:when test="ancestor::wp:inline">
                    <xsl:value-of select="'inline'"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="'anchored'"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:attribute name="{$image-position-attribute-name}">
            <xsl:value-of select="$position"/>
        </xsl:attribute>
        <!-- URI (Hyperlink assigned to image) -->
        <xsl:variable name="hlink-id" select="pic:nvPicPr/pic:cNvPr/a:hlinkClick/@r:id"/>
        <xsl:variable name="hlink-rel" select="$document-relationships/rel:Relationship[@Id = $hlink-id]"/>
        <xsl:variable name="hlink-target" select="$hlink-rel/@Target"/>
        <xsl:variable name="hlink-uri" select="normalize-space($hlink-target)"/>          
        <xsl:if test="$hlink-uri">
            <xsl:attribute name="{$image-uri-attribute-name}">
                <xsl:value-of select="$hlink-uri"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    
    <!-- Shape (WordprocessingShape) -->
    <xsl:template match="wps:wsp">
        <xsl:element name="{$shape-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-shape-attributes"/>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Shape -->
    <xsl:template name="insert-shape-attributes">
        <xsl:variable name="id">
            <xsl:choose>
                <xsl:when test="wps:cNvPr">
                    <xsl:value-of select="wps:cNvPr/@id"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="ancestor::a:graphic/preceding-sibling::wp:docPr[1]/@id"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="name">
            <xsl:choose>
                <xsl:when test="wps:cNvPr">
                    <xsl:value-of select="wps:cNvPr/@name"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="ancestor::a:graphic/preceding-sibling::wp:docPr[1]/@name"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="descr">
            <xsl:choose>
                <xsl:when test="wps:cNvPr">
                    <xsl:value-of select="wps:cNvPr/@descr"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="ancestor::a:graphic/preceding-sibling::wp:docPr[1]/@descr"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <!-- Index -->
        <xsl:if test="normalize-space($id)">
            <xsl:attribute name="{$shape-index-attribute-name}">
                <xsl:value-of select="$id"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Style -->
        <xsl:attribute name="{$shape-style-attribute-name}">
            <xsl:value-of select="$shape-style-attribute-value"/>
        </xsl:attribute>
        <!-- Name -->
        <xsl:if test="normalize-space($name)">
            <xsl:attribute name="{$shape-name-attribute-name}">
                <xsl:value-of select="$name"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Description (ALT text) -->
        <xsl:if test="normalize-space($descr)">
            <xsl:attribute name="{$shape-alt-attribute-name}">
                <xsl:value-of select="$descr"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    
    <!-- Textbox -->
    <xsl:template match="wps:wsp[descendant::w:txbxContent]">
        <xsl:element name="{$textbox-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-textbox-attributes"/>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Textbox -->
    <xsl:template name="insert-textbox-attributes">
        <xsl:variable name="id">
            <xsl:choose>
                <xsl:when test="wps:cNvPr">
                    <xsl:value-of select="wps:cNvPr/@id"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="ancestor::a:graphic/preceding-sibling::wp:docPr[1]/@id"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="name">
            <xsl:choose>
                <xsl:when test="wps:cNvPr">
                    <xsl:value-of select="wps:cNvPr/@name"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="ancestor::a:graphic/preceding-sibling::wp:docPr[1]/@name"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <xsl:variable name="descr">
            <xsl:choose>
                <xsl:when test="wps:cNvPr">
                    <xsl:value-of select="wps:cNvPr/@descr"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="ancestor::a:graphic/preceding-sibling::wp:docPr[1]/@descr"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:variable>
        <!-- Index -->
        <xsl:if test="normalize-space($id)">
            <xsl:attribute name="{$textbox-index-attribute-name}">
                <xsl:value-of select="$id"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Style -->
        <xsl:attribute name="{$textbox-style-attribute-name}">
            <xsl:value-of select="$textbox-style-attribute-value"/>
        </xsl:attribute>
        <!-- Name -->
        <xsl:if test="normalize-space($name)">
            <xsl:attribute name="{$textbox-name-attribute-name}">
                <xsl:value-of select="$name"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Description (ALT text) -->
        <xsl:if test="normalize-space($descr)">
            <xsl:attribute name="{$textbox-alt-attribute-name}">
                <xsl:value-of select="$descr"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Textbox Content -->
    <xsl:template match="w:txbxContent">
        <xsl:call-template name="structure-paragraphs">
            <xsl:with-param name="target-elements" select="*"/>
        </xsl:call-template>
    </xsl:template>
    
    <!-- Fallback content of Textbox -->
    <xsl:template match="mc:Fallback">
        <!-- Skip fallback elements -->
    </xsl:template>
    
    
    <!-- Footnote Reference -->
    <!-- (for inline style of footnote call see template »assign-inline-styles«) -->
    <xsl:template match="w:footnoteReference">
        <xsl:variable name="id" select="@w:id"/>
        <xsl:element name="{$footnote-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-footnote-attributes">
                <xsl:with-param name="id" select="$id"/>
            </xsl:call-template>
            <xsl:call-template name="structure-paragraphs">
                <xsl:with-param name="target-elements" select="$footnotes/w:footnote[@w:id = $id]/*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Footnote -->
    <xsl:template name="insert-footnote-attributes">
        <xsl:param name="id" select="''"/>
        <!-- ID -->
        <xsl:attribute name="{$footnote-index-attribute-name}">
            <xsl:value-of select="$id"/>
        </xsl:attribute>
        <!-- Style -->
        <xsl:attribute name="{$footnote-style-attribute-name}">
            <xsl:value-of select="$footnote-style-attribute-value"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Footnote Reference Marker -->
    <!-- (for inline style of footnote marker see template »assign-inline-styles«) -->
    <xsl:template match=" w:footnoteRef">
        <!--<xsl:element name="{$footnote-reference-tag-name}" namespace="{$ns}">-->
            <xsl:apply-templates />
        <!--</xsl:element>-->
    </xsl:template>
    
    
    <!-- Endnote Reference -->
    <!-- (for inline style of endnote call see template »assign-inline-styles«) -->
    <xsl:template match="w:endnoteReference">
        <xsl:variable name="id" select="@w:id"/>
        <xsl:element name="{$endnote-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-endnote-attributes">
                <xsl:with-param name="id" select="$id"/>
            </xsl:call-template>
            <xsl:call-template name="structure-paragraphs">
                <xsl:with-param name="target-elements" select="$endnotes/w:endnote[@w:id = $id]/*"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Endnote -->
    <xsl:template name="insert-endnote-attributes">
        <xsl:param name="id" select="''"/>
        <!-- ID -->
        <xsl:attribute name="{$endnote-index-attribute-name}">
            <xsl:value-of select="$id"/>
        </xsl:attribute>
        <!-- Style -->
        <xsl:attribute name="{$endnote-style-attribute-name}">
            <xsl:value-of select="$endnote-style-attribute-value"/>
        </xsl:attribute>
    </xsl:template>
    
    <!-- Endnote Reference Marker -->
    <!-- (for inline style of endnote marker see template »assign-inline-styles«) -->
    <xsl:template match="w:endnoteRef">
        <!--<xsl:element name="{$endnote-reference-tag-name}" namespace="{$ns}">-->
            <xsl:apply-templates />
        <!--</xsl:element>-->
    </xsl:template>
    
    
    <!-- Footnote/Endnote Separator Mark -->
    <xsl:template match="w:separator">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Continuation Separator Mark -->
    <xsl:template match="w:continuationSeparator">
        <xsl:apply-templates />
    </xsl:template>
    
    
    <!-- Comment -->
    <!-- 
        for inline style of comment call and comment marker 
        see template »assign-inline-styles« above
    -->
    <xsl:template match="w:commentReference">
        <xsl:variable name="id" select="@w:id"/>
        <xsl:element name="{$comment-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-comment-attributes">
                <xsl:with-param name="id" select="$id"/>
            </xsl:call-template>
            <xsl:call-template name="structure-paragraphs">
                <xsl:with-param name="target-elements" select="$comments/w:comment[@w:id = $id]"/>
            </xsl:call-template>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for comment element -->
    <xsl:template name="insert-comment-attributes">
        <xsl:param name="id"/>
        <xsl:attribute name="{$comment-index-attribute-name}">
            <xsl:value-of select="$id"/>
        </xsl:attribute>
        <xsl:attribute name="{$comment-style-attribute-name}">
            <xsl:value-of select="$comment-style-attribute-value"/>
        </xsl:attribute>
        <xsl:attribute name="{$comment-date-attribute-name}">
            <xsl:value-of select="$comments/w:comment[@w:id = $id]/@w:date" />
        </xsl:attribute>
        <xsl:attribute name="{$comment-initials-attribute-name}">
            <xsl:value-of select="$comments/w:comment[@w:id = $id]/@w:initials" />
        </xsl:attribute>
        <xsl:attribute name="{$comment-author-attribute-name}">
            <xsl:value-of select="$comments/w:comment[@w:id = $id]/@w:author" />
        </xsl:attribute>
    </xsl:template>
    
    <!-- Comment Reference Mark -->
    <xsl:template match=" w:annotationRef">
        <!--<xsl:element name="{$comment-reference-tag-name}" namespace="{$ns}">-->
            <xsl:apply-templates />
        <!--</xsl:element>-->
    </xsl:template>
    
    <!-- Comment Anchor Range Start -->
    <xsl:template match="w:commentRangeStart">
        <xsl:if test="$is-comment-inserted">
            <xsl:variable name="id" select="@w:id"/>
            <xsl:comment>
                <xsl:value-of select="concat('comment ', $id ,' range start')"/>
            </xsl:comment>
        </xsl:if>
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Comment Anchor Range End -->
    <xsl:template match="w:commentRangeEnd">
        <xsl:apply-templates />
        <xsl:if test="$is-comment-inserted">
            <xsl:variable name="id" select="@w:id"/>
            <xsl:comment>
                <xsl:value-of select="concat('comment ', $id ,' range end')"/>
            </xsl:comment>
        </xsl:if>
    </xsl:template>
    
    
    <!-- Bookmark Start -->
    <xsl:template match="w:bookmarkStart">
        <xsl:variable name="id" select="@w:id"/>
        <xsl:variable name="name" select="@w:name"/>
        <xsl:if test="$is-comment-inserted">
            <xsl:comment>
                <xsl:value-of select="concat('bookmark ', $id , ' ', $name, ' start')"/>
            </xsl:comment>
        </xsl:if>
        <xsl:choose>
            <!-- Check: Is _GoBack Bookmark? -->
            <xsl:when test="$name = '_GoBack'">
                <xsl:apply-templates />
            </xsl:when>
            <xsl:otherwise>
                <xsl:element name="{$bookmark-tag-name}" namespace="{$ns}">
                    <xsl:call-template name="insert-bookmark-attributes">
                        <xsl:with-param name="id" select="$id"/>
                    </xsl:call-template>
                    <xsl:apply-templates />
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Attributes for bookmark -->
    <xsl:template name="insert-bookmark-attributes">
        <xsl:param name="id"></xsl:param>
        <!-- Index -->
        <xsl:attribute name="{$bookmark-index-attribute-name}">
            <xsl:value-of select="$id"/>
        </xsl:attribute>
        <!-- Style -->
        <xsl:attribute name="{$bookmark-style-attribute-name}">
            <xsl:value-of select="$bookmark-style-attribute-value"/>
        </xsl:attribute>
        <!-- ID -->
        <xsl:attribute name="{$bookmark-id-attribute-name}">
            <xsl:value-of select="@w:name"/>
        </xsl:attribute>
        <!-- Content -->
        <xsl:variable name="content">
            <xsl:call-template name="get-bookmark-content">
                <xsl:with-param name="following-elements" select="following::w:r|following::w:bookmarkEnd[@w:id = $id]"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:if test="not($content = '')">
            <xsl:attribute name="{$bookmark-content-attribute-name}">
                <xsl:value-of select="$content"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Get content of bookmark -->
    <xsl:template name="get-bookmark-content">
        <xsl:param name="following-elements"/>
        <xsl:param name="bookmark-text" select="''"/>
        <xsl:variable name="target-element" select="$following-elements[1]"/>
        <xsl:variable name="target-element-text">
            <xsl:apply-templates select="$target-element/w:t"/>
        </xsl:variable>
        <xsl:choose>
            <xsl:when test="$target-element and not($target-element[name() = 'w:bookmarkEnd']) and string-length($bookmark-text) &lt; $max-bookmark-length">
                <xsl:call-template name="get-bookmark-content">
                    <xsl:with-param name="following-elements" select="$following-elements[position() &gt; 1]" />
                    <xsl:with-param name="bookmark-text" select="concat($bookmark-text, $target-element-text)"/>                    
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="normalize-space(concat($bookmark-text, $target-element-text))"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    <!-- Bookmark End -->
    <xsl:template match="w:bookmarkEnd">
        <xsl:apply-templates />
        <xsl:if test="$is-comment-inserted">
            <xsl:variable name="id" select="@w:id"/>
            <xsl:variable name="name" select="@w:name"/>
            <xsl:comment>
                <xsl:value-of select="concat('bookmark ', $id, ' ', $name, ' end')"/>
            </xsl:comment>
        </xsl:if>
    </xsl:template>
    
    
    <!-- Symbol Character -->
    <xsl:template match="w:sym">
        <xsl:element name="{$symbol-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-symbol-attributes"/>
            <xsl:apply-templates/>
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes of Symbol Character -->
    <xsl:template name="insert-symbol-attributes">
        <!-- Style -->
        <xsl:attribute name="{$symbol-style-attribute-name}">
            <xsl:value-of select="$symbol-style-attribute-value"/>
        </xsl:attribute>
        <!-- Font -->
        <xsl:attribute name="{$symbol-font-attribute-name}">
            <xsl:value-of select="@w:font"/>
        </xsl:attribute>
        <!-- Code -->
        <xsl:attribute name="{$symbol-code-attribute-name}">
            <xsl:choose>
                <xsl:when test="@w:code">
                    <xsl:value-of select="@w:code"/>
                </xsl:when>
                <xsl:when test="@w:char">
                    <xsl:value-of select="@w:char"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="'FFFD'"/>
                </xsl:otherwise>
            </xsl:choose>
        </xsl:attribute>
    </xsl:template>
    
    
    <!-- Inline Embedded Object -->
    <xsl:template match="w:object">
        <xsl:element name="{$embedded-object-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-object-attributes"/>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Inline Embedded Object -->
    <xsl:template name="insert-object-attributes">
        <xsl:variable name="id" select="o:OLEObject/@r:id"/>
        <xsl:variable name="target" select="$document-relationships/rel:Relationship[@Id = $id]/@Target"/>
        <!-- Style -->
        <xsl:attribute name="{$embedded-object-style-attribute-name}">
            <xsl:value-of select="$embedded-object-style-attribute-value"/>
        </xsl:attribute>
        <!-- Target -->
        <xsl:if test="$target">
            <xsl:attribute name="{$embedded-object-target-attribute-name}">
                <xsl:value-of select="$target"/>
            </xsl:attribute>
        </xsl:if>
        <!-- Programm -->
        <xsl:if test="o:OLEObject/@ProgID">
            <xsl:attribute name="{$embedded-object-program-attribute-name}">
                <xsl:value-of select="o:OLEObject/@ProgID"/>
            </xsl:attribute>
        </xsl:if>
    </xsl:template>
    
    <!-- Object properties -->
    <xsl:template match="o:OLEObject">
        <!-- Skip object properties -->
    </xsl:template>
    
    
    <!-- Previous Numbering Field Properties -->
    <xsl:template match="w:numberingChange">
        <!-- Empty content -->
    </xsl:template>
    
    
    <!-- Break -->
    <xsl:template match="w:br">
        <xsl:choose>	
            <xsl:when test="boolean(@w:type='page')">
                <!-- Seitenumbruch einfuegen -->
                <xsl:element name="{$page-break-tag-name}" namespace="{$ns}">
                    <xsl:attribute name="{$inline-style-attribute-name}">
                        <xsl:value-of select="'page-break-after'"/>
                    </xsl:attribute> 
                </xsl:element>
            </xsl:when>
            <xsl:when test="boolean(@w:type='column')">
                <!-- Spaltenumbruch einfuegen -->
                <xsl:element name="{$column-break-tag-name}" namespace="{$ns}">
                    <xsl:attribute name="{$inline-style-attribute-name}">
                        <xsl:value-of select="'column-break-after'"/>
                    </xsl:attribute> 
                </xsl:element>
            </xsl:when>
            <xsl:otherwise>
                <xsl:element name="{$forced-line-break-tag-name}" namespace="{$ns}"> 
                </xsl:element>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    
    <!-- Carriage Return -->
    <xsl:template match="w:cr">
        <xsl:element name="{$carriage-return-tag-name}" namespace="{$ns}"> 
        </xsl:element>
    </xsl:template>
    
    
    <!-- Soft Hyphen -->
    <xsl:template match="w:softHyphen">
        <xsl:text>&#xAD;</xsl:text>
    </xsl:template>
    
    
    <!-- No Break Hyphen -->
    <xsl:template match="w:noBreakHyphen">
        <xsl:text>&#x2011;</xsl:text>
    </xsl:template>
    
    
    <!-- Tab Char | Positional Tab -->
    <xsl:template match="w:tab[parent::w:r] | w:pTab[parent::w:r]">
        <xsl:element name="{$tab-tag-name}" namespace="{$ns}">
            <xsl:attribute name="{$tab-style-attribute-name}">
                <xsl:value-of select="$tab-style-attribute-value"/>
            </xsl:attribute>
            <!-- Comment -->
            <xsl:if test="$is-comment-inserted">
                <xsl:comment>
                    <xsl:text>tab</xsl:text>
                </xsl:comment>
            </xsl:if>
            <!-- Tab Character -->
            <xsl:if test="$is-tab-preserved">
                <xsl:text>&#x09;</xsl:text>
            </xsl:if>
        </xsl:element>
    </xsl:template>
    
    
    <!-- Position of Last Calculated Page Break -->
    <xsl:template match="w:lastRenderedPageBreak">
        <xsl:element name="{$section-break-tag-name}" namespace="{$ns}">
            <xsl:attribute name="{$section-break-type-attribute-name}">
                <xsl:value-of select="'section-break'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    
    
    
    <!-- +++++++++++++++++++++ -->
    <!-- + General templates + -->
    <!-- +++++++++++++++++++++ -->
    
    <!-- 
        Substring after last
        Determines the substring after the last occurrence of a given character (delimiter). 
    -->
    <xsl:template name="substring-after-last">
        <xsl:param name="target-string" />
        <xsl:param name="delimiter" />
        <xsl:choose>
            <xsl:when test="$target-string and $delimiter and contains($target-string, $delimiter)">
                <xsl:call-template name="substring-after-last">
                    <xsl:with-param name="target-string" select="substring-after($target-string, $delimiter)" />
                    <xsl:with-param name="delimiter" select="$delimiter" />
                </xsl:call-template>
            </xsl:when>
            <xsl:otherwise>
                <xsl:value-of select="$target-string" />
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
</xsl:transform>