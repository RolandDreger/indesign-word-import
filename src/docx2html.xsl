<?xml version="1.0" encoding="UTF-8"?>

<!--    
        
    Microsoft Word Document -> HTML
    
    Created: September 30, 2021
    Modified: November 09, 2025
    
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
    exclude-result-prefixes="rd pkg wpc cx cx1 cx2 cx3 cx4 cx5 cx6 cx7 cx8 mc aink am3d o r rel m v wp14 wp w10 w w14 w15 w16cex w16cid w16 w16sdtdh w16se wpg wpi wne wps cp extp cusp vt dc dcterms dcmitype dcmitype a pic xsi b"
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
    <xsl:param name="is-comment-inserted" select="false()"/> <!-- Comments for Complex Fields, Tab, ... -->
    <xsl:param name="is-tab-preserved" select="true()"/>  <!-- Tab Character -->
    <xsl:param name="style-mode" select="'minimized'"/> <!-- Values: 'extended' or 'minimized'. If minimized, ignore all local overrides (without tags, only via class) except: strong, i, em, u, superscript, subscript  -->
    <xsl:param name="math-mode" select="'mml'"/> <!-- Values: 'mml' or 'text' or 'remove'. -->
    
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
                        <xsl:variable name="custom-xml-relative-file-path" select="substring-after(@Target, '../')"/>
                        <xsl:if test="not($custom-xml-relative-file-path = '') and starts-with($custom-xml-relative-file-path, 'customXml')">
                            <xsl:variable name="custom-xml-file-path" select="concat($package-base-uri, $directory-separator, $custom-xml-relative-file-path)"/>
                            <xsl:variable name="sources-element" select="document($custom-xml-file-path)/b:Sources"/>
                            <xsl:if test="$sources-element and namespace-uri($sources-element) = 'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'">
                                <xsl:copy-of select="document($custom-xml-file-path)/b:Sources"/>
                            </xsl:if>
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
    <xsl:variable name="list-id-attribute-name" select="'data-list-id'"/>
    <xsl:variable name="list-item-level-attribute-name" select="'data-item-level'"/>
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
    <xsl:variable name="equation-tag-name" select="'span'"/>
    <xsl:variable name="equation-style-attribute-name" select="'class'"/>
    <xsl:variable name="equation-style-attribute-value" select="'math-expression'"/>
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
                        <xsl:with-param name="list-id" select="$first-num-id"/>
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
                        <xsl:with-param name="list-id" select="$first-num-id"/>
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
                            <xsl:with-param name="list-id" select="$first-lower-level-num-id"/>
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
                            <xsl:with-param name="list-id" select="$first-lower-level-num-id"/>
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
        <xsl:param name="list-id" select="''"/>
        <xsl:param name="list-item-level" select="''"/>
        <!-- ID -->
        <xsl:attribute name="{$list-id-attribute-name}">
            <xsl:value-of select="$list-id"/>
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
                    <xsl:call-template name="insert-paragraph-attributes"/>
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
        <xsl:variable name="paragraph-style-id" select="$target-element/w:pPr/w:pStyle/@w:val"/>
        <xsl:variable name="paragraph-style-name">
            <xsl:call-template name="get-style-name">
                <xsl:with-param name="style-id" select="$paragraph-style-id"/>
                <xsl:with-param name="style-type" select="'paragraph'"/>
            </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="heading-level">
            <xsl:choose>
                <!-- Heading: Default WordML Name -->
                <xsl:when test="starts-with($paragraph-style-name, 'heading')">
                    <xsl:variable name="heading-level-string" select="substring-after($paragraph-style-name, 'heading')"/>
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
                    translate($paragraph-style-name, 'H ', 'h') = 'h1' or 
                    translate($paragraph-style-name, 'H ', 'h') = 'h2' or
                    translate($paragraph-style-name, 'H ', 'h') = 'h3' or
                    translate($paragraph-style-name, 'H ', 'h') = 'h4' or
                    translate($paragraph-style-name, 'H ', 'h') = 'h5' or
                    translate($paragraph-style-name, 'H ', 'h') = 'h6'
                ">
                    <xsl:value-of select="number(substring-after(translate($paragraph-style-name, 'H','h'), 'h'))"/>
                </xsl:when>
                <!-- Heading: User Paragraph Name with $heading-marker -->
                <xsl:when test="contains($paragraph-style-name, $heading-marker)">
                    <xsl:value-of select="number(substring-after(translate($paragraph-style-name, 'h','H'), $heading-marker))"/>
                </xsl:when>
                <!-- Heading: User Paragraph Name in $h1-paragraph-style-names -->
                <xsl:when test="boolean($h1-paragraph-style-names)">
                    <xsl:choose>
                        <xsl:when test="($h1-paragraph-style-names = $paragraph-style-name or (contains($h1-paragraph-style-names, concat('»', $paragraph-style-name, '«'))))">
                            <xsl:value-of select="1"/>
                        </xsl:when>
                        <xsl:when test="($h2-paragraph-style-names = $paragraph-style-name or (contains($h2-paragraph-style-names, concat('»', $paragraph-style-name, '«'))))">
                            <xsl:value-of select="2"/>
                        </xsl:when>
                        <xsl:when test="($h3-paragraph-style-names = $paragraph-style-name or (contains($h3-paragraph-style-names, concat('»', $paragraph-style-name, '«'))))">
                            <xsl:value-of select="3"/>
                        </xsl:when>
                        <xsl:when test="($h4-paragraph-style-names = $paragraph-style-name or (contains($h4-paragraph-style-names, concat('»', $paragraph-style-name, '«'))))">
                            <xsl:value-of select="4"/>
                        </xsl:when>
                        <xsl:when test="($h5-paragraph-style-names = $paragraph-style-name or (contains($h5-paragraph-style-names, concat('»', $paragraph-style-name, '«'))))">
                            <xsl:value-of select="5"/>
                        </xsl:when>
                        <xsl:when test="($h6-paragraph-style-names = $paragraph-style-name or (contains($h6-paragraph-style-names, concat('»', $paragraph-style-name, '«'))))">
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
                           <xsl:when test="$styles/w:style[@w:type='paragraph' and @w:styleId=$paragraph-style-id]/w:rPr/w:sz">
                                <xsl:value-of select="$styles/w:style[@w:type='paragraph' and @w:styleId=$paragraph-style-id]/w:rPr/w:sz/@w:val"/>
                            </xsl:when>
                            <xsl:when test="$styles/w:style[@w:type='paragraph' and @w:styleId=$paragraph-style-id]/w:rPr/w:szCs">
                                <xsl:value-of select="$styles/w:style[@w:type='paragraph' and @w:styleId=$paragraph-style-id]/w:rPr/w:szCs/@w:val"/>
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
            <xsl:variable name="paragraph-style-name">
                <xsl:call-template name="get-style-name">
                    <xsl:with-param name="style-id" select="w:pPr/w:pStyle/@w:val"/>
                    <xsl:with-param name="style-type" select="'paragraph'"/>
                </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="paragraph-style-attribute-value">
                <xsl:choose>
                    <xsl:when test="contains($paragraph-style-name, $heading-marker)">
                        <xsl:value-of select="substring-before($paragraph-style-name, $heading-marker)"/>
                    </xsl:when>
                    <xsl:otherwise>
                        <xsl:value-of select="$paragraph-style-name"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:variable>
            <xsl:variable name="translated-paragraph-style-attribute-value" select="translate(translate($paragraph-style-attribute-value, $uppercase, $lowercase), ' ', '-')"/>
            <xsl:attribute name="{$paragraph-style-attribute-name}">
                <xsl:value-of select="$translated-paragraph-style-attribute-value"/>
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
            <xsl:otherwise>
                <xsl:attribute name="{$indexmark-type-attribute-name}">
                    <xsl:value-of select="''"/>
                </xsl:attribute>
            </xsl:otherwise>
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
            <xsl:otherwise>
                <xsl:attribute name="{$indexmark-format-attribute-name}">
                    <xsl:value-of select="''"/>
                </xsl:attribute>
            </xsl:otherwise>
        </xsl:choose>
        <!-- Entry (e.g. »Animal:Cat« for entry with subentry) -->
        <xsl:attribute name="{$indexmark-entry-attribute-name}">
            <xsl:value-of select="normalize-space(substring-before(substring-after($complex-field-content,'&quot;'), '&quot;'))"/>
        </xsl:attribute>
        <!-- Target (page range or see here entry, e.g. name of the textmark/bookmark or see here entry) -->
        <xsl:attribute name="{$indexmark-target-attribute-name}">
            <xsl:value-of select="normalize-space(substring-before(substring-after(substring-after(substring-after($complex-field-content,'&quot;'), '&quot;'), '&quot;'), '&quot;'))"/>
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
                        <xsl:variable name="character-style-name">
                            <xsl:call-template name="get-style-name">
                                <xsl:with-param name="style-id" select="$target-style-element/@w:val"/>
                                <xsl:with-param name="style-type" select="'character'"/>
                            </xsl:call-template>
                        </xsl:variable>
                        <xsl:variable name="translated-character-style-name" select="translate($character-style-name, ' ', '-')"/>
                        <xsl:call-template name="assign-inline-styles">
                            <xsl:with-param name="inline-style-elements" select="$inline-style-elements[position() != 1]"/>
                            <xsl:with-param name="class-names" select="concat($translated-character-style-name, ' ', $class-names)"/>
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
                        ($style-mode = 'extended') and 
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
    
    <!-- Drawing (includes Graphics, ...) -->
    <xsl:template match="w:drawing">
        <xsl:apply-templates />
    </xsl:template>
    
    <!-- Anchored/Inline Graphic -->
    <xsl:template match="wp:anchor | wp:inline">
        <xsl:apply-templates>
            <!-- Alt-text for graphic -->
            <xsl:with-param name="graphic-description" select="wp:docPr/@descr"/>
        </xsl:apply-templates>
    </xsl:template>
    
    <!-- Graphic (includes Groups, Shapes, Images, Textboxes) -->
    <xsl:template match="a:graphic">
        <xsl:param name="graphic-description" select="''"/>
        <xsl:apply-templates>
            <!-- Description (alt text for image) -->
            <xsl:with-param name="graphic-description" select="$graphic-description"/>
        </xsl:apply-templates>
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
        <xsl:param name="graphic-description" select="''"/>
        <xsl:element name="{$image-tag-name}" namespace="{$ns}">
            <xsl:call-template name="insert-image-attributes">
                <xsl:with-param name="graphic-description" select="$graphic-description"/>
            </xsl:call-template>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    <!-- Attributes for Image -->
    <xsl:template name="insert-image-attributes">
        <xsl:param name="graphic-description" select="''"/>
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
        <xsl:variable name="pic-description">
            <xsl:choose>
                <xsl:when test="not($graphic-description = '')">
                    <xsl:value-of select="normalize-space($graphic-description)"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="normalize-space(pic:nvPicPr/pic:cNvPr/@descr)"/>
                </xsl:otherwise>
            </xsl:choose>  
        </xsl:variable>
        <xsl:if test="not($pic-description = '')">
            <xsl:attribute name="{$image-alt-attribute-name}">
                <xsl:value-of select="$pic-description"/>
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
            <xsl:apply-templates select="$target-element/w:t/text()"/>
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
    
    
    <!-- Section Break -->
    <xsl:template match="w:sectPr">
        <xsl:element name="{$section-break-tag-name}" namespace="{$ns}">
            <xsl:attribute name="{$section-break-type-attribute-name}">
                <xsl:value-of select="'section-break'"/>
            </xsl:attribute>
            <xsl:apply-templates />
        </xsl:element>
    </xsl:template>
    
    
    <!-- Position of Last Calculated Page Break (application specific) -->
    <xsl:template match="w:lastRenderedPageBreak">
        <xsl:apply-templates/>
    </xsl:template>
    
    
    
    <!-- +++++++++++++++++++++ -->
    <!-- + General templates + -->
    <!-- +++++++++++++++++++++ -->
    
    <!-- Value of Style Name (Paragraph Style, Character Style) -->
    <xsl:template name="get-style-name">
        <xsl:param name="style-id" select="''"/>
        <xsl:param name="style-type" select="''"/>
        <xsl:variable name="style-element" select="$styles/w:style[@w:type=$style-type and @w:styleId=$style-id]"/>
        <xsl:choose>
            <xsl:when test="$style-element">
                <xsl:choose>
                    <!-- Style aliases -->
                    <xsl:when test="$style-element/w:aliases">
                        <xsl:variable name="style-aliases" select="$style-element/w:aliases/@w:val"/>
                        <xsl:choose>
                            <xsl:when test="contains($style-aliases, ',')">
                                <xsl:value-of select="substring-before($style-aliases,',')"/>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:value-of select="$style-aliases"/>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:when>
                    <!-- Style name -->
                    <xsl:otherwise>
                        <xsl:value-of select="$style-element/w:name/@w:val"/>
                    </xsl:otherwise>
                </xsl:choose>
            </xsl:when>
            <!-- Style ID -->
            <xsl:otherwise>
                <xsl:value-of select="$style-id"/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
    
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
    
    
    
    
    <!-- 
        Math Expression 
        Office Math Markup Language (OMML) to Mathematical Markup Language (MathML)
    -->
    <xsl:template match="m:oMath">
        <xsl:choose>
            <xsl:when test="$math-mode ='mml'">
                <xsl:element name="math">
                    <xsl:apply-templates select="*" />
                </xsl:element>
            </xsl:when>
            <xsl:when test="$math-mode ='text'">
                <xsl:element name="{$equation-tag-name}" namespace="{$ns}">
                    <xsl:call-template name="insert-equation-attributes"/>
                    <xsl:value-of select="normalize-space(.)"/>
                </xsl:element>
            </xsl:when>
        </xsl:choose>
    </xsl:template>
 
 
    <!-- Attributes for Math Expression -->
    <xsl:template name="insert-equation-attributes">
        <!-- Style -->
        <xsl:attribute name="{$equation-style-attribute-name}">
            <xsl:value-of select="$equation-style-attribute-value"/>
        </xsl:attribute>
    </xsl:template>
    
    
    <!--
        OMML to MML
        This part of stylesheet is derived from the TEI XSL stylesheet: 
        https://github.com/TEIC/Stylesheets/blob/dev/docx/from/omml2mml.xsl
    -->
    
    <!-- Global Definitions -->

    <!-- Every single unicode character that is recognized by OMML as an operator -->
    <xsl:variable name="sOperators"
  		select="concat(
            '&#x00A8;&#x0021;&#x0022;&#x0023;&#x0026;&#x0028;&#x0029;&#x002B;&#x002C;&#x002D;&#x002E;&#x002F;&#x003A;',
            '&#x003B;&#x003C;&#x003D;&#x003E;&#x003F;&#x0040;&#x005B;&#x005C;&#x005D;&#x005E;&#x005F;&#x0060;&#x007B;',
            '&#x007C;&#x007D;&#x007E;&#x00A1;&#x00A6;&#x00AC;&#x00AF;&#x00B0;&#x00B1;&#x00B2;&#x00B3;&#x00B4;&#x00B7;&#x00B9;&#x00BF;',
            '&#x00D7;&#x007E;&#x00F7;&#x02C7;&#x02D8;&#x02D9;&#x02DC;&#x02DD;&#x0300;&#x0301;&#x0302;&#x0303;&#x0304;&#x0305;&#x0306;&#x0307;&#x0308;&#x0309;',
            '&#x030A;&#x030B;&#x030C;&#x030D;&#x030E;&#x030F;&#x0310;&#x0311;&#x0312;&#x0313;&#x0314;&#x0315;',
            '&#x0316;&#x0317;&#x0318;&#x0319;&#x031A;&#x031B;&#x031C;&#x031D;&#x031E;&#x031F;&#x0320;&#x0321;',
            '&#x0322;&#x0323;&#x0324;&#x0325;&#x0326;&#x0327;&#x0328;&#x0329;&#x032A;&#x032B;&#x032C;&#x032D;',
            '&#x032E;&#x032F;&#x0330;&#x0331;&#x0332;&#x0333;&#x0334;&#x0335;&#x0336;&#x0337;&#x0338;&#x033F;',
            '&#x2000;&#x2001;&#x2002;&#x2003;&#x2004;&#x2005;&#x2006;&#x2009;&#x200A;&#x2010;&#x2012;&#x2013;',
            '&#x2014;&#x2016;&#x2020;&#x2021;&#x2022;&#x2024;&#x2025;&#x2026;&#x2032;&#x2033;&#x2034;&#x203C;',
            '&#x2040;&#x2044;&#x204E;&#x204F;&#x2050;&#x2057;&#x2061;&#x2062;&#x2063;&#x2070;&#x2074;&#x2075;',
            '&#x2076;&#x2077;&#x2078;&#x2079;&#x207A;&#x207B;&#x207C;&#x207D;&#x207E;&#x2080;&#x2081;&#x2082;',
            '&#x2083;&#x2084;&#x2085;&#x2086;&#x2087;&#x2088;&#x2089;&#x208A;&#x208B;&#x208C;&#x208D;&#x208E;',
            '&#x20D0;&#x20D1;&#x20D2;&#x20D3;&#x20D4;&#x20D5;&#x20D6;&#x20D7;&#x20D8;&#x20D9;&#x20DA;&#x20DB;',
            '&#x20DC;&#x20DD;&#x20DE;&#x20DF;&#x20E0;&#x20E1;&#x20E4;&#x20E5;&#x20E6;&#x20E7;&#x20E8;&#x20E9;',
            '&#x20EA;&#x2140;&#x2146;&#x2190;&#x2191;&#x2192;&#x2193;&#x2194;&#x2195;&#x2196;&#x2197;&#x2198;&#x2199;',
            '&#x219A;&#x219B;&#x219C;&#x219D;&#x219E;&#x219F;&#x21A0;&#x21A1;&#x21A2;&#x21A3;&#x21A4;&#x21A5;',
            '&#x21A6;&#x21A7;&#x21A8;&#x21A9;&#x21AA;&#x21AB;&#x21AC;&#x21AD;&#x21AE;&#x21AF;&#x21B0;&#x21B1;',
            '&#x21B2;&#x21B3;&#x21B6;&#x21B7;&#x21BA;&#x21BB;&#x21BC;&#x21BD;&#x21BE;&#x21BF;&#x21C0;&#x21C1;',
            '&#x21C2;&#x21C3;&#x21C4;&#x21C5;&#x21C6;&#x21C7;&#x21C8;&#x21C9;&#x21CA;&#x21CB;&#x21CC;&#x21CD;',
            '&#x21CE;&#x21CF;&#x21D0;&#x21D1;&#x21D2;&#x21D3;&#x21D4;&#x21D5;&#x21D6;&#x21D7;&#x21D8;&#x21D9;',
            '&#x21DA;&#x21DB;&#x21DC;&#x21DD;&#x21DE;&#x21DF;&#x21E0;&#x21E1;&#x21E2;&#x21E3;&#x21E4;&#x21E5;',
            '&#x21E6;&#x21E7;&#x21E8;&#x21E9;&#x21F3;&#x21F4;&#x21F5;&#x21F6;&#x21F7;&#x21F8;&#x21F9;&#x21FA;',
            '&#x21FB;&#x21FC;&#x21FD;&#x21FE;&#x21FF;&#x2200;&#x2201;&#x2202;&#x2203;&#x2204;&#x2206;&#x2207;',
            '&#x2208;&#x2209;&#x220A;&#x220B;&#x220C;&#x220D;&#x220F;&#x2210;&#x2211;&#x2212;&#x2213;&#x2214;',
            '&#x2215;&#x2216;&#x2217;&#x2218;&#x2219;&#x221A;&#x221B;&#x221C;&#x221D;&#x2223;&#x2224;&#x2225;',
            '&#x2226;&#x2227;&#x2228;&#x2229;&#x222A;&#x222B;&#x222C;&#x222D;&#x222E;&#x222F;&#x2230;&#x2231;',
            '&#x2232;&#x2233;&#x2234;&#x2235;&#x2236;&#x2237;&#x2238;&#x2239;&#x223A;&#x223B;&#x223C;&#x223D;',
            '&#x223E;&#x2240;&#x2241;&#x2242;&#x2243;&#x2244;&#x2245;&#x2246;&#x2247;&#x2248;&#x2249;&#x224A;',
            '&#x224B;&#x224C;&#x224D;&#x224E;&#x224F;&#x2250;&#x2251;&#x2252;&#x2253;&#x2254;&#x2255;&#x2256;',
            '&#x2257;&#x2258;&#x2259;&#x225A;&#x225B;&#x225C;&#x225D;&#x225E;&#x225F;&#x2260;&#x2261;&#x2262;',
            '&#x2263;&#x2264;&#x2265;&#x2266;&#x2267;&#x2268;&#x2269;&#x226A;&#x226B;&#x226C;&#x226D;&#x226E;',
            '&#x226F;&#x2270;&#x2271;&#x2272;&#x2273;&#x2274;&#x2275;&#x2276;&#x2277;&#x2278;&#x2279;&#x227A;',
            '&#x227B;&#x227C;&#x227D;&#x227E;&#x227F;&#x2280;&#x2281;&#x2282;&#x2283;&#x2284;&#x2285;&#x2286;',
            '&#x2287;&#x2288;&#x2289;&#x228A;&#x228B;&#x228C;&#x228D;&#x228E;&#x228F;&#x2290;&#x2291;&#x2292;',
            '&#x2293;&#x2294;&#x2295;&#x2296;&#x2297;&#x2298;&#x2299;&#x229A;&#x229B;&#x229C;&#x229D;&#x229E;',
            '&#x229F;&#x22A0;&#x22A1;&#x22A2;&#x22A3;&#x22A5;&#x22A6;&#x22A7;&#x22A8;&#x22A9;&#x22AA;&#x22AB;',
            '&#x22AC;&#x22AD;&#x22AE;&#x22AF;&#x22B0;&#x22B1;&#x22B2;&#x22B3;&#x22B4;&#x22B5;&#x22B6;&#x22B7;',
            '&#x22B8;&#x22B9;&#x22BA;&#x22BB;&#x22BC;&#x22BD;&#x22C0;&#x22C1;&#x22C2;&#x22C3;&#x22C4;&#x22C5;',
            '&#x22C6;&#x22C7;&#x22C8;&#x22C9;&#x22CA;&#x22CB;&#x22CC;&#x22CD;&#x22CE;&#x22CF;&#x22D0;&#x22D1;',
            '&#x22D2;&#x22D3;&#x22D4;&#x22D5;&#x22D6;&#x22D7;&#x22D8;&#x22D9;&#x22DA;&#x22DB;&#x22DC;&#x22DD;',
            '&#x22DE;&#x22DF;&#x22E0;&#x22E1;&#x22E2;&#x22E3;&#x22E4;&#x22E5;&#x22E6;&#x22E7;&#x22E8;&#x22E9;',
            '&#x22EA;&#x22EB;&#x22EC;&#x22ED;&#x22EE;&#x22EF;&#x22F0;&#x22F1;&#x22F2;&#x22F3;&#x22F4;&#x22F5;',
            '&#x22F6;&#x22F7;&#x22F8;&#x22F9;&#x22FA;&#x22FB;&#x22FC;&#x22FD;&#x22FE;&#x22FF;&#x2305;&#x2306;',
            '&#x2308;&#x2309;&#x230A;&#x230B;&#x231C;&#x231D;&#x231E;&#x231F;&#x2322;&#x2323;&#x2329;&#x232A;',
            '&#x233D;&#x233F;&#x23B0;&#x23B1;&#x23DC;&#x23DD;&#x23DE;&#x23DF;&#x23E0;&#x2502;&#x251C;&#x2524;',
            '&#x252C;&#x2534;&#x2581;&#x2588;&#x2592;&#x25A0;&#x25A1;&#x25AD;&#x25B2;&#x25B3;&#x25B4;&#x25B5;',
            '&#x25B6;&#x25B7;&#x25B8;&#x25B9;&#x25BC;&#x25BD;&#x25BE;&#x25BF;&#x25C0;&#x25C1;&#x25C2;&#x25C3;',
            '&#x25C4;&#x25C5;&#x25CA;&#x25CB;&#x25E6;&#x25EB;&#x25EC;&#x25F8;&#x25F9;&#x25FA;&#x25FB;&#x25FC;',
            '&#x25FD;&#x25FE;&#x25FF;&#x2605;&#x2606;&#x2772;&#x2773;&#x27D1;&#x27D2;&#x27D3;&#x27D4;&#x27D5;',
            '&#x27D6;&#x27D7;&#x27D8;&#x27D9;&#x27DA;&#x27DB;&#x27DC;&#x27DD;&#x27DE;&#x27DF;&#x27E0;&#x27E1;',
            '&#x27E2;&#x27E3;&#x27E4;&#x27E5;&#x27E6;&#x27E7;&#x27E8;&#x27E9;&#x27EA;&#x27EB;&#x27F0;&#x27F1;',
            '&#x27F2;&#x27F3;&#x27F4;&#x27F5;&#x27F6;&#x27F7;&#x27F8;&#x27F9;&#x27FA;&#x27FB;&#x27FC;&#x27FD;',
            '&#x27FE;&#x27FF;&#x2900;&#x2901;&#x2902;&#x2903;&#x2904;&#x2905;&#x2906;&#x2907;&#x2908;&#x2909;',
            '&#x290A;&#x290B;&#x290C;&#x290D;&#x290E;&#x290F;&#x2910;&#x2911;&#x2912;&#x2913;&#x2914;&#x2915;',
            '&#x2916;&#x2917;&#x2918;&#x2919;&#x291A;&#x291B;&#x291C;&#x291D;&#x291E;&#x291F;&#x2920;&#x2921;',
            '&#x2922;&#x2923;&#x2924;&#x2925;&#x2926;&#x2927;&#x2928;&#x2929;&#x292A;&#x292B;&#x292C;&#x292D;',
            '&#x292E;&#x292F;&#x2930;&#x2931;&#x2932;&#x2933;&#x2934;&#x2935;&#x2936;&#x2937;&#x2938;&#x2939;',
            '&#x293A;&#x293B;&#x293C;&#x293D;&#x293E;&#x293F;&#x2940;&#x2941;&#x2942;&#x2943;&#x2944;&#x2945;',
            '&#x2946;&#x2947;&#x2948;&#x2949;&#x294A;&#x294B;&#x294C;&#x294D;&#x294E;&#x294F;&#x2950;&#x2951;',
            '&#x2952;&#x2953;&#x2954;&#x2955;&#x2956;&#x2957;&#x2958;&#x2959;&#x295A;&#x295B;&#x295C;&#x295D;',
            '&#x295E;&#x295F;&#x2960;&#x2961;&#x2962;&#x2963;&#x2964;&#x2965;&#x2966;&#x2967;&#x2968;&#x2969;',
            '&#x296A;&#x296B;&#x296C;&#x296D;&#x296E;&#x296F;&#x2970;&#x2971;&#x2972;&#x2973;&#x2974;&#x2975;',
            '&#x2976;&#x2977;&#x2978;&#x2979;&#x297A;&#x297B;&#x297C;&#x297D;&#x297E;&#x297F;&#x2980;&#x2982;',
            '&#x2983;&#x2984;&#x2985;&#x2986;&#x2987;&#x2988;&#x2989;&#x298A;&#x298B;&#x298C;&#x298D;&#x298E;',
            '&#x298F;&#x2990;&#x2991;&#x2992;&#x2993;&#x2994;&#x2995;&#x2996;&#x2997;&#x2998;&#x2999;&#x299A;',
            '&#x29B6;&#x29B7;&#x29B8;&#x29B9;&#x29C0;&#x29C1;&#x29C4;&#x29C5;&#x29C6;&#x29C7;&#x29C8;&#x29CE;',
            '&#x29CF;&#x29D0;&#x29D1;&#x29D2;&#x29D3;&#x29D4;&#x29D5;&#x29D6;&#x29D7;&#x29D8;&#x29D9;&#x29DA;',
            '&#x29DB;&#x29DF;&#x29E1;&#x29E2;&#x29E3;&#x29E4;&#x29E5;&#x29E6;&#x29EB;&#x29F4;&#x29F5;&#x29F6;',
            '&#x29F7;&#x29F8;&#x29F9;&#x29FA;&#x29FB;&#x29FC;&#x29FD;&#x29FE;&#x29FF;&#x2A00;&#x2A01;&#x2A02;',
            '&#x2A03;&#x2A04;&#x2A05;&#x2A06;&#x2A07;&#x2A08;&#x2A09;&#x2A0A;&#x2A0B;&#x2A0C;&#x2A0D;&#x2A0E;',
            '&#x2A0F;&#x2A10;&#x2A11;&#x2A12;&#x2A13;&#x2A14;&#x2A15;&#x2A16;&#x2A17;&#x2A18;&#x2A19;&#x2A1A;',
            '&#x2A1B;&#x2A1C;&#x2A1D;&#x2A1E;&#x2A1F;&#x2A20;&#x2A21;&#x2A22;&#x2A23;&#x2A24;&#x2A25;&#x2A26;',
            '&#x2A27;&#x2A28;&#x2A29;&#x2A2A;&#x2A2B;&#x2A2C;&#x2A2D;&#x2A2E;&#x2A2F;&#x2A30;&#x2A31;&#x2A32;',
            '&#x2A33;&#x2A34;&#x2A35;&#x2A36;&#x2A37;&#x2A38;&#x2A39;&#x2A3A;&#x2A3B;&#x2A3C;&#x2A3D;&#x2A3E;',
            '&#x2A3F;&#x2A40;&#x2A41;&#x2A42;&#x2A43;&#x2A44;&#x2A45;&#x2A46;&#x2A47;&#x2A48;&#x2A49;&#x2A4A;',
            '&#x2A4B;&#x2A4C;&#x2A4D;&#x2A4E;&#x2A4F;&#x2A50;&#x2A51;&#x2A52;&#x2A53;&#x2A54;&#x2A55;&#x2A56;',
            '&#x2A57;&#x2A58;&#x2A59;&#x2A5A;&#x2A5B;&#x2A5C;&#x2A5D;&#x2A5E;&#x2A5F;&#x2A60;&#x2A61;&#x2A62;',
            '&#x2A63;&#x2A64;&#x2A65;&#x2A66;&#x2A67;&#x2A68;&#x2A69;&#x2A6A;&#x2A6B;&#x2A6C;&#x2A6D;&#x2A6E;',
            '&#x2A6F;&#x2A70;&#x2A71;&#x2A72;&#x2A73;&#x2A74;&#x2A75;&#x2A76;&#x2A77;&#x2A78;&#x2A79;&#x2A7A;',
            '&#x2A7B;&#x2A7C;&#x2A7D;&#x2A7E;&#x2A7F;&#x2A80;&#x2A81;&#x2A82;&#x2A83;&#x2A84;&#x2A85;&#x2A86;',
            '&#x2A87;&#x2A88;&#x2A89;&#x2A8A;&#x2A8B;&#x2A8C;&#x2A8D;&#x2A8E;&#x2A8F;&#x2A90;&#x2A91;&#x2A92;',
            '&#x2A93;&#x2A94;&#x2A95;&#x2A96;&#x2A97;&#x2A98;&#x2A99;&#x2A9A;&#x2A9B;&#x2A9C;&#x2A9D;&#x2A9E;',
            '&#x2A9F;&#x2AA0;&#x2AA1;&#x2AA2;&#x2AA3;&#x2AA4;&#x2AA5;&#x2AA6;&#x2AA7;&#x2AA8;&#x2AA9;&#x2AAA;',
            '&#x2AAB;&#x2AAC;&#x2AAD;&#x2AAE;&#x2AAF;&#x2AB0;&#x2AB1;&#x2AB2;&#x2AB3;&#x2AB4;&#x2AB5;&#x2AB6;',
            '&#x2AB7;&#x2AB8;&#x2AB9;&#x2ABA;&#x2ABB;&#x2ABC;&#x2ABD;&#x2ABE;&#x2ABF;&#x2AC0;&#x2AC1;&#x2AC2;',
            '&#x2AC3;&#x2AC4;&#x2AC5;&#x2AC6;&#x2AC7;&#x2AC8;&#x2AC9;&#x2ACA;&#x2ACB;&#x2ACC;&#x2ACD;&#x2ACE;',
            '&#x2ACF;&#x2AD0;&#x2AD1;&#x2AD2;&#x2AD3;&#x2AD4;&#x2AD5;&#x2AD6;&#x2AD7;&#x2AD8;&#x2AD9;&#x2ADA;',
            '&#x2ADB;&#x2ADC;&#x2ADD;&#x2ADE;&#x2ADF;&#x2AE0;&#x2AE2;&#x2AE3;&#x2AE4;&#x2AE5;&#x2AE6;&#x2AE7;',
            '&#x2AE8;&#x2AE9;&#x2AEA;&#x2AEB;&#x2AEC;&#x2AED;&#x2AEE;&#x2AEF;&#x2AF0;&#x2AF2;&#x2AF3;&#x2AF4;',
            '&#x2AF5;&#x2AF6;&#x2AF7;&#x2AF8;&#x2AF9;&#x2AFA;&#x2AFB;&#x2AFC;&#x2AFD;&#x2AFE;&#x2AFF;&#x2B04;',
            '&#x2B06;&#x2B07;&#x2B0C;&#x2B0D;&#x3014;&#x3015;&#x3016;&#x3017;&#x3018;&#x3019;&#xFF01;&#xFF06;',
            '&#xFF08;&#xFF09;&#xFF0B;&#xFF0C;&#xFF0D;&#xFF0E;&#xFF0F;&#xFF1A;&#xFF1B;&#xFF1C;&#xFF1D;&#xFF1E;',
            '&#xFF1F;&#xFF20;&#xFF3B;&#xFF3C;&#xFF3D;&#xFF3E;&#xFF3F;&#xFF5B;&#xFF5C;&#xFF5D;')" />
  
    <!-- A string of '-'s repeated exactly as many times as the operators above -->
    <xsl:variable name="sMinuses">
      <xsl:call-template name="SRepeatChar">
        <xsl:with-param name="cchRequired" select="string-length($sOperators)" />
        <xsl:with-param name="ch" select="'-'" />
      </xsl:call-template>
    </xsl:variable>
  
    <!-- Every single unicode character that is recognized by OMML as a number -->
    <xsl:variable name="sNumbers" select="'0123456789'" />
  
    <!-- A string of '0's repeated exactly as many times as the list of numbers above -->
    <xsl:variable name="sZeros">
      <xsl:call-template name="SRepeatChar">
        <xsl:with-param name="cchRequired" select="string-length($sNumbers)" />
        <xsl:with-param name="ch" select="'0'" />
      </xsl:call-template>
    </xsl:variable>
  
    <!-- 
      Template: SReplace
      
      Replace all occurences of sOrig in sInput with sReplacement
      and return the resulting string. 
    -->
    <xsl:template name="SReplace">
      
      <xsl:param name="sInput" />
      <xsl:param name="sOrig" />
      <xsl:param name="sReplacement" />
  
      <xsl:choose>
        <xsl:when test="not(contains($sInput, $sOrig))">
          <xsl:value-of select="$sInput" />
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="sBefore" select="substring-before($sInput, $sOrig)" />
          <xsl:variable name="sAfter" select="substring-after($sInput, $sOrig)" />
          <xsl:variable name="sAfterProcessed">
            <xsl:call-template name="SReplace">
              <xsl:with-param name="sInput" select="$sAfter" />
              <xsl:with-param name="sOrig" select="$sOrig" />
              <xsl:with-param name="sReplacement" select="$sReplacement" />
            </xsl:call-template>
          </xsl:variable>
          <xsl:value-of select="concat($sBefore, concat($sReplacement, $sAfterProcessed))" />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!-- Templates -->
  
    <xsl:template match="m:borderBox">
  
      <!-- Get Lowercase versions of properties -->
      <xsl:variable name="sLowerCaseHideTop" select="translate(m:borderBoxPr[last()]/m:hideTop[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseHideBot" select="translate(m:borderBoxPr[last()]/m:hideBot[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseHideLeft" select="translate(m:borderBoxPr[last()]/m:hideLeft[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseHideRight" select="translate(m:borderBoxPr[last()]/m:hideRight[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseStrikeH" select="translate(m:borderBoxPr[last()]/m:strikeH[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseStrikeV" select="translate(m:borderBoxPr[last()]/m:strikeV[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseStrikeBLTR" select="translate(m:borderBoxPr[last()]/m:strikeBLTR[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseStrikeTLBR" select="translate(m:borderBoxPr[last()]/m:strikeTLBR[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="fHideTop">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseHideTop" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fHideBot">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseHideBot" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fHideLeft">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseHideLeft" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fHideRight">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseHideRight" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fStrikeH">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseStrikeH" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fStrikeV">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseStrikeV" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fStrikeBLTR">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseStrikeBLTR" />
        </xsl:call-template>
      </xsl:variable>
      <xsl:variable name="fStrikeTLBR">
        <xsl:call-template name="ForceTrueStrVal">
          <xsl:with-param name="str" select="$sLowerCaseStrikeTLBR" />
        </xsl:call-template>
      </xsl:variable>
  
      <xsl:choose>
        <xsl:when test="$fHideTop=1 
                        and $fHideBot=1 
                        and $fHideLeft=1 
                        and $fHideRight=1 
                        and $fStrikeH=0 
                        and $fStrikeV=0 
                        and $fStrikeBLTR=0 
                        and $fStrikeTLBR=0">
          <mrow>
            <xsl:apply-templates select="m:e[1]" />
          </mrow>
        </xsl:when>
        <xsl:otherwise>
          <menclose>
            <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
              <xsl:with-param name="fHideTop" select="$fHideTop" />
              <xsl:with-param name="fHideBot" select="$fHideBot" />
              <xsl:with-param name="fHideLeft" select="$fHideLeft" />
              <xsl:with-param name="fHideRight" select="$fHideRight" />
              <xsl:with-param name="fStrikeH" select="$fStrikeH" />
              <xsl:with-param name="fStrikeV" select="$fStrikeV" />
              <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
              <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
            </xsl:call-template>
            <xsl:apply-templates select="m:e[1]" />
          </menclose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!--<xsl:template match="*">
    <xsl:apply-templates select="*" />
    </xsl:template>-->
  
    <!--
      { Non-combining, Upper-combining, Lower-combining }
      {U+02D8, U+0306, U+032E}, // BREVE
      {U+00B8, U+0312, U+0327}, // CEDILLA
      {U+0060, U+0300, U+0316}, // GRAVE ACCENT
      {U+002D, U+0305, U+0332}, // HYPHEN-MINUS/OVERLINE
      {U+2212, U+0305, U+0332}, // MINUS SIGN/OVERLINE
      {U+002E, U+0305, U+0323}, // FULL STOP/DOT ABOVE
      {U+02D9, U+0307, U+0323}, // DOT ABOVE
      {U+02DD, U+030B, U+02DD}, // DOUBLE ACUTE ACCENT
      {U+00B4, U+0301, U+0317}, // ACUTE ACCENT
      {U+007E, U+0303, U+0330}, // TILDE
      {U+02DC, U+0303, U+0330}, // SMALL TILDE
      {U+00A8, U+0308, U+0324}, // DIAERESIS
      {U+02C7, U+030C, U+032C}, // CARON
      {U+005E, U+0302, U+032D}, // CIRCUMFLEX ACCENT
      {U+00AF, U+0305, ::::::}, // MACRON
      {U+005F, ::::::, U+0332}, // LOW LINE
      {U+2192, U+20D7, U+20EF}, // RIGHTWARDS ARROW
      {U+27F6, U+20D7, U+20EF}, // LONG RIGHTWARDS ARROW
      {U+2190, U+20D6, U+20EE}, // LEFT ARROW
    -->
    <xsl:template name="ToNonCombining">
      <xsl:param name="ch" />
      <xsl:choose>
        <!-- BREVE -->
        <xsl:when test="$ch='&#x0306;' or $ch='&#x032e;'">&#x02D8;</xsl:when>
        <!-- CEDILLA -->
        <xsl:when test="$ch='&#x0312;' or $ch='&#x0327;'">&#x00B8;</xsl:when>
        <!-- GRAVE ACCENT -->
        <xsl:when test="$ch='&#x0300;' or $ch='&#x0316;'">&#x0060;</xsl:when>
        <!-- HYPHEN-MINUS/OVERLINE -->
        <xsl:when test="$ch='&#x0305;' or $ch='&#x0332;'">&#x002D;</xsl:when>
        <!-- MINUS SIGN/OVERLINE -->
        <xsl:when test="$ch='&#x0305;' or $ch='&#x0332;'">&#x2212;</xsl:when>
        <!-- FULL STOP/DOT ABOVE -->
        <xsl:when test="$ch='&#x0305;' or $ch='&#x0323;'">&#x002E;</xsl:when>
        <!-- DOT ABOVE -->
        <xsl:when test="$ch='&#x0307;' or $ch='&#x0323;'">&#x02D9;</xsl:when>
        <!-- DOUBLE ACUTE ACCENT -->
        <xsl:when test="$ch='&#x030B;' or $ch='&#x02DD;'">&#x02DD;</xsl:when>
        <!-- ACUTE ACCENT -->
        <xsl:when test="$ch='&#x0301;' or $ch='&#x0317;'">&#x00B4;</xsl:when>
        <!-- TILDE -->
        <xsl:when test="$ch='&#x0303;' or $ch='&#x0330;'">&#x007E;</xsl:when>
        <!-- SMALL TILDE -->
        <xsl:when test="$ch='&#x0303;' or $ch='&#x0330;'">&#x02DC;</xsl:when>
        <!-- DIAERESIS -->
        <xsl:when test="$ch='&#x0308;' or $ch='&#x0324;'">&#x00A8;</xsl:when>
        <!-- CARON -->
        <xsl:when test="$ch='&#x030C;' or $ch='&#x032C;'">&#x02C7;</xsl:when>
        <!-- CIRCUMFLEX ACCENT -->
        <xsl:when test="$ch='&#x0302;' or $ch='&#x032D;'">&#x005E;</xsl:when>
        <!-- MACRON -->
        <xsl:when test="$ch='&#x0305;'                   ">&#x00AF;</xsl:when>
        <!-- LOW LINE -->
        <xsl:when test="                   $ch='&#x0332;'">&#x005F;</xsl:when>
        <!-- RIGHTWARDS ARROW -->
        <xsl:when test="$ch='&#x20D7;' or $ch='&#x20EF;'">&#x2192;</xsl:when>
        <!-- LONG RIGHTWARDS ARROW -->
        <xsl:when test="$ch='&#x20D7;' or $ch='&#x20EF;'">&#x27F6;</xsl:when>
        <!-- LEFT ARROW -->
        <xsl:when test="$ch='&#x20D6;' or $ch='&#x20EE;'">&#x2190;</xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$ch"/>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template match="m:acc">
      <mover>
        <xsl:attribute name="accent">true</xsl:attribute>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <xsl:variable name="chAcc">
          <xsl:choose>
            <xsl:when test="not(m:accPr[last()]/m:chr)">
              <xsl:value-of select="'&#x0302;'" />
            </xsl:when>
            <xsl:otherwise>
              <xsl:value-of select="substring(m:accPr/m:chr/@m:val,1,1)" />
            </xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="chNonComb">
          <xsl:call-template name="ToNonCombining">
            <xsl:with-param name="ch" select="$chAcc" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:choose>
          <xsl:when test="string-length($chAcc)=0">
            <mo/>
          </xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ParseMt">
              <xsl:with-param name="sToParse" select="$chNonComb" />
              <xsl:with-param name="scr" select="m:e[1]/*/m:rPr[last()]/m:scr/@m:val" />
              <xsl:with-param name="sty" select="m:e[1]/*/m:rPr[last()]/m:sty/@m:val" />
              <xsl:with-param name="nor">
                <xsl:choose>
                  <xsl:when test="count(m:e[1]/*/m:rPr[last()]/m:nor) = 0">0</xsl:when>
                  <xsl:otherwise>
                    <xsl:call-template name="ForceFalseStrVal">
                      <xsl:with-param name="str" select="translate(m:e[1]/*/m:rPr[last()]/m:nor/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
                    </xsl:call-template>
                  </xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </mover>
    </xsl:template>
  
    <xsl:template name="OutputScript">
      <xsl:param name="ndCur" select="." />
      <xsl:choose>
        <!-- Only output contents of $ndCur if $ndCur exists and $ndCur has children -->
        <xsl:when test="count($ndCur/*) &gt; 0">
          <mrow>
            <xsl:apply-templates select="$ndCur" />
          </mrow>
        </xsl:when>
        <xsl:otherwise>
          <none />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template match="m:sPre">
      <mmultiscripts>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <mprescripts />
        <xsl:call-template name="OutputScript">
          <xsl:with-param name="ndCur" select="m:sub[1]"/>
        </xsl:call-template>
        <xsl:call-template name="OutputScript">
          <xsl:with-param name="ndCur" select="m:sup[1]" />
        </xsl:call-template>
      </mmultiscripts>
    </xsl:template>
  
    <xsl:template match="m:m">
      <mtable>
        <xsl:call-template name="CreateMathMLMatrixAttr">
          <xsl:with-param name="mcJc" select="m:mPr[last()]/m:mcs/m:mc/m:mcPr[last()]/m:mcJc/@m:val" />
        </xsl:call-template>
        <xsl:for-each select="m:mr">
          <mtr>
            <xsl:for-each select="m:e">
              <mtd>
                <xsl:apply-templates select="." />
              </mtd>
            </xsl:for-each>
          </mtr>
        </xsl:for-each>
      </mtable>
    </xsl:template>
  
    <xsl:template name="CreateMathMLMatrixAttr">
      <xsl:param name="mcJc" />
      <xsl:variable name="sLowerCaseMcjc" select="translate($mcJc, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:choose>
        <xsl:when test="$sLowerCaseMcjc='left'">
          <xsl:attribute name="columnalign">left</xsl:attribute>
        </xsl:when>
        <xsl:when test="$sLowerCaseMcjc='right'">
          <xsl:attribute name="columnalign">right</xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template match="m:phant">
      <xsl:variable name="sLowerCaseZeroWidVal" select="translate(m:phantPr[last()]/m:zeroWid[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseZeroAscVal" select="translate(m:phantPr[last()]/m:zeroAsc[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseZeroDescVal" select="translate(m:phantPr[last()]/m:zeroDesc[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseShowVal" select="translate(m:phantPr[last()]/m:show[last()]/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <!-- The following properties default to 'yes' unless the last value equals 'no' or there isn't any node for the property -->
      <xsl:variable name="fZeroWid">
        <xsl:choose>
          <xsl:when test="count(m:phantPr[last()]/m:zeroWid[last()]) = 0">0</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ForceFalseStrVal">
              <xsl:with-param name="str" select="$sLowerCaseZeroWidVal" />
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="fZeroAsc">
        <xsl:choose>
          <xsl:when test="count(m:phantPr[last()]/m:zeroAsc[last()]) = 0">0</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ForceFalseStrVal">
              <xsl:with-param name="str" select="$sLowerCaseZeroAscVal" />
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:variable name="fZeroDesc">
        <xsl:choose>
          <xsl:when test="count(m:phantPr[last()]/m:zeroDesc[last()]) = 0">0</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ForceFalseStrVal">
              <xsl:with-param name="str" select="$sLowerCaseZeroDescVal" />
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
  
      <!-- The show property defaults to 'on' unless there exists a show property and its value is 'off' -->
  
      <xsl:variable name="fShow">
        <xsl:call-template name="ForceFalseStrVal">
          <xsl:with-param name="str" select="$sLowerCaseShowVal" />
        </xsl:call-template>
      </xsl:variable>
  
      <xsl:choose>
        <!-- Show the phantom contents, therefore, just use mpadded. -->
        <xsl:when test="$fShow = 1">
          <xsl:element name="mpadded">
            <xsl:call-template name="CreateMpaddedAttributes">
              <xsl:with-param name="fZeroWid" select="$fZeroWid" />
              <xsl:with-param name="fZeroAsc" select="$fZeroAsc" />
              <xsl:with-param name="fZeroDesc" select="$fZeroDesc" />
            </xsl:call-template>
            <mrow>
              <xsl:apply-templates select="m:e" />
            </mrow>
          </xsl:element>
        </xsl:when>
        <!-- Don't show phantom contents, but don't smash anything, therefore, just use mphantom -->
        <xsl:when test="$fZeroWid=0 and $fZeroAsc=0 and $fZeroDesc=0">
          <xsl:element name="mphantom">
            <mrow>
              <xsl:apply-templates select="m:e" />
            </mrow>
          </xsl:element>
        </xsl:when>
        <!-- Combination -->
        <xsl:otherwise>
          <xsl:element name="mphantom">
            <xsl:element name="mpadded">
              <xsl:call-template name="CreateMpaddedAttributes">
                <xsl:with-param name="fZeroWid" select="$fZeroWid" />
                <xsl:with-param name="fZeroAsc" select="$fZeroAsc" />
                <xsl:with-param name="fZeroDesc" select="$fZeroDesc" />
              </xsl:call-template>
              <mrow>
                <xsl:apply-templates select="m:e" />
              </mrow>
            </xsl:element>
          </xsl:element>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template name="CreateMpaddedAttributes">
      <xsl:param name="fZeroWid" />
      <xsl:param name="fZeroAsc" />
      <xsl:param name="fZeroDesc" />
      <xsl:if test="$fZeroWid=1">
        <xsl:attribute name="width">0in</xsl:attribute>
      </xsl:if>
      <xsl:if test="$fZeroAsc=1">
        <xsl:attribute name="height">0in</xsl:attribute>
      </xsl:if>
      <xsl:if test="$fZeroDesc=1">
        <xsl:attribute name="depth">0in</xsl:attribute>
      </xsl:if>
    </xsl:template>

    <xsl:template match="m:rad">
      <xsl:variable name="fDegHide">
        <xsl:choose>
          <xsl:when test="count(m:radPr[last()]/m:degHide)=0">0</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ForceFalseStrVal">
              <xsl:with-param name="str" select="translate(m:radPr[last()]/m:degHide/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$fDegHide=1">
          <msqrt>
            <xsl:apply-templates select="m:e[1]" />
          </msqrt>
        </xsl:when>
        <xsl:otherwise>
          <mroot>
            <mrow>
              <xsl:apply-templates select="m:e[1]" />
            </mrow>
            <mrow>
              <xsl:apply-templates select="m:deg[1]" />
            </mrow>
          </mroot>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template name="OutputNaryMo">
      <xsl:param name="ndCur" select="." />
      <xsl:param name="fGrow" select="0" />
      <mo>
        <xsl:choose>
          <xsl:when test="$fGrow=1">
            <xsl:attribute name="stretchy">true</xsl:attribute>
          </xsl:when>
          <xsl:otherwise>
            <xsl:attribute name="stretchy">false</xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
        <xsl:choose>
          <xsl:when test="not($ndCur/m:naryPr[last()]/m:chr/@m:val) or
  			                            $ndCur/m:naryPr[last()]/m:chr/@m:val=''">
            <xsl:text>&#x222b;</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="$ndCur/m:naryPr[last()]/m:chr/@m:val" />
          </xsl:otherwise>
        </xsl:choose>
      </mo>
    </xsl:template>
  
    <!-- 
      Template match m:nary 
      Process an n-ary. 
      
      Decides, based on which arguments are supplied, between
      using an mo, msup, msub, or msubsup for the n-ary operator    
    -->
    <xsl:template match="m:nary">
      <xsl:variable name="sLowerCaseSubHide">
        <xsl:choose>
          <xsl:when test="count(m:naryPr[last()]/m:subHide) = 0">
            <xsl:text>off</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="translate(m:naryPr[last()]/m:subHide/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
  
      <xsl:variable name="sLowerCaseSupHide">
        <xsl:choose>
          <xsl:when test="count(m:naryPr[last()]/m:supHide) = 0">
            <xsl:text>off</xsl:text>
          </xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="translate(m:naryPr[last()]/m:supHide/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
  
      <xsl:variable name="sLowerCaseLimLoc">
        <xsl:value-of select="translate(m:naryPr[last()]/m:limLoc/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      </xsl:variable>
  
      <xsl:variable name="sLowerGrow">
        <xsl:choose>
          <xsl:when test="count(m:naryPr[last()]/m:grow)=0">off</xsl:when>
          <xsl:otherwise>
            <xsl:value-of select="translate(m:naryPr[last()]/m:grow/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
  
      <xsl:variable name="fLimLocSubSup">
        <xsl:choose>
          <xsl:when test="count(m:naryPr[last()]/m:limLoc)=0 or $sLowerCaseLimLoc='subsup'">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
  
      <xsl:variable name="fGrow">
        <xsl:call-template name="ForceFalseStrVal">
          <xsl:with-param name="str" select="$sLowerGrow" />
        </xsl:call-template>
      </xsl:variable>
  
      <xsl:variable name="fSupHide">
        <xsl:call-template name="ForceFalseStrVal">
          <xsl:with-param name="str" select="$sLowerCaseSupHide" />
        </xsl:call-template>
      </xsl:variable>
  
      <xsl:variable name="fSubHide">
        <xsl:call-template name="ForceFalseStrVal">
          <xsl:with-param name="str" select="$sLowerCaseSubHide" />
        </xsl:call-template>
      </xsl:variable>
  
      <mrow>
        <xsl:choose>
          <xsl:when test="$fSupHide=1 and $fSubHide=1">
            <xsl:call-template name="OutputNaryMo">
              <xsl:with-param name="ndCur" select="." />
              <xsl:with-param name="fGrow" select="$fGrow" />
            </xsl:call-template>
          </xsl:when>
          <xsl:when test="$fSubHide=1">
            <xsl:choose>
              <xsl:when test="$fLimLocSubSup=1">
                <msup>
                  <xsl:call-template name="OutputNaryMo">
                    <xsl:with-param name="ndCur" select="." />
                    <xsl:with-param name="fGrow" select="$fGrow" />
                  </xsl:call-template>
                  <mrow>
                    <xsl:apply-templates select="m:sup[1]" />
                  </mrow>
                </msup>
              </xsl:when>
              <xsl:otherwise>
                <mover>
                  <xsl:call-template name="OutputNaryMo">
                    <xsl:with-param name="ndCur" select="." />
                    <xsl:with-param name="fGrow" select="$fGrow" />
                  </xsl:call-template>
                  <mrow>
                    <xsl:apply-templates select="m:sup[1]" />
                  </mrow>
                </mover>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:when test="$fSupHide=1">
            <xsl:choose>
              <xsl:when test="$fLimLocSubSup=1">
                <msub>
                  <xsl:call-template name="OutputNaryMo">
                    <xsl:with-param name="ndCur" select="." />
                    <xsl:with-param name="fGrow" select="$fGrow" />
                  </xsl:call-template>
                  <mrow>
                    <xsl:apply-templates select="m:sub[1]" />
                  </mrow>
                </msub>
              </xsl:when>
              <xsl:otherwise>
                <munder>
                  <xsl:call-template name="OutputNaryMo">
                    <xsl:with-param name="ndCur" select="." />
                    <xsl:with-param name="fGrow" select="$fGrow" />
                  </xsl:call-template>
                  <mrow>
                    <xsl:apply-templates select="m:sub[1]" />
                  </mrow>
                </munder>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:when>
          <xsl:otherwise>
            <xsl:choose>
              <xsl:when test="$fLimLocSubSup=1">
                <msubsup>
                  <xsl:call-template name="OutputNaryMo">
                    <xsl:with-param name="ndCur" select="." />
                    <xsl:with-param name="fGrow" select="$fGrow" />
                  </xsl:call-template>
                  <mrow>
                    <xsl:apply-templates select="m:sub[1]" />
                  </mrow>
                  <mrow>
                    <xsl:apply-templates select="m:sup[1]" />
                  </mrow>
                </msubsup>
              </xsl:when>
              <xsl:otherwise>
                <munderover>
                  <xsl:call-template name="OutputNaryMo">
                    <xsl:with-param name="ndCur" select="." />
                    <xsl:with-param name="fGrow" select="$fGrow" />
                  </xsl:call-template>
                  <mrow>
                    <xsl:apply-templates select="m:sub[1]" />
                  </mrow>
                  <mrow>
                    <xsl:apply-templates select="m:sup[1]" />
                  </mrow>
                </munderover>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
      </mrow>
    </xsl:template>
  
    <xsl:template match="m:limLow">
      <munder>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <mrow>
          <xsl:apply-templates select="m:lim[1]" />
        </mrow>
      </munder>
    </xsl:template>
  
    <xsl:template match="m:limUpp">
      <mover>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <mrow>
          <xsl:apply-templates select="m:lim[1]" />
        </mrow>
      </mover>
    </xsl:template>
  
    <xsl:template match="m:sSub">
      <msub>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <mrow>
          <xsl:apply-templates select="m:sub[1]" />
        </mrow>
      </msub>
    </xsl:template>
  
    <xsl:template match="m:sSup">
      <msup>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <mrow>
          <xsl:apply-templates select="m:sup[1]" />
        </mrow>
      </msup>
    </xsl:template>
  
    <xsl:template match="m:sSubSup">
      <msubsup>
        <mrow>
          <xsl:apply-templates select="m:e[1]" />
        </mrow>
        <mrow>
          <xsl:apply-templates select="m:sub[1]" />
        </mrow>
        <mrow>
          <xsl:apply-templates select="m:sup[1]" />
        </mrow>
      </msubsup>
    </xsl:template>
  
    <xsl:template match="m:groupChr">
      <xsl:variable name="ndLastGroupChrPr" select="m:groupChrPr[last()]" />
      <xsl:variable name="sLowerCasePos" select="translate($ndLastGroupChrPr/m:pos/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="sLowerCaseVertJc" select="translate($ndLastGroupChrPr/m:vertJc/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="ndLastChr" select="$ndLastGroupChrPr/m:chr" />
      <xsl:variable name="chr">
        <xsl:choose>
          <xsl:when test="$ndLastChr and (not($ndLastChr/@m:val) or string-length($ndLastChr/@m:val) = 0)"></xsl:when>
          <xsl:when test="string-length($ndLastChr/@m:val) &gt;= 1">
            <xsl:value-of select="substring($ndLastChr/@m:val,1,1)" />
          </xsl:when>
          <xsl:otherwise>
            <xsl:text>&#x023DF;</xsl:text>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$sLowerCasePos = 'top'">
          <xsl:choose>
            <xsl:when test="$sLowerCaseVertJc = 'bot'">
              <mover accent="false">
                <mrow>
                  <xsl:apply-templates select="m:e[1]" />
                </mrow>
                <mo>
                  <xsl:value-of select="$chr" />
                </mo>
              </mover>
            </xsl:when>
            <xsl:otherwise>
              <munder accentunder="false">
                <mo>
                  <xsl:value-of select="$chr" />
                </mo>
                <mrow>
                  <xsl:apply-templates select="m:e[1]" />
                </mrow>
              </munder>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="$sLowerCaseVertJc = 'bot'">
              <mover accent="false">
                <mo>
                  <xsl:value-of select="$chr" />
                </mo>
                <mrow>
                  <xsl:apply-templates select="m:e[1]" />
                </mrow>
              </mover>
            </xsl:when>
            <xsl:otherwise>
              <munder accentunder="false">
                <mrow>
                  <xsl:apply-templates select="m:e[1]" />
                </mrow>
                <mo>
                  <xsl:value-of select="$chr" />
                </mo>
              </munder>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template name="fName">
      <xsl:for-each select="m:fName/*">
        <xsl:apply-templates select="." />
      </xsl:for-each>
    </xsl:template>
  
    <xsl:template match="m:func">
      <mrow>
        <mrow>
          <xsl:call-template name="fName" />
        </mrow>
        <mo>&#x02061;</mo>
        <mrow>
          <xsl:apply-templates select="m:e" />
        </mrow>
      </mrow>
    </xsl:template>
  
    <!-- 
      Template: match m:f 
      m:f maps directly to mfrac. 
    -->
    <xsl:template match="m:f">
      <xsl:variable name="sLowerCaseType" select="translate(m:fPr[last()]/m:type/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:choose>
        <xsl:when test="$sLowerCaseType='lin'">
          <mrow>
            <mrow>
              <xsl:apply-templates select="m:num[1]" />
            </mrow>
            <mo>/</mo>
            <mrow>
              <xsl:apply-templates select="m:den[1]" />
            </mrow>
          </mrow>
        </xsl:when>
        <xsl:otherwise>
          <mfrac>
            <xsl:call-template name="CreateMathMLFracProp">
              <xsl:with-param name="type" select="$sLowerCaseType" />
            </xsl:call-template>
            <mrow>
              <xsl:apply-templates select="m:num[1]" />
            </mrow>
            <mrow>
              <xsl:apply-templates select="m:den[1]" />
            </mrow>
          </mfrac>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
  
    <!-- 
      Template: CreateMathMLFracProp 
      
      Make fraction properties based on supplied parameters.
      OMML differentiates between a linear fraction and a skewed
      one. For MathML, we write both as bevelled.
    -->
    <xsl:template name="CreateMathMLFracProp">
      <xsl:param name="type" />
      <xsl:variable name="sLowerCaseType" select="translate($type, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:if test="$sLowerCaseType='skw' or $sLowerCaseType='lin'">
        <xsl:attribute name="bevelled">true</xsl:attribute>
      </xsl:if>
      <xsl:if test="$sLowerCaseType='nobar'">
        <xsl:attribute name="linethickness">0pt</xsl:attribute>
      </xsl:if>
      <xsl:choose>
        <xsl:when test="sLowerCaseNumJc='right'">
          <xsl:attribute name="numalign">right</xsl:attribute>
        </xsl:when>
        <xsl:when test="sLowerCaseNumJc='left'">
          <xsl:attribute name="numalign">left</xsl:attribute>
        </xsl:when>
      </xsl:choose>
      <xsl:choose>
        <xsl:when test="sLowerCaseDenJc='right'">
          <xsl:attribute name="numalign">right</xsl:attribute>
        </xsl:when>
        <xsl:when test="sLowerCaseDenJc='left'">
          <xsl:attribute name="numalign">left</xsl:attribute>
        </xsl:when>
      </xsl:choose>
    </xsl:template>
  
    <!-- 
      Template: match m:e | m:den | m:num | m:lim | m:sup | m:sub 
      
      These element delinate parts of an expression (like the numerator).
    -->
    <xsl:template match="m:e | m:den | m:num | m:lim | m:sup | m:sub">
      <xsl:choose>
        <!-- If there is no scriptLevel specified, just call through -->
        <xsl:when test="not(m:argPr[last()]/m:scrLvl/@m:val)">
          <xsl:apply-templates select="*" />
        </xsl:when>
        <!-- Otherwise, create an mstyle and set the script level -->
        <xsl:otherwise>
          <mstyle>
            <xsl:attribute name="scriptlevel">
              <xsl:value-of select="m:argPr[last()]/m:scrLvl/@m:val" />
            </xsl:attribute>
            <xsl:apply-templates select="*" />
          </mstyle>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template match="m:bar">
      <xsl:variable name="sLowerCasePos" select="translate(m:barPr/m:pos/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
      <xsl:variable name="fTop">
        <xsl:choose>
          <xsl:when test="$sLowerCasePos='top'">1</xsl:when>
          <xsl:otherwise>0</xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$fTop=1">
          <mover accent="false">
            <mrow>
              <xsl:apply-templates select="m:e[1]" />
            </mrow>
            <mo>
              <xsl:text>&#x00AF;</xsl:text>
            </mo>
          </mover>
        </xsl:when>
        <xsl:otherwise>
          <munder underaccent="false">
            <mrow>
              <xsl:apply-templates select="m:e[1]" />
            </mrow>
            <mo>
              <xsl:text>&#x005F;</xsl:text>
            </mo>
          </munder>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!-- 
      Template match m:d
      
      Process a delimiter. 
    -->
    <xsl:template match="m:d">
      <mfenced>
        <!-- open: default is '(' for both OMML and MathML -->
        <xsl:if test="m:dPr[1]/m:begChr/@m:val and not(m:dPr[1]/m:begChr/@m:val ='(')">
          <xsl:attribute name="open">
            <xsl:value-of select="m:dPr[1]/m:begChr/@m:val" />
          </xsl:attribute>
        </xsl:if>
  
        <!-- close: default is ')' for both OMML and MathML -->
        <xsl:if test="m:dPr[1]/m:endChr/@m:val and not(m:dPr[1]/m:endChr/@m:val =')')">
          <xsl:attribute name="close">
            <xsl:value-of select="m:dPr[1]/m:endChr/@m:val" />
          </xsl:attribute>
        </xsl:if>
  
        <!-- separator: the default is ',' for MathML, and '|' for OMML -->
        <xsl:choose>
          <!-- Matches MathML default. Write nothing -->
          <xsl:when test="m:dPr[1]/m:sepChr/@m:val = ','" />
  
          <!-- OMML default: | -->
          <xsl:when test="not(m:dPr[1]/m:sepChr/@m:val)">
            <xsl:attribute name="separators">
              <xsl:value-of select="'|'" />
            </xsl:attribute>
          </xsl:when>
  
          <xsl:otherwise>
            <xsl:attribute name="separators">
              <xsl:value-of select="m:dPr[1]/m:sepChr/@m:val" />
            </xsl:attribute>
          </xsl:otherwise>
        </xsl:choose>
  
        <!-- now write all the children. Put each one into an mrow just in case it produces multiple runs, etc -->
        <xsl:for-each select="m:e">
          <mrow>
            <xsl:apply-templates select="." />
          </mrow>
        </xsl:for-each>
      </mfenced>
    </xsl:template>
  
    <xsl:template match="m:r">
      <xsl:variable name="fNor">
        <xsl:choose>
          <xsl:when test="count(child::m:rPr[last()]/m:nor) = 0">0</xsl:when>
          <xsl:otherwise>
            <xsl:call-template name="ForceFalseStrVal">
              <xsl:with-param name="str" select="translate(child::m:rPr[last()]/m:nor/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$fNor=1">
          <mtext>
            <xsl:variable name="sOutput" select="translate(.//m:t, ' ', '&#xa0;')" />
            <xsl:value-of select="$sOutput" />
          </mtext>
        </xsl:when>
        <xsl:otherwise>
          <xsl:for-each select=".//m:t">
            <xsl:call-template name="ParseMt">
              <xsl:with-param name="sToParse" select="text()" />
              <xsl:with-param name="scr" select="../m:rPr[last()]/m:scr/@m:val" />
              <xsl:with-param name="sty" select="../m:rPr[last()]/m:sty/@m:val" />
              <xsl:with-param name="nor">0</xsl:with-param>
            </xsl:call-template>
          </xsl:for-each>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>

    <xsl:template name="CreateTokenAttributes">
      <xsl:param name="scr" />
      <xsl:param name="sty" />
      <xsl:param name="nor" />
      <xsl:param name="nCharToPrint" />
      <xsl:param name="sTokenType" />
      <xsl:choose>
        <xsl:when test="$nor=1">
          <xsl:attribute name="mathvariant">normal</xsl:attribute>
        </xsl:when>
        <xsl:otherwise>
          <xsl:variable name="mathvariant">
            <xsl:choose>
              <!-- numbers don't care -->
              <xsl:when test="$sTokenType='mn'" />
              <xsl:when test="$scr='monospace'">monospace</xsl:when>
              <xsl:when test="$scr='sans-serif' and $sty='i'">sans-serif-italic</xsl:when>
              <xsl:when test="$scr='sans-serif' and $sty='b'">bold-sans-serif</xsl:when>
              <xsl:when test="$scr='sans-serif' and $sty='bi'">sans-serif-bold-italic</xsl:when>
              <xsl:when test="$scr='sans-serif'">sans-serif</xsl:when>
              <xsl:when test="$scr='fraktur' and ($sty='b' or $sty='bi')">bold-fraktur</xsl:when>
              <xsl:when test="$scr='fraktur'">fraktur</xsl:when>
              <xsl:when test="$scr='double-struck'">double-struck</xsl:when>
              <xsl:when test="$scr='script' and ($sty='b' or $sty='bi')">bold-script</xsl:when>
              <xsl:when test="$scr='script'">script</xsl:when>
              <xsl:when test="($scr='roman' or not($scr) or $scr='') and $sty='b'">bold</xsl:when>
              <xsl:when test="($scr='roman' or not($scr) or $scr='') and $sty='i'">italic</xsl:when>
              <xsl:when test="($scr='roman' or not($scr) or $scr='') and $sty='p'">normal</xsl:when>
              <xsl:when test="($scr='roman' or not($scr) or $scr='') and $sty='bi'">bold-italic</xsl:when>
              <xsl:otherwise />
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="fontweight">
            <xsl:choose>
              <xsl:when test="$sty='b' or $sty='bi'">bold</xsl:when>
              <xsl:otherwise>normal</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <xsl:variable name="fontstyle">
            <xsl:choose>
              <xsl:when test="$sty='p' or $sty='b'">normal</xsl:when>
              <xsl:otherwise>italic</xsl:otherwise>
            </xsl:choose>
          </xsl:variable>
          <!-- Writing of attributes begins here -->
          <xsl:choose>
            <!-- Don't write mathvariant for operators unless they want to be normal -->
            <xsl:when test="$sTokenType='mo' and $mathvariant!='normal'" />
            <!-- A single character within an mi is already italics, don't write -->
            <xsl:when test="$sTokenType='mi' and $nCharToPrint=1 and ($mathvariant='' or $mathvariant='italic')" />
            <xsl:when test="$sTokenType='mi' and $nCharToPrint &gt; 1 and ($mathvariant='' or $mathvariant='italic')">
              <xsl:attribute name="mathvariant">
                <xsl:value-of select="'italic'" />
              </xsl:attribute>
            </xsl:when>
            <xsl:when test="$mathvariant!='italic' and $mathvariant!=''">
              <xsl:attribute name="mathvariant">
                <xsl:value-of select="$mathvariant" />
              </xsl:attribute>
            </xsl:when>
            <xsl:otherwise>
              <xsl:if test="not($sTokenType='mi' and $nCharToPrint=1) and $fontstyle='italic'">
                <xsl:attribute name="fontstyle">italic</xsl:attribute>
              </xsl:if>
              <xsl:if test="$fontweight='bold'">
                <xsl:attribute name="fontweight">bold</xsl:attribute>
              </xsl:if>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <xsl:template match="m:eqArr">
      <mtable>
        <xsl:for-each select="m:e">
          <mtr>
            <mtd>
              <xsl:choose>
                <xsl:when test="m:argPr[last()]/m:scrLvl/@m:val!='0' or 
  					            not(m:argPr[last()]/m:scrLvl/@m:val)  or 
  					            m:argPr[last()]/m:scrLvl/@m:val=''">
                  <mrow>
                    <maligngroup />
                    <xsl:call-template name="CreateEqArrRow">
                      <xsl:with-param name="align" select="1" />
                      <xsl:with-param name="ndCur" select="*[1]" />
                    </xsl:call-template>
                  </mrow>
                </xsl:when>
                <xsl:otherwise>
                  <mstyle>
                    <xsl:attribute name="scriptlevel">
                      <xsl:value-of select="m:argPr[last()]/m:scrLvl/@m:val" />
                    </xsl:attribute>
                    <maligngroup />
                    <xsl:call-template name="CreateEqArrRow">
                      <xsl:with-param name="align" select="1" />
                      <xsl:with-param name="ndCur" select="*[1]" />
                    </xsl:call-template>
                  </mstyle>
                </xsl:otherwise>
              </xsl:choose>
            </mtd>
          </mtr>
        </xsl:for-each>
      </mtable>
    </xsl:template>
  
    <xsl:template name="CreateEqArrRow">
      <xsl:param name="align" />
      <xsl:param name="ndCur" />
      <xsl:variable name="sAllMt">
        <xsl:for-each select="$ndCur/m:t">
          <xsl:value-of select="." />
        </xsl:for-each>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$ndCur/self::m:r">
          <xsl:call-template name="ParseEqArrMr">
            <xsl:with-param name="sToParse" select="$sAllMt" />
            <xsl:with-param name="scr" select="../m:rPr[last()]/m:scr/@m:val" />
            <xsl:with-param name="sty" select="../m:rPr[last()]/m:sty/@m:val" />
            <xsl:with-param name="nor">
              <xsl:choose>
                <xsl:when test="count($ndCur/m:rPr[last()]/m:nor) = 0">0</xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="ForceFalseStrVal">
                    <xsl:with-param name="str" select="translate($ndCur/m:rPr[last()]/m:nor/@m:val, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:with-param>
            <xsl:with-param name="align" select="$align" />
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:apply-templates select="$ndCur" />
        </xsl:otherwise>
      </xsl:choose>
      <xsl:if test="count($ndCur/following-sibling::*) &gt; 0">
        <xsl:variable name="cAmp">
          <xsl:call-template name="CountAmp">
            <xsl:with-param name="sAllMt" select="$sAllMt" />
            <xsl:with-param name="cAmp" select="0" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:call-template name="CreateEqArrRow">
          <xsl:with-param name="align" select="($align+($cAmp mod 2)) mod 2" />
          <xsl:with-param name="ndCur" select="$ndCur/following-sibling::*[1]" />
        </xsl:call-template>
      </xsl:if>
    </xsl:template>
  
    <xsl:template name="CountAmp">
      <xsl:param name="sAllMt" />
      <xsl:param name="cAmp" />
      <xsl:choose>
        <xsl:when test="string-length(substring-after($sAllMt, '&amp;')) &gt; 0 or 
  			                substring($sAllMt, string-length($sAllMt))='&#x0026;'">
          <xsl:call-template name="CountAmp">
            <xsl:with-param name="sAllMt" select="substring-after($sAllMt, '&#x0026;')" />
            <xsl:with-param name="cAmp" select="$cAmp+1" />
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="$cAmp" />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!-- 
      Template: ParseEqArrMr
      
      Similar to ParseMt, but this one has to do more for an equation array. 
      In equation arrays &amp; is a special character which denotes alignment.
      
      The &amp; in an equation works by alternating between meaning insert alignment spacing
      and insert alignment mark.  For each equation in the equation array
      there is an implied align space at the beginning of the equation.  Within each equation,
      the first &amp; means alignmark, the second, align space, the third, alignmark, etc.
      
      For this reason when parsing m:r's in equation arrays it is important to keep track of what
      the next ampersand will mean.
      
      $align=0 => Omml's align space, which is similar to MathML's maligngroup.
      $align=1 => Omml's alignment mark, which is similar to MathML's malignmark.
    -->
    <xsl:template name="ParseEqArrMr">
      <xsl:param name="sToParse" />
      <xsl:param name="sty" />
      <xsl:param name="scr" />
      <xsl:param name="nor" />
      <xsl:param name="align" />
      <xsl:if test="string-length($sToParse) &gt; 0">
        <xsl:choose>
          <xsl:when test="substring($sToParse,1,1) = '&amp;'">
            <xsl:choose>
              <xsl:when test="$align='0'">
                <maligngroup />
              </xsl:when>
              <xsl:when test="$align='1'">
                <malignmark />
              </xsl:when>
            </xsl:choose>
            <xsl:call-template name="ParseEqArrMr">
              <xsl:with-param name="sToParse" select="substring($sToParse,2)" />
              <xsl:with-param name="scr" select="$scr" />
              <xsl:with-param name="sty" select="$sty" />
              <xsl:with-param name="nor" select="$nor" />
              <xsl:with-param name="align">
                <xsl:choose>
                  <xsl:when test="$align='1'">0</xsl:when>
                  <xsl:otherwise>1</xsl:otherwise>
                </xsl:choose>
              </xsl:with-param>
            </xsl:call-template>
          </xsl:when>
          <xsl:otherwise>
            <xsl:variable name="sRepNumWith0">
              <xsl:call-template name="SReplaceNumWithZero">
                <xsl:with-param name="sToParse" select="$sToParse" />
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="sRepOperWith-">
              <xsl:call-template name="SReplaceOperWithMinus">
                <xsl:with-param name="sToParse" select="$sRepNumWith0" />
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="iFirstOper" select="string-length($sRepOperWith-) - string-length(substring-after($sRepOperWith-, '-'))" />
            <xsl:variable name="iFirstNum" select="string-length($sRepOperWith-) - string-length(substring-after($sRepOperWith-, '0'))" />
            <xsl:variable name="iFirstAmp" select="string-length($sRepOperWith-) - string-length(substring-after($sRepOperWith-, '&#x0026;'))" />
            <xsl:variable name="fNumAtPos1">
              <xsl:choose>
                <xsl:when test="substring($sRepOperWith-,1,1)='0'">1</xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:variable name="fOperAtPos1">
              <xsl:choose>
                <xsl:when test="substring($sRepOperWith-,1,1)='-'">1</xsl:when>
                <xsl:otherwise>0</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <xsl:choose>
              
              <!-- Case I: The string begins with neither a number, nor an operator -->
              <xsl:when test="$fNumAtPos1='0' and $fOperAtPos1='0'">
                <xsl:choose>
                  <xsl:when test="$nor = 0">
                    <mi>
                      <xsl:call-template name="CreateTokenAttributes">
                        <xsl:with-param name="scr" select="$scr" />
                        <xsl:with-param name="sty" select="$sty" />
                        <xsl:with-param name="nor" select="$nor" />
                        <xsl:with-param name="nCharToPrint" select="1" />
                        <xsl:with-param name="sTokenType" select="'mi'" />
                      </xsl:call-template>
                      <xsl:variable name="sOutput" select="translate(substring($sToParse, 1, 1), ' ', '&#xa0;')" />
                      <xsl:value-of select="$sOutput" />
                    </mi>
                  </xsl:when>
                  <xsl:otherwise>
                    <mtext>
                      <xsl:variable name="sOutput" select="translate(substring($sToParse, 1, 1), ' ', '&#xa0;')" />
                      <xsl:value-of select="$sOutput" />
                    </mtext>
                  </xsl:otherwise>
                </xsl:choose>
                <xsl:call-template name="ParseEqArrMr">
                  <xsl:with-param name="sToParse" select="substring($sToParse, 2)" />
                  <xsl:with-param name="scr" select="$scr" />
                  <xsl:with-param name="sty" select="$sty" />
                  <xsl:with-param name="nor" select="$nor" />
                  <xsl:with-param name="align" select="$align" />
                </xsl:call-template>
              </xsl:when>
  
              <!-- Case II: There is an operator at position 1 -->
              <xsl:when test="$fOperAtPos1='1'">
                <xsl:choose>
                  <xsl:when test="$nor = 0">
                    <mo>
                      <xsl:call-template name="CreateTokenAttributes">
                        <xsl:with-param name="scr" />
                        <xsl:with-param name="sty" />
                        <xsl:with-param name="nor" select="$nor" />
                        <xsl:with-param name="sTokenType" select="'mo'" />
                      </xsl:call-template>
                      <xsl:value-of select="substring($sToParse,1,1)" />
                    </mo>
                  </xsl:when>
                  <xsl:otherwise>
                    <mtext>
                      <xsl:value-of select="substring($sToParse,1,1)" />
                    </mtext>                  
                  </xsl:otherwise>
                </xsl:choose>
                <xsl:call-template name="ParseEqArrMr">
                  <xsl:with-param name="sToParse" select="substring($sToParse, 2)" />
                  <xsl:with-param name="scr" select="$scr" />
                  <xsl:with-param name="sty" select="$sty" />
                  <xsl:with-param name="nor" select="$nor" />
                  <xsl:with-param name="align" select="$align" />
                </xsl:call-template>
              </xsl:when>
  
              <!-- Case III: There is a number at position 1 -->
              <xsl:otherwise>
                <xsl:variable name="sConsecNum">
                  <xsl:call-template name="SNumStart">
                    <xsl:with-param name="sToParse" select="$sToParse" />
                    <xsl:with-param name="sPattern" select="$sRepNumWith0" />
                  </xsl:call-template>
                </xsl:variable>
                <xsl:choose>
                  <xsl:when test="$nor = 0">
                    <mn>
                      <xsl:call-template name="CreateTokenAttributes">
                        <xsl:with-param name="scr" />
                        <xsl:with-param name="sty" select="'p'"/>
                        <xsl:with-param name="nor" select="$nor" />
                        <xsl:with-param name="sTokenType" select="'mn'" />
                      </xsl:call-template>
                      <xsl:value-of select="$sConsecNum" />
                    </mn>
                  </xsl:when>
                  <xsl:otherwise>
                    <mtext>
                      <xsl:value-of select="$sConsecNum" />
                    </mtext>
                  </xsl:otherwise>
                </xsl:choose>
                <xsl:call-template name="ParseEqArrMr">
                  <xsl:with-param name="sToParse" select="substring-after($sToParse, $sConsecNum)" />
                  <xsl:with-param name="scr" select="$scr" />
                  <xsl:with-param name="sty" select="$sty" />
                  <xsl:with-param name="nor" select="$nor" />
                  <xsl:with-param name="align" select="$align" />
                </xsl:call-template>
              </xsl:otherwise>
            </xsl:choose>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:template>
  
    <!-- 
      Template: ParseMt
      
      Produce a run of text. Technically, OMML makes no distinction 
      between numbers, operators, and other characters in a run. For 
      MathML we need to break these into mi, mn, or mo elements. 
      
      See also ParseEqArrMr
    -->
    <xsl:template name="ParseMt">
      <xsl:param name="sToParse" />
      <xsl:param name="sty" />
      <xsl:param name="scr" />
      <xsl:param name="nor" />
      <xsl:if test="string-length($sToParse) &gt; 0">
        <xsl:variable name="sRepNumWith0">
          <xsl:call-template name="SReplaceNumWithZero">
            <xsl:with-param name="sToParse" select="$sToParse" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="sRepOperWith-">
          <xsl:call-template name="SReplaceOperWithMinus">
            <xsl:with-param name="sToParse" select="$sRepNumWith0" />
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="iFirstOper" select="string-length($sRepOperWith-) - string-length(substring-after($sRepOperWith-, '-'))" />
        <xsl:variable name="iFirstNum" select="string-length($sRepOperWith-) - string-length(substring-after($sRepOperWith-, '0'))" />
        <xsl:variable name="fNumAtPos1">
          <xsl:choose>
            <xsl:when test="substring($sRepOperWith-,1,1)='0'">1</xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:variable name="fOperAtPos1">
          <xsl:choose>
            <xsl:when test="substring($sRepOperWith-,1,1)='-'">1</xsl:when>
            <xsl:otherwise>0</xsl:otherwise>
          </xsl:choose>
        </xsl:variable>
        <xsl:choose>
  
          <!-- Case I: The string begins with neither a number, nor an operator -->
          <xsl:when test="$fOperAtPos1='0' and $fNumAtPos1='0'">
            <xsl:variable name="nCharToPrint">
              <xsl:choose>
                <xsl:when test="ancestor::m:fName">
                  <xsl:choose>
                    <xsl:when test="($iFirstOper=$iFirstNum) and 
  											($iFirstOper=string-length($sToParse)) and
  							                (substring($sRepOperWith-, string-length($sRepOperWith-))!='0') and 
  							                (substring($sRepOperWith-, string-length($sRepOperWith-))!='-')">
                      <xsl:value-of select="string-length($sToParse)" />
                    </xsl:when>
                    <xsl:when test="$iFirstOper &lt; $iFirstNum">
                      <xsl:value-of select="$iFirstOper - 1" />
                    </xsl:when>
                    <xsl:otherwise>
                      <xsl:value-of select="$iFirstNum - 1" />
                    </xsl:otherwise>
                  </xsl:choose>
                </xsl:when>
                <xsl:otherwise>1</xsl:otherwise>
              </xsl:choose>
            </xsl:variable>
            <mi>
              <xsl:call-template name="CreateTokenAttributes">
                <xsl:with-param name="scr" select="$scr" />
                <xsl:with-param name="sty" select="$sty" />
                <xsl:with-param name="nor" select="$nor" />
                <xsl:with-param name="nCharToPrint" select="$nCharToPrint" />
                <xsl:with-param name="sTokenType" select="'mi'" />
              </xsl:call-template>
              <xsl:variable name="sWrite" select="translate(substring($sToParse, 1, $nCharToPrint), ' ', '&#xa0;')" />
              <xsl:value-of select="$sWrite" />
            </mi>
            <xsl:call-template name="ParseMt">
              <xsl:with-param name="sToParse" select="substring($sToParse, $nCharToPrint+1)" />
              <xsl:with-param name="scr" select="$scr" />
              <xsl:with-param name="sty" select="$sty" />
              <xsl:with-param name="nor" select="$nor" />
            </xsl:call-template>
          </xsl:when>
  
          <!-- Case II: There is an operator at position 1 -->
          <xsl:when test="$fOperAtPos1='1'">
            <mo>
              <xsl:call-template name="CreateTokenAttributes">
                <xsl:with-param name="scr" />
                <xsl:with-param name="sty" />
                <xsl:with-param name="nor" select="$nor" />
                <xsl:with-param name="sTokenType" select="'mo'" />
              </xsl:call-template>
              <xsl:value-of select="substring($sToParse,1,1)" />
            </mo>
            <xsl:call-template name="ParseMt">
              <xsl:with-param name="sToParse" select="substring($sToParse, 2)" />
              <xsl:with-param name="scr" select="$scr" />
              <xsl:with-param name="sty" select="$sty" />
              <xsl:with-param name="nor" select="$nor" />
            </xsl:call-template>
          </xsl:when>
  
          <!-- Case III: There is a number at position 1 -->
          <xsl:otherwise>
            <xsl:variable name="sConsecNum">
              <xsl:call-template name="SNumStart">
                <xsl:with-param name="sToParse" select="$sToParse" />
                <xsl:with-param name="sPattern" select="$sRepNumWith0" />
              </xsl:call-template>
            </xsl:variable>
            <mn>
              <xsl:call-template name="CreateTokenAttributes">
                <xsl:with-param name="scr" select="$scr" />
                <xsl:with-param name="sty" select="'p'" />
                <xsl:with-param name="nor" select="$nor" />
                <xsl:with-param name="sTokenType" select="'mn'" />
              </xsl:call-template>
              <xsl:value-of select="$sConsecNum" />
            </mn>
            <xsl:call-template name="ParseMt">
              <xsl:with-param name="sToParse" select="substring-after($sToParse, $sConsecNum)" />
              <xsl:with-param name="scr" select="$scr" />
              <xsl:with-param name="sty" select="$sty" />
              <xsl:with-param name="nor" select="$nor" />
            </xsl:call-template>
          </xsl:otherwise>
        </xsl:choose>
      </xsl:if>
    </xsl:template>
  
    <!-- 
      Template: SNumStart 
      
      Return the longest substring of sToParse starting from the 
      start of sToParse that is a number. In addition, it takes the
      pattern string, which is sToParse with all of its numbers 
      replaced with a 0. sPattern should be the same length 
      as sToParse  
    -->
    <xsl:template name="SNumStart">
      <xsl:param name="sToParse" select="''" />
      <!-- if we don't get anything, take the string itself -->
      <xsl:param name="sPattern" select="'$sToParse'" />
      <xsl:choose>
        <!-- the pattern says this is a number, recurse with the rest -->
        <xsl:when test="substring($sPattern, 1, 1) = '0'">
          <xsl:call-template name="SNumStart">
            <xsl:with-param name="sToParse" select="$sToParse" />
            <xsl:with-param name="sPattern" select="substring($sPattern, 2)" />
          </xsl:call-template>
        </xsl:when>
        <!-- the pattern says we've run out of numbers. Take as many characters from sToParse as we shaved off sPattern -->
        <xsl:otherwise>
          <xsl:value-of select="substring($sToParse, 1, string-length($sToParse) - string-length($sPattern))" />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!--  
      Template SRepeatCharAcc
      
      The core of SRepeatChar with an accumulator. The current
      string is in param $acc, and we will double and recurse,
      if we're less than half of the required length or else just 
      add the right amount of characters to the accumulator and
      return
    -->
    <xsl:template name="SRepeatCharAcc">
      <xsl:param name="cchRequired" select="1" />
      <xsl:param name="ch" select="'-'" />
      <xsl:param name="acc" select="$ch" />
      <xsl:variable name="cchAcc" select="string-length($acc)" />
      <xsl:choose>
        <xsl:when test="(2 * $cchAcc) &lt; $cchRequired">
          <xsl:call-template name="SRepeatCharAcc">
            <xsl:with-param name="cchRequired" select="$cchRequired" />
            <xsl:with-param name="ch" select="$ch" />
            <xsl:with-param name="acc" select="concat($acc, $acc)" />
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <xsl:value-of select="concat($acc, substring($acc, 1, $cchRequired - $cchAcc))" />
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
  
    <!-- 
      Template SRepeatChar
      
      Generates a string nchRequired long by repeating the given character ch
    -->
    <xsl:template name="SRepeatChar">
      <xsl:param name="cchRequired" select="1" />
      <xsl:param name="ch" select="'-'" />
      <xsl:call-template name="SRepeatCharAcc">
        <xsl:with-param name="cchRequired" select="$cchRequired" />
        <xsl:with-param name="ch" select="$ch" />
        <xsl:with-param name="acc" select="$ch" />
      </xsl:call-template>
    </xsl:template>
  
    <!--  
      Template SReplaceOperWithMinus
      
      Go through the given string and replace every instance
      of an operator with a minus '-'. This helps quickly identify
      the first instance of an operator.
    -->
    <xsl:template name="SReplaceOperWithMinus">
      <xsl:param name="sToParse" select="''" />
      <xsl:value-of select="translate($sToParse, $sOperators, $sMinuses)" />
    </xsl:template>
  
    <!-- Template SReplaceNumWithZero
    
      Go through the given string and replace every instance
      of an number with a zero '0'. This helps quickly identify
      the first occurence of a number. 
      
      Considers the '.' and ',' part of a number iff they are sandwiched 
      between two other numbers. 0.3 will be recognized as a number,
      x.3 will not be. Since these characters can also be an operator, this 
      should be called before SReplaceOperWithMinus.
    -->
    <xsl:template name="SReplaceNumWithZero">
      <xsl:param name="sToParse" select="''" />
  
      <!-- First do a simple replace. Numbers will all be come 0's.
        After this point, the pattern involving the . or , that 
        we are looking for will become 0.0 or 0,0 -->
      <xsl:variable name="sSimpleReplace" select="translate($sToParse, $sNumbers, $sZeros)" />
  
      <!-- And then, replace 0.0 with just 000. This means that the . will 
        become part of the number -->
      <xsl:variable name="sReplacePeriod">
        <xsl:call-template name="SReplace">
          <xsl:with-param name="sInput" select="$sSimpleReplace" />
          <xsl:with-param name="sOrig" select="'0.0'" />
          <xsl:with-param name="sReplacement" select="'000'" />
        </xsl:call-template>
      </xsl:variable>
  
      <!-- And then, replace 0,0 with just 000. This means that the, will become part of the number -->
      <xsl:call-template name="SReplace">
        <xsl:with-param name="sInput" select="$sReplacePeriod" />
        <xsl:with-param name="sOrig" select="'0,0'" />
        <xsl:with-param name="sReplacement" select="'000'" />
      </xsl:call-template>
    </xsl:template>
  
    <!-- 
      Template to translate Word's borderBox properties into the menclose notation attribute
      The initial call to this SHOULD NOT pass an sAttribute.  Subsequent calls to 
      CreateMencloseNotationAttrFromBorderBoxAttr by CreateMencloseNotationAttrFromBorderBoxAttr will
      update the sAttribute as appropriate.
      
      CreateMencloseNotationAttrFromBorderBoxAttr looks at each attribute (fHideTop, fHideBot, etc.) one at a time
      in the order they are listed and passes a modified sAttribute to CreateMencloseNotationAttrFromBorderBoxAttr.
      Each successive call to CreateMencloseNotationAttrFromBorderBoxAttr knows which attribute to look at because 
      the previous call should have omitted passing the attribute it just analyzed.  This is why as you read lower 
      and lower in the template that each call to CreateMencloseNotationAttrFromBorderBoxAttr has fewer and fewer 
      attributes.
    -->
    <xsl:template name="CreateMencloseNotationAttrFromBorderBoxAttr">
      <xsl:param name="fHideTop" />
      <xsl:param name="fHideBot" />
      <xsl:param name="fHideLeft" />
      <xsl:param name="fHideRight" />
      <xsl:param name="fStrikeH" />
      <xsl:param name="fStrikeV" />
      <xsl:param name="fStrikeBLTR" />
      <xsl:param name="fStrikeTLBR" />
      <xsl:param name="sAttribute" />
  
      <xsl:choose>
        <xsl:when test="string-length($sAttribute) = 0">
          <xsl:choose>
            <xsl:when test="string-length($fHideTop) &gt; 0
                        and string-length($fHideBot) &gt; 0 
                        and string-length($fHideLeft) &gt; 0
                        and string-length($fHideRight) &gt; 0">
  
              <xsl:choose>
                <xsl:when test="$fHideTop = 0 
                                and $fHideBot = 0
                                and $fHideLeft = 0 
                                and $fHideRight = 0">
                  <!-- We can use 'box' instead of top, bot, left, and right.  Therefore, replace sAttribute with 'box' and begin analyzing params fStrikeH and below. -->
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute">
                      <xsl:text>box</xsl:text>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <!-- Can't use 'box', theremore, must analyze all attributes -->
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideTop" select="$fHideTop" />
                    <xsl:with-param name="fHideBot" select="$fHideBot" />
                    <xsl:with-param name="fHideLeft" select="$fHideLeft" />
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute">
                      <!-- Assume using all four (left right top bottom).  Subsequent calls will remove the sides which aren't to be includes. -->
                      <xsl:text>left right top bottom</xsl:text>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
          </xsl:choose>
        </xsl:when>
        <xsl:otherwise>
          <xsl:choose>
            <xsl:when test="string-length($fHideTop) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fHideTop=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideBot" select="$fHideBot" />
                    <xsl:with-param name="fHideLeft" select="$fHideLeft" />
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute">
                      <xsl:call-template name="SReplace">
                        <xsl:with-param name="sInput" select="$sAttribute" />
                        <xsl:with-param name="sOrig" select="'top'" />
                        <xsl:with-param name="sReplacement" select="''" />
                      </xsl:call-template>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideBot" select="$fHideBot" />
                    <xsl:with-param name="fHideLeft" select="$fHideLeft" />
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fHideBot) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fHideBot=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideLeft" select="$fHideLeft" />
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute">
                      <xsl:call-template name="SReplace">
                        <xsl:with-param name="sInput" select="$sAttribute" />
                        <xsl:with-param name="sOrig" select="'bottom'" />
                        <xsl:with-param name="sReplacement" select="''" />
                      </xsl:call-template>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideLeft" select="$fHideLeft" />
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fHideLeft) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fHideLeft=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute">
                      <xsl:call-template name="SReplace">
                        <xsl:with-param name="sInput" select="$sAttribute" />
                        <xsl:with-param name="sOrig" select="'left'" />
                        <xsl:with-param name="sReplacement" select="''" />
                      </xsl:call-template>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fHideRight" select="$fHideRight" />
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fHideRight) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fHideRight=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute">
                      <xsl:call-template name="SReplace">
                        <xsl:with-param name="sInput" select="$sAttribute" />
                        <xsl:with-param name="sOrig" select="'right'" />
                        <xsl:with-param name="sReplacement" select="''" />
                      </xsl:call-template>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeH" select="$fStrikeH" />
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fStrikeH) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fStrikeH=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="concat($sAttribute, ' horizontalstrike')" />
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeV" select="$fStrikeV" />
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fStrikeV) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fStrikeV=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="concat($sAttribute, ' verticalstrike')" />
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeBLTR" select="$fStrikeBLTR" />
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fStrikeBLTR) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fStrikeBLTR=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="concat($sAttribute, ' updiagonalstrike')" />
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="fStrikeTLBR" select="$fStrikeTLBR" />
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:when test="string-length($fStrikeTLBR) &gt; 0">
              <xsl:choose>
                <xsl:when test="$fStrikeTLBR=1">
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="sAttribute" select="concat($sAttribute, ' downdiagonalstrike')" />
                  </xsl:call-template>
                </xsl:when>
                <xsl:otherwise>
                  <xsl:call-template name="CreateMencloseNotationAttrFromBorderBoxAttr">
                    <xsl:with-param name="sAttribute" select="$sAttribute" />
                  </xsl:call-template>
                </xsl:otherwise>
              </xsl:choose>
            </xsl:when>
            <xsl:otherwise>
              <xsl:attribute name="notation">
                <xsl:value-of select="normalize-space($sAttribute)" />
              </xsl:attribute>
            </xsl:otherwise>
          </xsl:choose>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!-- Tristate (true, false, neither) from string value -->
    <xsl:template name="TFromStrVal">
      <xsl:param name="str" />
      <xsl:choose>
        <xsl:when test="$str = 'on' or $str = '1' or $str = 'true'">1</xsl:when>
        <xsl:when test="$str = 'off' or $str = '0' or $str = 'false'">0</xsl:when>
        <xsl:otherwise>-1</xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!-- Return 0 iff $str is explicitly set to a false value. Return true otherwise -->
    <xsl:template name="ForceFalseStrVal">
      <xsl:param name="str" />
      <xsl:variable name="tValue">
        <xsl:call-template name="TFromStrVal">
          <xsl:with-param name="str" select="$str"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$tValue = '0'">0</xsl:when>
        <xsl:otherwise>1</xsl:otherwise>
      </xsl:choose>
    </xsl:template>
  
    <!-- Return 1 iff $str is explicitly set to a true value. Return false otherwise -->
    <xsl:template name="ForceTrueStrVal">
      <xsl:param name="str" />
      <xsl:variable name="tValue">
        <xsl:call-template name="TFromStrVal">
          <xsl:with-param name="str" select="$str"/>
        </xsl:call-template>
      </xsl:variable>
      <xsl:choose>
        <xsl:when test="$tValue = '1'">1</xsl:when>
        <xsl:otherwise>0</xsl:otherwise>
      </xsl:choose>
    </xsl:template>
    
    <!-- END OMML to MML -->
    
</xsl:transform>