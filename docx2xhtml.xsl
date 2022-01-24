<?xml version="1.0" encoding="UTF-8"?>

<!--    
        
    Microsoft Word Document to XHTML
    
    30. September 2021
    14. Dezember 2021
    
    Author: Roland Dreger, www.rolanddreger.net
    
    ToDo:
    
    - Tag names without dash "-"
    - Citation source elements
    
-->

<xsl:transform 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
    xmlns:dc="http://purl.org/dc/elements/1.1/" 
    xmlns:dcterms="http://purl.org/dc/terms/" 
    exclude-result-prefixes="xs"
    version="1.0"
>
    
    <xsl:import href="docx2html.xsl"/>
    
    <xsl:output 
        method="xhtml" 
        version="1.0" 
        doctype-public="-//W3C//DTD XHTML 1.1//EN" 
        doctype-system="http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd" 
        media-type="application/xhtml+xml" 
        omit-xml-declaration="no" 
        indent="yes" 
        standalone="yes"
    />
    
    <!-- Document Namespace -->
    <xsl:param name="ns" select="'http://www.w3.org/1999/xhtml'"/>
    
    <!-- Head -->
    <xsl:template name="create-head-section">
        <xsl:element name="head" namespace="{$ns}">
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
        </xsl:element>
    </xsl:template>
    
</xsl:transform>