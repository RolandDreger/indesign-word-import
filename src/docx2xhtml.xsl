<?xml version="1.0" encoding="UTF-8"?>

<!--    
        
    Microsoft Word Document => HTML => XHTML
    (XHTML Module)
    
    Created: September 30, 2021
    Modified: February 5, 2022
    
    Author: Roland Dreger, www.rolanddreger.net
    
    # Default namespaces:
    
    <html xmlns="http://www.w3.org/1999/xhtml"> 
    <math xmlns="http://www.w3.org/1998/Math/MathML"> 
    <svg xmlns="http://www.w3.org/2000/svg"> 

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
        method="xml" 
        version="1.0" 
        encoding="UTF-8"
        doctype-public="" 
        doctype-system=""
        media-type="application/xhtml+xml" 
        omit-xml-declaration="yes" 
        indent="yes" 
    />
    
    <!-- Document Namespace -->
    <xsl:param name="ns" select="'http://www.w3.org/1999/xhtml'"/>
    

    <!-- +++++++++++++ -->
    <!-- + Templates + -->
    <!-- +++++++++++++ -->
    
    <xsl:template match="/">
        <xsl:text disable-output-escaping='yes'>&lt;!DOCTYPE html&gt;&#x0d;</xsl:text>
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
            <xsl:element name="title" namespace="{$ns}">
                <xsl:value-of select="$core-props/dc:title"/>
            </xsl:element>
            <xsl:element name="meta" namespace="{$ns}">
                <xsl:attribute name="charset">
                    <xsl:value-of select="'UTF-8'"/>
                </xsl:attribute>
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