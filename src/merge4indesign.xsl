<?xml version="1.0" encoding="UTF-8"?>
<!--
    
    Merge Stylesheets for InDesign
    
    src/docs2html.xsl | src/docx2indesign.xsl (with xsl:import) -> docx2indesign.xsl
    
    Created: 16. January 2022
    Modified: 2. February 2022
    
-->
<xsl:stylesheet 
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    exclude-result-prefixes="xs"
    version="2.0">
    
    <xsl:key name="match-templates" match="xsl:template" use="@match"/>
    <xsl:key name="named-templates" match="xsl:template" use="@name"/>
    <xsl:key name="params" match="xsl:param" use="@name"/>
    <xsl:key name="vars" match="xsl:variable" use="@name"/>
    <xsl:key name="attribute-sets" match="xsl:attribute-set" use="@name"/>
    <xsl:key name="keys" match="xsl:key" use="@name"/>
    
    <xsl:variable name="root" select="/"/>
    
    <xsl:output method="xml" indent="no"/>
    
    <!--<xsl:strip-space elements="*"/>
    <xsl:preserve-space elements="xsl:text"/>-->
    
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>
    
    <xsl:template match="xsl:import">
        <!-- Parameters (Toplevel), Variables (Toplevel), Attribute Sets, Key Defs, Templates, Comments -->
        <xsl:apply-templates select="
            document(@href)/*/xsl:param[not(key('params', @name, $root))] | 
            document(@href)/*/xsl:variable[not(key('vars', @name, $root))] | 
            document(@href)/*/xsl:attribute-set[not(key('attribute-sets', @name, $root))] | 
            document(@href)/*/xsl:key[not(key('keys', @name, $root))] | 
            document(@href)/*/xsl:template[not(key('match-templates', @match, $root) or key('named-templates', @name, $root))] | 
            document(@href)/*/comment()
         "/>
    </xsl:template>
    
</xsl:stylesheet>