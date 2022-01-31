<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    xmlns:sd="http://www.w3.org/2001/XMLSchema"
    xmlns:xx="http://www.w3.org/2001/XMLSchema"
    exclude-result-prefixes="xs"
    version="1.0">
    
    <xsl:import href="stylesheet_one.xsl"/>
    
    <xsl:output method="xml"/>
    
    
    <!-- Start Two -->
    
    <xsl:param name="base-uri" select="''"/>
    <xsl:variable name="tag" select="'root'"/>
    
    <xsl:template match="/">
        <xsl:variable name="two"/>
    </xsl:template>
    
    <xsl:template name="root">
        <xsl:variable name="two"/>
    </xsl:template>
   
    <!-- END Two -->
    
</xsl:stylesheet>