<?xml version="1.0" encoding="UTF-8"?>
<p:declare-step name="transform4test" 
    xmlns:p="http://www.w3.org/ns/xproc" 
    xmlns:c="http://www.w3.org/ns/xproc-step" 
    xmlns:rd="http://rolanddreger.net" 
    xmlns:xs="http://www.w3.org/2001/XMLSchema"
    exclude-inline-prefixes="c rd xs" 
    version="3.0">
    
    <p:input port="source" sequence="true" >
        <rd:config>
            <!-- Path to Word-XML-Document (must be relative to XSLT stylesheet) -->
            <rd:item name="word-xml-document" href="../tests/list/list.xml"/>
            <!-- Selected modes (space separated values) -->
            <rd:item name="modes" value="indesign html xhtml"/> 
            <!-- XSLT stylesheets -->
            <rd:item name="stylesheet" mode="indesign" href="docx2indesign.xsl" result-document-extension="xml"/>
            <rd:item name="stylesheet" mode="html" href="docx2html.xsl" result-document-extension="html"/>
            <rd:item name="stylesheet" mode="xhtml" href="docx2xhtml.xsl" result-document-extension="xhtml"/>
        </rd:config>
    </p:input>
    <p:output port="result" sequence="true" serialization="map{'indent':false()}"/>
    
    <p:variable name="modes" as="xs:string" select="/rd:config/rd:item[@name = 'modes']/@value"/>
    <p:variable name="document-file-name" select="/rd:config/rd:item[@name = 'word-xml-document']/@href"/>
    
    <!-- Make absolute URIs for Word-XML-Document and stylesheets -->
    <p:make-absolute-uris name="absolute-uri-config" match="/rd:config/rd:item/@href"/>
    
    <p:variable name="word-xml-document-uri" select="/rd:config/rd:item[@name = 'word-xml-document']/@href"/>
    
    <!-- Loop over stylesheets (selected modes)  -->
    <p:for-each>
        <p:with-input select="/rd:config/rd:item[@name = 'stylesheet'][contains($modes,@mode)]"/>
        
        <p:output port="result" sequence="true">
            <p:pipe port="result" step="xslt"/>
            <p:pipe port="result-uri" step="store"/>
        </p:output>
        
        <!-- Transform Word-XML-Document -->
        <p:variable name="current-mode" as="xs:string" select="string(/rd:item/@mode)"/>
        <p:variable name="stylesheet-uri" select="/rd:item/@href"></p:variable>
        <p:variable name="result-document-extension" select="/rd:item/@result-document-extension"/>

        <p:xslt name="xslt" version="3.0">
            <p:with-input port="source" href="{$word-xml-document-uri}"/>
            <p:with-input port="stylesheet" href="{$stylesheet-uri}"/>
            <p:with-option name="parameters" select="map{ 
                    'document-file-name': $document-file-name 
                }"/>
        </p:xslt>
        
        <!-- Store result document -->
        <p:variable name="result-document-base-uri" select="p:document-property(.,'base-uri')"/>
        <p:variable name="result-document-name" select="concat(substring-before(tokenize($result-document-base-uri,'/')[last()], '.xml'), '.', $result-document-extension)"/>
        <p:variable name="result-document-path" select="resolve-uri(concat($current-mode,'/', $result-document-name), $result-document-base-uri)"/>
        
        <p:store name="store" href="{$result-document-path}" serialization="map{'indent':false()}"/>

    </p:for-each>
    
</p:declare-step>