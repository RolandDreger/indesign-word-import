<?xml version="1.0" encoding="UTF-8"?>
<!--
    
    Merge stylesheets and scripts for InDesign import
    
    src/docs2html.xsl + src/docx2indesign.xsl => docx2indesign.xsl
    
    Created: February 21, 2022
    Modified: February 21, 2022
    
-->
<p:declare-step 
    xmlns:p="http://www.w3.org/ns/xproc" 
    xmlns:c="http://www.w3.org/ns/xproc-step" 
    name="build4indesign"
    version="1.0">
    
    <p:input port="source">
        <p:document href="docx2indesign.xsl"></p:document>
    </p:input>
    <p:output port="result"/>
    
    <p:xslt>
        <p:input port="stylesheet">
            <p:document href="merge4indesign.xsl"></p:document>
        </p:input>
        <p:input port="parameters">
            <p:empty/>
        </p:input>  
    </p:xslt>
    
    <p:store href="../docx2indesign.xsl" method="xml"/>
    
</p:declare-step>