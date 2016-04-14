<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:svg="http://www.w3.org/2000/svg"> 
    
    <xsl:template match="@*|node()">
        <xsl:copy>
            <xsl:apply-templates select="@*|node()"/>
        </xsl:copy>
    </xsl:template>

    <xsl:template match ="/">       
        <xsl:for-each select="./svg:svgs/svg:svg">
            <xsl:variable name="fileNAME" select="./svg:title[1]/text()"/>
            <xsl:result-document method="xml" href="{$fileNAME}.svg"> 
             <xsl:copy>
                 <xsl:apply-templates select="@*|node()"/>
             </xsl:copy>             
            </xsl:result-document> 
        </xsl:for-each> 
    </xsl:template>

    <xsl:template match="svg:text">
        <xsl:choose>
            <xsl:when test="contains(./@id,'axlabel')">
                <xsl:variable name="elemNAME">
                    <xsl:value-of select="name(.)"/>
                </xsl:variable>
                <xsl:variable name="chartHEIGHT">
                    <xsl:value-of select="substring-before(./ancestor::svg:svg/@height,'cm')"/>
                </xsl:variable>
                <xsl:variable name="fontSIZE">
                    <xsl:value-of select="./@font-size"/>
                </xsl:variable>
                <xsl:element name="{$elemNAME}" namespace="http://www.w3.org/2000/svg">
                    <xsl:attribute name="x" select="./@x"></xsl:attribute>
                    <xsl:attribute name="y" select="$chartHEIGHT - ./@y + ($fontSIZE div 1.3)"></xsl:attribute>
                    <xsl:copy-of select="*|@*[local-name() != 'x'][local-name() != 'y']|text()"/>
                </xsl:element>
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="."/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>

    <xsl:template match="svg:path">
        <xsl:choose>
            <xsl:when test="contains(./@id,'PATH')">
                <xsl:variable name="elemNAME">
                    <xsl:value-of select="name(.)"/>
                </xsl:variable>
                <xsl:variable name="chartHEIGHT">
                    <xsl:value-of select="substring-before(./ancestor::svg:svg/@height,'cm')"/>
                </xsl:variable>
                <xsl:element name="{$elemNAME}" namespace="http://www.w3.org/2000/svg">
                    <xsl:attribute name="transform" select="concat('translate(0,',$chartHEIGHT,') scale(1,-1)')"></xsl:attribute>
                    <xsl:copy-of select="*|@*"/>
                </xsl:element>
            </xsl:when>
            <xsl:otherwise>
                <xsl:copy-of select="."/>
            </xsl:otherwise>
        </xsl:choose>
    </xsl:template>
    
</xsl:stylesheet>
