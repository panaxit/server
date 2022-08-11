<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:x="http://panax.io/xover"
xmlns:source="http://panax.io/source"
xmlns:query="http://panax.io/xover/binding/query"
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
>
  <xsl:output method="xml" indent="no" omit-xml-declaration="yes"/>
  <xsl:template match="node() | text()" priority="-1">
    <xsl:copy>
      <xsl:if test="not(parent::*)">
        <xsl:copy-of select="//namespace::*[name()!='']"/>
      </xsl:if>
      <xsl:apply-templates select="@*|*|text()"/>
    </xsl:copy>
  </xsl:template>

  <xsl:template match="@*" priority="-1">
    <xsl:attribute name="{name()}" namespace="{namespace-uri()}">
      <xsl:value-of select="normalize-space(.)"/>
    </xsl:attribute>
  </xsl:template>

</xsl:stylesheet>