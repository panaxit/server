<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xson="http://panax.io/xson" xmlns="" version="1.0" id="PrettifyJSON">
	<xsl:variable name="validChars" select="'abcdefghijklmnñopqrstuvwxyzABCDEFGHIJKLMNÑOPQRSTUVWXYZ0123456789-_'"/>
	<xsl:template match="/">
		<xsl:apply-templates mode="raw-to-xson"/>
	</xsl:template>
	<xsl:template match="*" mode="raw-to-xson">
		<xsl:apply-templates mode="raw-to-xson"/>
	</xsl:template>
	<xsl:template match="o|l" mode="raw-to-xson">
		<xsl:apply-templates mode="raw-to-xson"/>
	</xsl:template>
	<xsl:template match="l/v" mode="raw-to-xson">
		<xsl:element name="xson:item">
			<xsl:apply-templates mode="raw-to-xson"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="a" mode="raw-to-xson">
		<xsl:variable name="name">
			<xsl:choose>
				<xsl:when test="number(translate(n,'&quot;',''))=translate(n,'&quot;','')">
					<xsl:value-of select="concat('@',translate(n,'&quot;',''))"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:value-of select="translate(translate(n,'&quot;',''),translate(n,$validChars,''),'@@@@@@@@@@@@@@@')"/>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:variable>
		<xsl:element name="{translate($name,'@','_')}">
			<xsl:if test="contains($name,'@')">
				<xsl:attribute name="xson:originalName">
					<xsl:value-of select="translate(n,'&quot;','')"/>
				</xsl:attribute>
			</xsl:if>
			<xsl:if test="l">
				<xsl:attribute name="xsi:type">xson:array</xsl:attribute>
			</xsl:if>
			<xsl:apply-templates select="*" mode="raw-to-xson"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="text()" mode="raw-to-xson">
		<xsl:value-of select="."/>
	</xsl:template>
	<xsl:template match="text()[starts-with(.,'&quot;')]" mode="raw-to-xson">
		<xsl:value-of select="substring(.,2,string-length(.)-2)"/>
	</xsl:template>
	<xsl:template match="text()[.='null']|*[.='']" mode="raw-to-xson"/>
	<xsl:template match="text()[.='null']" mode="raw-to-xson">
		<xsl:attribute name="xsi:nil">true</xsl:attribute>
	</xsl:template>
	<xsl:template match="n" mode="raw-to-xson"/>
	<xsl:template match="a[v='true' or v='false']/n" mode="raw-to-xson">
		<xsl:attribute name="xsi:type">boolean</xsl:attribute>
	</xsl:template>
	<xsl:template match="e" mode="raw-to-xson">
		<xsl:value-of select="@v"/>
	</xsl:template>
	<xsl:template match="a[number(v)=v]/n" mode="raw-to-xson">
		<xsl:attribute name="xsi:type">numeric</xsl:attribute>
	</xsl:template>
	<xsl:template match="a[starts-with(v,'&quot;')]/n" mode="raw-to-xson">
		<xsl:attribute name="xsi:type">string</xsl:attribute>
	</xsl:template>
	<xsl:template match="a[l]/n" mode="raw-to-xson">
		<xsl:attribute name="xsi:type">xson:array</xsl:attribute>
	</xsl:template>
	<xsl:template match="a[o]/n" mode="raw-to-xson">
		<xsl:attribute name="xsi:type">xson:object</xsl:attribute>
	</xsl:template>
	<xsl:template match="o[not(preceding-sibling::n)]" mode="raw-to-xson">
		<xsl:element name="xson:object">
			<xsl:apply-templates mode="raw-to-xson"/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="l[not(preceding-sibling::n)]" mode="raw-to-xson">
		<xsl:element name="xson:array">
			<xsl:apply-templates mode="raw-to-xson"/>
		</xsl:element>
	</xsl:template>
</xsl:stylesheet>