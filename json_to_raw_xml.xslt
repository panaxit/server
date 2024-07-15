<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns="" version="1.0" id="raw_json_compatibility">
	<xsl:variable name="node_name">olsc</xsl:variable>
	<xsl:variable name="translate-o">{[ ,</xsl:variable>
	<xsl:variable name="translate-c">}] </xsl:variable>
	<xsl:template match="/">
		<xsl:apply-templates/>
	</xsl:template>
	<xsl:template match="*" mode="value">
		<xsl:copy>
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="o|l|c" mode="value">
		<xsl:param name="is_string" select="false()"/>
		<xsl:value-of select="translate(name(),$node_name,$translate-o)"/>
		<xsl:apply-templates select="(text()|*)[1]" mode="value">
			<xsl:with-param name="is_string" select="$is_string"/>
		</xsl:apply-templates>
		<xsl:value-of select="translate(name(),$node_name,$translate-c)"/>
		<xsl:apply-templates select="(following-sibling::text()|following-sibling::*)[1]" mode="value">
			<xsl:with-param name="is_string" select="$is_string"/>
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="s" mode="value">
		<xsl:param name="is_string" select="false()"/>
		<xsl:value-of select="' '"/>
		<xsl:if test="$is_string">
			<xsl:apply-templates select="(following-sibling::text()|following-sibling::*)[1]" mode="value">
				<xsl:with-param name="is_string" select="$is_string"/>
			</xsl:apply-templates>
		</xsl:if>
	</xsl:template>
	<xsl:template match="r|f" mode="value">
		<xsl:param name="is_string" select="false()"/>
		<xsl:text>\n</xsl:text>
		<xsl:apply-templates select="(text()|*)[1]" mode="value">
			<xsl:with-param name="is_string" select="$is_string"/>
		</xsl:apply-templates>
		<xsl:apply-templates select="(following-sibling::text()|following-sibling::*)[1]" mode="value">
			<xsl:with-param name="is_string" select="$is_string"/>
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="e" mode="value">
		<xsl:param name="is_string" select="false()"/>
		<xsl:text>\\</xsl:text>
		<xsl:value-of select="text()"/>
		<xsl:apply-templates select="(following-sibling::text()|following-sibling::*)[1]" mode="value">
			<xsl:with-param name="is_string" select="$is_string"/>
		</xsl:apply-templates>
	</xsl:template>
	<xsl:template match="text()" mode="value">
		<xsl:param name="is_string" select="false()"/>
		<xsl:copy/>
		<xsl:if test="$is_string and not(substring(.,string-length(.),1)='&quot;')">
			<xsl:apply-templates select="(following-sibling::text()|following-sibling::*)[1]" mode="value">
				<xsl:with-param name="is_string" select="$is_string"/>
			</xsl:apply-templates>
		</xsl:if>
	</xsl:template>
	<xsl:template match="text()[substring(.,1,1)='&quot;']" mode="value">
		<xsl:param name="is_string" select="false()"/>
		<xsl:copy/>
		<xsl:if test="$is_string = false() and (string-length(.)=1 or not(substring(.,string-length(.),1)='&quot;'))">
			<xsl:apply-templates select="(following-sibling::text()|following-sibling::*)[1]" mode="value">
				<xsl:with-param name="is_string" select="true()"/>
			</xsl:apply-templates>
		</xsl:if>
	</xsl:template>
	<xsl:template match="l/text()">
		<xsl:element name="v">
			<xsl:apply-templates mode="value" select="."/>
		</xsl:element>
	</xsl:template>
	<xsl:template match="l">
		<xsl:copy>
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="o|text()[normalize-space(.)!='']|c"/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="o">
		<xsl:copy>
			<xsl:copy-of select="@*"/>
			<xsl:apply-templates select="a"/>
		</xsl:copy>
	</xsl:template>
	<xsl:template match="a">
		<xsl:variable name="following" select="(following-sibling::text()[normalize-space(.)!='']|following-sibling::*[not(self::f or self::r or self::c or self::s)])[1]"/>
		<xsl:copy>
			<xsl:element name="n">
				<xsl:value-of select="text()"/>
			</xsl:element>
			<xsl:choose>
				<xsl:when test="$following/self::o or $following/self::l">
					<xsl:apply-templates select="$following"/>
				</xsl:when>
				<xsl:otherwise>
					<xsl:element name="v">
						<xsl:apply-templates select="$following" mode="value"/>
					</xsl:element>
				</xsl:otherwise>
			</xsl:choose>
		</xsl:copy>
	</xsl:template>
</xsl:stylesheet>