<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <description>
    	<xsl:for-each select="//rss/channel/item">
      		<A><xsl:value-of select="description"/></A>
    	</xsl:for-each>
    </description>
  </xsl:template>
</xsl:stylesheet>
