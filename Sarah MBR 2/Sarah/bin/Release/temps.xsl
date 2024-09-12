<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:template match="/">
    <rss version="2.0" xmlns:yweather="http://xml.weather.yahoo.com/ns/rss/1.0">
      <Current>
        <Current_Condition><xsl:value-of select="rss/channel/item/yweather:condition/@text"/></Current_Condition>
        <Current_Temp><xsl:value-of select="rss/channel/item/yweather:condition/@temp"/></Current_Temp>
      </Current>
    	<xsl:for-each select="//rss/channel/item/yweather:forecast">
        <Temps>
      		<High><xsl:value-of select="@high"/></High>
          <Low><xsl:value-of select="@low"/></Low>
        </Temps>
    	</xsl:for-each>
    </rss>
  </xsl:template>
</xsl:stylesheet>
