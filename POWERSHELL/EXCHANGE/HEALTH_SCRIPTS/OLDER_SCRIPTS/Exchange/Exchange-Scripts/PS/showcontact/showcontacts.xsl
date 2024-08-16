<?xml version="1.0" encoding="ISO-8859-1"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:template match="/"  xmlns:a='DAV:' xmlns:d='urn:schemas:contacts:' xmlns:e='http://schemas.microsoft.com/mapi/'>
    <table>
      <tr bgcolor="#9acd32">
        <th>Name</th>
        <th>Email Address</th>
      </tr>
      <xsl:for-each select='a:multistatus/a:response'>
      <xsl:sort select="a:propstat/a:prop/d:cn"/>
        <tr>
          <td><xsl:attribute name='target'>_blank</xsl:attribute>
	      <xsl:value-of select='a:propstat/a:prop/d:cn'/>
          </td>
	  <td><xsl:attribute name='target'>_blank</xsl:attribute>
	      <xsl:value-of select='a:propstat/a:prop/e:email1emailaddress'/>
          </td>
        </tr>
      </xsl:for-each>
    </table>
</xsl:template>
</xsl:stylesheet>
