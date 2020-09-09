<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format">
<xsl:template match="/">
	<datatypes>
	<xsl:for-each select="//TYPE/*">
		<datatype>
			<xsl:for-each select="./*">
				<xsl:if test="local-name(.) != 'HPSLENGTH' and local-name(.) != 'HPSTYPE' and local-name(.) != 'OWNER' and local-name(.) != 'LOCKED' and local-name(.) != 'LOWLIMIT' and local-name(.) != 'UPLIMIT' and local-name(.) != 'EDITMASK' and local-name(.) != 'DISPLAYLENGTH' and local-name(.) != 'INMASK' and local-name(.) != 'ALIGN' and local-name(.) != 'OUTMASK' and local-name(.) != 'VALIDATIONTYPE' and local-name(.) != 'EDITTYPE' and local-name(.) != 'VALIDATIONCODE' and local-name(.) != 'CD' and local-name(.) != 'NAME' and local-name(.) != 'EDITLENGTH' and local-name(.) != 'DISPLAYMASK' ">

					<xsl:attribute name="{local-name(.)}">
						<xsl:value-of select="."></xsl:value-of>
					</xsl:attribute>
				</xsl:if>
			</xsl:for-each>
					<xsl:attribute name="cd">
					<xsl:value-of select="./CD/."></xsl:value-of>
					</xsl:attribute>
					<xsl:attribute name="name">
					<xsl:value-of select="./NAME/."></xsl:value-of>
					</xsl:attribute>
					
			<xsl:if test="./EDITLENGTH/.!=''">
					<xsl:attribute name="editlength">
					<xsl:value-of select="./EDITLENGTH/."></xsl:value-of>
					</xsl:attribute>
			</xsl:if>					
			<xsl:if test="./DISPLAYMASK/.!=''">
					<xsl:attribute name="displaymask">
					<xsl:value-of select="./DISPLAYMASK/."></xsl:value-of>
					</xsl:attribute>
			</xsl:if>					
			<xsl:choose >
				<xsl:when test="./EDITTYPE/.='1'">
					<xsl:attribute name="edittype">text</xsl:attribute>
				</xsl:when>
				<xsl:when test="./EDITTYPE/.='2'">
					<xsl:attribute name="edittype">number</xsl:attribute>
				</xsl:when>
				<xsl:otherwise >
					<xsl:attribute name="edittype"><xsl:value-of select="./EDITTYPE/."></xsl:value-of></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

			<xsl:choose >
				<xsl:when test="./VALIDATIONCODE/.='0'">
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='1'">
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='2'">
					<xsl:attribute name="validation">account2cd</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='3'">
					<xsl:attribute name="validation">account0cd</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='4'">
					<xsl:attribute name="validation">Δάνειο με CD</xsl:attribute>
					<xsl:attribute name="validationcode">4</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='5'">
					<xsl:attribute name="validation">Δάνειο χωρίς CD</xsl:attribute>
					<xsl:attribute name="validationcode">5</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='6'">
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='7'">
					<xsl:attribute name="validation">Αριθμός Εγγραφής</xsl:attribute>
					<xsl:attribute name="validationcode">7</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='8'">
					<xsl:attribute name="validation">Ειδικός με CD</xsl:attribute>
					<xsl:attribute name="validationcode">8</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='9'">
					<xsl:attribute name="validation">Γενικός Λογαριασμός Δανείου</xsl:attribute>
					<xsl:attribute name="validationcode">9</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='10'">
					<xsl:attribute name="validation">account1cd</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='11'">
					<xsl:attribute name="validation">Τραπεζική Επιταγή</xsl:attribute>
					<xsl:attribute name="validationcode">11</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='12'">
					<xsl:attribute name="validation">Λογαριασμός ΕΘΝΟΚΑΡΤΑΣ</xsl:attribute>
					<xsl:attribute name="validationcode">12</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='13'">
					<xsl:attribute name="validation">Τραπεζική Εντολή</xsl:attribute>
					<xsl:attribute name="validationcode">13</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONCODE/.='14'">
					<xsl:attribute name="validation">Ιδιωτική Επιταγή</xsl:attribute>
					<xsl:attribute name="validationcode">14</xsl:attribute>
				</xsl:when>
				<xsl:otherwise >
					<xsl:attribute name="validationcode"><xsl:value-of select="./VALIDATIONCODE/."></xsl:value-of></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>


			<xsl:choose >
				<xsl:when test="./VALIDATIONTYPE/.='1'">
					<xsl:attribute name="validationtype">text</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONTYPE/.='2'">
					<xsl:attribute name="validationtype">number</xsl:attribute>
				</xsl:when>
				<xsl:when test="./VALIDATIONTYPE/.='3'">
					<xsl:attribute name="validationtype">date</xsl:attribute>
				</xsl:when>
				<xsl:otherwise >
					<xsl:attribute name="validationtype"><xsl:value-of select="./VALIDATIONTYPE/."></xsl:value-of></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

			<xsl:if test="./OUTMASK/.!=''">
					<xsl:attribute name="outmask">
					<xsl:value-of select="./OUTMASK/."></xsl:value-of>
					</xsl:attribute>
			</xsl:if>					
			<xsl:choose >
				<xsl:when test="./ALIGN/.=2">
					<xsl:attribute name="align">right</xsl:attribute>
				</xsl:when>
				<xsl:when test="./ALIGN/.=1">
					<xsl:attribute name="align">left</xsl:attribute>
				</xsl:when>
				<xsl:otherwise >
					<xsl:attribute name="align"><xsl:value-of select="./ALIGN/."></xsl:value-of></xsl:attribute>
				</xsl:otherwise>
			</xsl:choose>

				
		</datatype>
	</xsl:for-each>
	</datatypes>
</xsl:template>
</xsl:stylesheet>
