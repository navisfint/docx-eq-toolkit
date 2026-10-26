<?xml version="1.0" encoding="UTF-8"?>
<!-- MathType OMML2MML wrapper stylesheet.
     Purpose: MathType's CheckVersion_Omml2MmlXslFile2 reads the whole XSL via
     VBA Input$(LOF(f), f) on a text-mode handle, which raises runtime error 62
     ("Input past end of file") for files above roughly 64 KB. The stock
     Office16 OMML2MML.XSL is 115781 bytes, so that check always fails and
     XformOMML2MML degrades to the sentinel "omml_xform_error".
     This small wrapper keeps the version signature and delegates the real
     transform to the stock stylesheet via xsl:import.
     Signature token expected by MathType: &#x2146 -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:import href="file:///C:/Program%20Files/Microsoft%20Office/root/Office16/OMML2MML.XSL"/>
  <xsl:output method="xml" encoding="UTF-16"/>
</xsl:stylesheet>
