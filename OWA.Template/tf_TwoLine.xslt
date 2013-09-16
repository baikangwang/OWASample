<view>
  <query>
    &quot;http://schemas.microsoft.com/exchange/smallicon&quot; as smicon, &quot;http://schemas.microsoft.com/mapi/sent_representing_name&quot; as from, &quot;urn:schemas:httpmail:datereceived&quot; as recvd, &quot;http://schemas.microsoft.com/mapi/proptag/x10900003&quot; as flag, &quot;http://schemas.microsoft.com/mapi/subject&quot; as subj, &quot;http://schemas.microsoft.com/exchange/x-priority-long&quot; as prio, &quot;urn:schemas:httpmail:hasattachment&quot; as fattach,
  </query>
  <group>
    <xsl:template xmlns:xsl="uri:xsl" xmlns:a="DAV:">
      <xsl:script>
        var re=/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2}).\d{3}Z?$/;

        function formatDateTimeString(szISODate,szDateFormat,szTimeFormat)
        {
        re.lastIndex=0;
        var arr = re.exec( szISODate );
        if (null == arr) return "";
        var objDate = new Date(Date.UTC(arr[1],arr[2]-1,arr[3],arr[4],arr[5],arr[6],0));
        var szFormattedDate = formatDate(objDate.getVarDate(), szDateFormat);
        var szFormattedTime = formatTime(objDate.getVarDate(), szTimeFormat);
        //Check for null character and strip it off if present
        if(szFormattedDate.charCodeAt(szFormattedDate.length -1) == 0)
        {
        szFormattedDate = szFormattedDate.substring(0,szFormattedDate.length-1)
        }
        return szFormattedDate + " " + szFormattedTime;
        }
      </xsl:script>
      <TABLE isTabularView="true" cellpadding="0" cellspacing="0" style="table-layout:fixed;width:100%;height:100%">
        <TR>
          <TD>
            <TABLE id="tblHeader" class="tblHdr" cellpadding="1" cellspacing="0" style="table-layout:fixed;width:100%">
              <TR id="sizingRow" style="display:none">
                <TD style="width:20px;"></TD>
                <TD colspan="16"></TD>
                <TD style="width:10px;"></TD>
                <TD style="width:10px;"></TD>
                <TD style="width:10px;"></TD>
                <TD style="width:10px;"></TD>
                <TD style="width:18px;"></TD>
                <TD id="idFill" style="width:16px;"></TD>
              </TR>
              <TR class="vwHdrTR">
                <TD id="tdColumn" class="vwHdrTD vwHdrNoBrdr" rowspan="2" valign="top" colName="" dataType="string" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" nowrap="true" style="width:20px;text-align:center;" prop="http://schemas.microsoft.com/exchange/smallicon">
                  <IMG id="imgCol" width="10" height="14" style="cursor:hand" src="/exchweb/img/view-document.gif"/>
                </TD>
                <TD id="tdColumn" class="vwHdrTD" colspan="11" colName="" dataType="string" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" nowrap="true" style="width:55%;padding:0px 3px;" prop="http://schemas.microsoft.com/mapi/sent_representing_name" ></TD>
                <TD id="tdColumn" class="vwHdrTD RTLrevA" colspan="9"  colName="" dataType="date" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" nowrap="true" style="width:45%;padding:0px 3px;" prop="urn:schemas:httpmail:datereceived"></TD>
                <TD id="tdColumn" class="vwHdrTD vwHdrNoBrdr" rowspan="2" colName="" dataType="i4" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" nowrap="true" style="width:18px;text-align:center;padding-left:2px;padding:0px 2px;" prop="http://schemas.microsoft.com/mapi/proptag/x10900003">
                  <IMG id="imgCol" width="10" height="14" src="/exchweb/img/view-flag.gif"/>
                </TD>
              </TR>
              <TR class="vwHdrTR">
                <TD id="tdColumn" class="vwHdrTD vwHdrNoBrdr" colspan="16" colName="" dataType="string" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" nowrap="true" style="width:100%;padding:0px 2px;" prop="http://schemas.microsoft.com/mapi/subject"></TD>
                <TD id="tdColumn" class="vwHdrTD vwHdrNoBrdr" colspan="2" colName="" dataType="i4" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" align="center" prop="http://schemas.microsoft.com/exchange/x-priority-long">
                  <IMG id="imgCol" width="10" height="14" src="/exchweb/img/view-importance.gif"/>
                </TD>
                <TD id="tdColumn" class="vwHdrTD vwHdrNoBrdr" colspan="2" colName="" dataType="boolean" sortdir="" sortable="1" onclick="if(null == event.returnValue) event.returnValue=this;" align="center" prop="urn:schemas:httpmail:hasattachment">
                  <IMG id="imgCol" width="10" height="14" src="/exchweb/img/view-paperclip.gif"/>
                </TD>
              </TR>
              <TR>
                <TD colspan="22" class="vwHdrBrdr1">&#160;</TD>
              </TR>
              <TR>
                <TD colspan="22" class="vwHdrBrdr2">&#160;</TD>
              </TR>
              <TR>
                <TD colspan="22" class="vwHdrBrdr3">&#160;</TD>
              </TR>
            </TABLE>
          </TD>
        </TR>
        <TR>
          <TD style="height:100%">
            <DIV class="msgViewerCont" id="dvContents">
              <TABLE cellpadding="1" cellspacing="0" style="table-layout:fixed;width:100%">
                <TR id="sizingRow" style="display:none">
                  <TD style="width:20px;"></TD>
                  <TD colspan="16"></TD>
                  <TD style="width:10px;"></TD>
                  <TD style="width:10px;"></TD>
                  <TD style="width:10px;"></TD>
                  <TD style="width:17px;"></TD>
                  <TD style="width:17px;"></TD>
                </TR>
                <xsl:for-each select="a:multistatus/a:response">
                  <TR id="unSelected" itemTR="1" onmousedown="if(null == event.returnValue) event.returnValue=this" onclick="if(null == event.returnValue) event.returnValue=this" ondblclick="if(null == event.returnValue) event.returnValue=this" oncontextmenu="if(null == event.returnValue) event.returnValue=this">
                    <xsl:attribute name="style">
                      <xsl:if test="a:propstat/a:prop/r[.='0']">font-weight:bold;</xsl:if>
                    </xsl:attribute>
                    <xsl:attribute name="_href">
                      <xsl:value-of select="a:href"/>
                    </xsl:attribute>
                    <xsl:attribute name="messageClass">
                      <xsl:value-of select="a:propstat/a:prop/m"/>
                    </xsl:attribute>
                    <xsl:attribute name="read">
                      <xsl:value-of select="a:propstat/a:prop/r"/>
                    </xsl:attribute>
                    <A onclick="return(false)" tabindex="-1">
                      <xsl:attribute name="id">
                        <xsl:value-of select="a:href"/>
                      </xsl:attribute>
                      <xsl:attribute name="href">
                        <xsl:value-of select="a:href"/>
                      </xsl:attribute>
                      <TD rowspan="2" nowrap="true" valign="top" class="vwItemSep" style="cursor:default;text-align:center;width:20px;">
                        <IMG width="16" height="16">
                          <xsl:attribute name="src">
                            <xsl:value-of select="a:propstat/a:prop/smicon"/>
                          </xsl:attribute>
                        </IMG>
                      </TD>
                      <TD colspan="11" nowrap="true" class="vwItmTd" style="padding:2px 0px 1px 0px;">
                        <xsl:value-of select="a:propstat/a:prop/from"/>
                      </TD>
                      <TD colspan="9" nowrap="true" class="RTLrevA vwItmTd" style="padding:2px 3px 1px 3px;">
                        <xsl:for-each select="a:propstat/a:prop/recvd">
                          <xsl:eval>formatDateTime</xsl:eval>
                        </xsl:for-each>
                      </TD>
                      <TD id="idFlag" rowspan="2" nowrap="true" style="width:18px;cursor:hand;text-align: center;border-bottom:1px solid #DDDDDD;padding:0px 2px;" onmouseout='this.style.backgroundColor = "";' onmouseover='this.style.backgroundColor = "#CCCCCC";' onclick="if(null == event.returnValue) event.returnValue=this" onmousedown="if(null == event.returnValue) event.returnValue=this" ondblclick="if(null == event.returnValue) event.returnValue=this"  oncontextmenu="if(null == event.returnValue) event.returnValue=this">
                        <xsl:attribute name="flagState">
                          <xsl:value-of select="a:propstat/a:prop/flag"/>
                        </xsl:attribute>
                        <xsl:attribute name="flagColor">
                          <xsl:value-of select="a:propstat/a:prop/flagcolor"/>
                        </xsl:attribute>
                        <xsl:attribute name="class">
                          <xsl:choose>
                            <xsl:when test="a:propstat/a:prop/flagcolor[.='1']">purpleFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flagcolor[.='2']">orangeFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flagcolor[.='3']">greenFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flagcolor[.='4']">yellowFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flagcolor[.='5']">blueFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flagcolor[.='6']">redFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flag[.='1']">completeFlg</xsl:when>
                            <xsl:when test="a:propstat/a:prop/flag[.='2']">noFlg</xsl:when>
                            <xsl:otherwise>noFlg</xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <IMG width="16" height="16">
                          <xsl:attribute name="src">
                            /exchweb/img/
                            <xsl:choose>
                              <xsl:when test="a:propstat/a:prop/flagcolor[.='1']">flg-1.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flagcolor[.='2']">flg-2.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flagcolor[.='3']">flg-3.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flagcolor[.='4']">flg-4.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flagcolor[.='5']">flg-5.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flagcolor[.='6']">flg-6.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flag[.='1']">flg-compl.gif</xsl:when>
                              <xsl:when test="a:propstat/a:prop/flag[.='2']">flg-sender.gif</xsl:when>
                              <xsl:otherwise>flg-empty.gif</xsl:otherwise>
                            </xsl:choose>
                          </xsl:attribute>
                        </IMG>
                      </TD>
                    </A>
                  </TR>
                  <TR id="trChld" style="color:#716F64" onmousedown="if(null == event.returnValue) event.returnValue=this" onclick="if(null == event.returnValue) event.returnValue=this" ondblclick="if(null == event.returnValue) event.returnValue=this" oncontextmenu="if(null == event.returnValue) event.returnValue=this">
                    <A onclick="return(false)" href="#" tabindex="-1">
                      <TD nowrap="true" class="vwItemSep vwItmTd">
                        <xsl:attribute name="colspan">
                          <xsl:choose>
                            <xsl:when test="a:propstat/a:prop[prio ='0' or prio ='2' or fattach ='1']">17</xsl:when>
                            <xsl:otherwise>20</xsl:otherwise>
                          </xsl:choose>
                        </xsl:attribute>
                        <xsl:value-of select="a:propstat/a:prop/subj"/>&#160;
                      </TD>
                      <xsl:if test="a:propstat/a:prop[prio ='0' or prio ='2' or fattach ='1']">
                        <TD colspan="3" nowrap="true" align="right" class="vwItemSep RTLrevA" style="cursor:default;padding:0px 3px;">
                          <xsl:choose>
                            <xsl:when test="a:propstat/a:prop/prio[.='0']">
                              <IMG width="10" height="16" src="/exchweb/img/imp-low.gif"/>
                            </xsl:when>
                            <xsl:when test="a:propstat/a:prop/prio[.='2']">
                              <IMG width="10" height="16" src="/exchweb/img/imp-high.gif"/>
                            </xsl:when>
                          </xsl:choose>
                          <xsl:choose>
                            <xsl:when test="a:propstat/a:prop/fattach[.='1']">
                              <IMG width="16" height="16" src="/exchweb/img/icon-paperclip.gif"/>
                            </xsl:when>
                          </xsl:choose>
                        </TD>
                      </xsl:if>
                    </A>
                  </TR>
                </xsl:for-each>
              </TABLE>
            </DIV>
          </TD>
        </TR>
      </TABLE>
    </xsl:template>
  </group>
</view>