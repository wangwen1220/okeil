	<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="35" valign="middle"> 
	  <input type="button" name="action2" onClick="javascript:location.href='?gb=1';" value="英文版">
	  <input type="button" name="action2" onClick="javascript:location.href='?gb=0';" value="西班牙文">
      &nbsp;&nbsp;<strong><font color="#FF6600">&nbsp;您现在管理的是： 
	  <% if gb="0" then%> 西班牙文  <%end if%>
	  <% if gb="1" then%> 英文版 <%end if%>
	  </font></strong></td>
        </tr>
      </table>
