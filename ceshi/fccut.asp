<!--#include virtual="/etc/dbconn.asp"-->
<!--#include virtual="/etc/sqlInject.asp"-->
<!--#include virtual="/etc/isLogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<link rel="stylesheet" href="/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" href="/bootstrap/css/bootstrap-theme.min.css">
<script type="text/javascript" src="/js/datepicker/WdatePicker.js"></script>
<script src="/js/jquery-2.1.3.min.js"></script>
<style>
.table th, .table td {
text-align: center; vertical-align:middle !important;
}
</style>
<script>
setInterval(function() {
    var now = (new Date()).toLocaleString();
    $('#current-time').text(now);
}, 1000);
</script>
</head>
<body>
<% Dates=Trim(request.querystring("dated"))%>
<br>
<div style="width:1800px">
<div style="text-align:center">
<span style="font-size:26px"><strong>�ٶȷ�����Ŀҵ�����ȱ���(<%=Dates%>)</strong></span>
</div>
<br>
<div style="text-align:left">
<%if hour(now())>=10 then %>
<span style="font-size:18px">��ͼʱ�䣺<span id="current-time" style="color:red"></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:red">�ﱧǸ�����Ľ�ͼʱ���Ѿ�����10�㣬����Ⱥ�﷢200Ԫ�����Ϊ�ͷ��������������������Ⱥ��˵��!</span></span>
<%else%>
<span style="font-size:18px">��ͼʱ�䣺<span id="current-time"></span></span>
<%end if %>
</div>
<br>
<table style="margin:0px;padding:0px;" class="table table-bordered table-hover table-condensed" >
 <thead>
  <tr >
	<th width="110px" rowspan="3" class="warning">����</th>
	<th colspan="7" class="success">¥����������Ϣ</th>
	<th colspan="6" class="info">��������</th>
	<th colspan="8" class="danger">��������</th>
	<th colspan="10" class="success">�¶�����</th>
	<th colspan="2" rowspan="2" class="info">��������</th>
  </tr>
  <tr style="font-size:11px">
    <th rowspan="2" class="success">����<br>¥����</th>
    <th rowspan="2" class="success">������</th>
    <th rowspan="2" class="success">������</th>
    <th rowspan="2" class="success">������</th>
    <th rowspan="2" class="success">�ϼ�<br>¥����</th>
	<th rowspan="2" class="success">�¼�<br>¥����</th>
	<th rowspan="2" class="success">¥��<br>�ϸ���</th>
	
	<th class="info" colspan="3">��������</th>
	<th class="info" colspan="3">��������</th>
	
	<th class="danger" colspan="2">������</th>
	<th class="danger" rowspan="2">̽����</th>
	<th class="danger" rowspan="2">����<br>¥����</th>
	<th class="danger" colspan="4">ҵ���Ƽ��ͻ���<br><span style="color:red">����ָ����ʾ������4����3��</span></th>
	
	<th class="success" colspan="2">����ֵ</th>
	<th class="success" colspan="2">������</th>
	<th class="success" rowspan="2">̽����</th>
	<th class="success" rowspan="2">����<br>¥����</th>
	<th class="success" colspan="4">ҵ���Ƽ��ͻ���<br><span style="color:red">����ָ�꣺����16 ����12��</span></th>
	</tr>
  <tr style="font-size:11px">
	<th class="info">������</th>
	<th class="info">������</th>
	<th class="info">��Ч��</th>
	<th class="info">������</th>
	<th class="info">������</th>
	<th class="info">��Ч��</th>
	
	<th class="danger">�ͻ���</th>
	<th class="danger">���</th>
	<th class="danger">����<br>��Ա</th>
	<th class="danger">����</th>
	<th class="danger">��Ϣ��<br>������</th>
	<th class="danger">�ϼ�</th>
	
	<th class="success">�ͻ���</th>
	<th class="success">���</th>
	<th class="success">�ͻ���</th>
	<th class="success">���</th>
	<th class="success">����<br>��Ա</th>
	<th class="success">����</th>
	<th class="success">��Ϣ��<br>������</th>
	<th class="success">�ϼ�</th>
	
	<th class="info">�ͻ���</th>
	<th class="info">���</th>
	
  </tr>
  </thead>
   <tbody>
<%
conn.open
sql="select * from OA_dailydsdata where ds_company='�ٶȷ���' and ds_date='"&Dates&"' order by ds_id"
Set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
	if rs.EOF then
		response.write "<p>�ٶȷ���δ�ύ"
	else
	count=conn.execute ("select count(1) from OA_dailydsdata where ds_company='�ٶȷ���' and ds_date='"&Dates&"'")(0)
	
	for i=1 to count%>
	<tr <%if i=3 then%>class="danger"<%end if%>>
	
		<td><%=rs("ds_area")%></td>
		<%for a=1 to 10 %>
				<td><%=rs("ds_data"&a&"")%></td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data8) from OA_dailydsdata where ds_company='�ٶȷ���' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data9) from OA_dailydsdata where ds_company='�ٶȷ���' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data10) from OA_dailydsdata where ds_company='�ٶȷ���' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%for a=11 to 20 %>
				<td><%=rs("ds_data"&a&"")%></td>
		<%next%>
		<%for a=11 to 18 %>
				<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='�ٶȷ���' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data11) from OA_dailydsdata where ds_company='�ٶȷ���' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data12) from OA_dailydsdata where ds_company='�ٶȷ���' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
	</tr>
	<%rs.movenext
	next%>
	<tr class="danger">
	
		<td>�ϼ�</td>
		<%for a=1 to 10 %>
		<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='�ٶȷ���'  and ds_date='"&Dates&"'")(0)%>
		</td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data8) from OA_dailydsdata where ds_company='�ٶȷ���' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data9) from OA_dailydsdata where ds_company='�ٶȷ���' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data10) from OA_dailydsdata where ds_company='�ٶȷ���' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%for a=11 to 20 %>
		<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='�ٶȷ���'  and ds_date='"&Dates&"'")(0)%>
		</td>
		<%next%>
		<%for a=11 to 18 %>
				<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='�ٶȷ���' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data11) from OA_dailydsdata where ds_company='�ٶȷ���' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data12) from OA_dailydsdata where ds_company='�ٶȷ���' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
	</tr>
<%
end if
rs.close
set rs=nothing
conn.close%>
</tbody>
 </table>
 </div>
</body>
</html>

