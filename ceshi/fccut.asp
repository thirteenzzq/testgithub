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
<span style="font-size:26px"><strong>百度房产项目业绩进度报表(<%=Dates%>)</strong></span>
</div>
<br>
<div style="text-align:left">
<%if hour(now())>=10 then %>
<span style="font-size:18px">截图时间：<span id="current-time" style="color:red"></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:red">★抱歉，您的截图时间已经超过10点，请在群里发200元红包作为惩罚，如有特殊情况，请在群内说明!</span></span>
<%else%>
<span style="font-size:18px">截图时间：<span id="current-time"></span></span>
<%end if %>
</div>
<br>
<table style="margin:0px;padding:0px;" class="table table-bordered table-hover table-condensed" >
 <thead>
  <tr >
	<th width="110px" rowspan="3" class="warning">城市</th>
	<th colspan="7" class="success">楼盘入库基础信息</th>
	<th colspan="6" class="info">线索数据</th>
	<th colspan="8" class="danger">当日数据</th>
	<th colspan="10" class="success">月度数据</th>
	<th colspan="2" rowspan="2" class="info">季度数据</th>
  </tr>
  <tr style="font-size:11px">
    <th rowspan="2" class="success">地区<br>楼盘数</th>
    <th rowspan="2" class="success">在售数</th>
    <th rowspan="2" class="success">待售数</th>
    <th rowspan="2" class="success">售罄数</th>
    <th rowspan="2" class="success">上架<br>楼盘数</th>
	<th rowspan="2" class="success">下架<br>楼盘数</th>
	<th rowspan="2" class="success">楼盘<br>合格数</th>
	
	<th class="info" colspan="3">当日线索</th>
	<th class="info" colspan="3">当月线索</th>
	
	<th class="danger" colspan="2">完成情况</th>
	<th class="danger" rowspan="2">探盘数</th>
	<th class="danger" rowspan="2">意向<br>楼盘数</th>
	<th class="danger" colspan="4">业务推荐客户数<br><span style="color:red">（周指标提示：徐州4，镇江3）</span></th>
	
	<th class="success" colspan="2">任务值</th>
	<th class="success" colspan="2">完成情况</th>
	<th class="success" rowspan="2">探盘数</th>
	<th class="success" rowspan="2">意向<br>楼盘数</th>
	<th class="success" colspan="4">业务推荐客户数<br><span style="color:red">（月指标：徐州16 ，镇江12）</span></th>
	</tr>
  <tr style="font-size:11px">
	<th class="info">线索数</th>
	<th class="info">跟进数</th>
	<th class="info">有效数</th>
	<th class="info">线索数</th>
	<th class="info">跟进数</th>
	<th class="info">有效数</th>
	
	<th class="danger">客户数</th>
	<th class="danger">金额</th>
	<th class="danger">房产<br>会员</th>
	<th class="danger">搜索</th>
	<th class="danger">信息流<br>及其他</th>
	<th class="danger">合计</th>
	
	<th class="success">客户数</th>
	<th class="success">金额</th>
	<th class="success">客户数</th>
	<th class="success">金额</th>
	<th class="success">房产<br>会员</th>
	<th class="success">搜索</th>
	<th class="success">信息流<br>及其他</th>
	<th class="success">合计</th>
	
	<th class="info">客户数</th>
	<th class="info">金额</th>
	
  </tr>
  </thead>
   <tbody>
<%
conn.open
sql="select * from OA_dailydsdata where ds_company='百度房产' and ds_date='"&Dates&"' order by ds_id"
Set rs=server.CreateObject("adodb.recordset")
rs.open sql,conn,1,1
	if rs.EOF then
		response.write "<p>百度房产未提交"
	else
	count=conn.execute ("select count(1) from OA_dailydsdata where ds_company='百度房产' and ds_date='"&Dates&"'")(0)
	
	for i=1 to count%>
	<tr <%if i=3 then%>class="danger"<%end if%>>
	
		<td><%=rs("ds_area")%></td>
		<%for a=1 to 10 %>
				<td><%=rs("ds_data"&a&"")%></td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data8) from OA_dailydsdata where ds_company='百度房产' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data9) from OA_dailydsdata where ds_company='百度房产' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data10) from OA_dailydsdata where ds_company='百度房产' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%for a=11 to 20 %>
				<td><%=rs("ds_data"&a&"")%></td>
		<%next%>
		<%for a=11 to 18 %>
				<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='百度房产' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data11) from OA_dailydsdata where ds_company='百度房产' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data12) from OA_dailydsdata where ds_company='百度房产' and ds_area='"&rs("ds_area")&"' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
	</tr>
	<%rs.movenext
	next%>
	<tr class="danger">
	
		<td>合计</td>
		<%for a=1 to 10 %>
		<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='百度房产'  and ds_date='"&Dates&"'")(0)%>
		</td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data8) from OA_dailydsdata where ds_company='百度房产' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data9) from OA_dailydsdata where ds_company='百度房产' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data10) from OA_dailydsdata where ds_company='百度房产' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%for a=11 to 20 %>
		<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='百度房产'  and ds_date='"&Dates&"'")(0)%>
		</td>
		<%next%>
		<%for a=11 to 18 %>
				<td><%=conn.execute("select sum(ds_data"&a&") from OA_dailydsdata where ds_company='百度房产' and year(ds_date)=year('"&Dates&"') and month(ds_date)=month('"&Dates&"') ")(0)%></td>
		<%next%>
		<td><%=conn.execute("select sum(ds_data11) from OA_dailydsdata where ds_company='百度房产' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
		<td><%=conn.execute("select sum(ds_data12) from OA_dailydsdata where ds_company='百度房产' and year(ds_date)=year('"&Dates&"') and datepart(quarter,ds_date)="&DatePart("q",Dates)&"  and ds_date<='"&Dates&"' ")(0)%></td>
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

