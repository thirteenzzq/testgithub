<!--#include virtual="/etc/dbconn.asp"-->
<!--#include virtual="/etc/sqlInject.asp"-->
<!--#include virtual="/etc/isLogin.asp"-->
<!--#include virtual="/sysmgr/func.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
	<link href="/Financial/newCaiwu/css/style.css" rel="stylesheet">
	<link rel="stylesheet" href="/etc/c.css">
	<link rel="stylesheet" href="/bootstrap/css/bootstrap.min.css">
	<link rel="stylesheet" href="/bootstrap/css/bootstrap-theme.min.css">
	<script src="/Financial/newCaiwu/js/jquery-1.10.2.min.js"></script>
	<script src="/Financial/newCaiwu/js/bootstrap.min.js"></script>
  <!-- HTML5 shim and rspond.js IE8 support of HTML5 elements and media queries -->
  <!--[if lt IE 9]>
  <script src="js/html5shiv.js"></script>
  <script src="js/rspond.min.js"></script>
  <![endif]-->
</head>
<body>
<br>
sfsdfsdfsdfsdfsd
sdfsdfsd
ʤ�ฺ��
<ol class="breadcrumb">
    <li><a href="/main.asp">�����ۺϹ���ϵͳ</a></li>
    <li><a href="#">��Ч��ع���</a></li>
    <li class="active">�ٶȴ�������</li>
</ol>
 <!--body wrapper start-->
<div class="container-fluid">
<section class="panel">
	<header class="panel-heading custom-tab turquoise-tab">
		<ul class="nav nav-pills">
		<li class="active" style="text-align: center;width: 18%;">
			<a href="jixiaozhidu.asp">��Ч�ƶ�</a>
		</li>
		<li  style="text-align: center;width: 18%;">
			<a href="dailydata.asp">���̼�ر���д</a>
		</li>
		<li  style="text-align: center;width: 18%;">
			<a href="dailyview.asp">���̼�ر�鿴</a>
		</li>
		<li class="" style="text-align: center;width: 18%;">
			<a href="#" class=" dropdown-toggle"  data-toggle="dropdown">
			�ձ�����
			<span class="caret"></span>
			</a>
			<ul class="dropdown-menu dropdown-menu-usermenu pull-right">
				<li>
					<a href="dailyribao.asp">
					<span style="color:#3ed29a">
					<i class="glyphicon glyphicon-pencil"></i>
					ÿ���ձ���д
					</span>
					</a>
				</li>
				<li>
					<a href="../ribao.asp" target="view_window">
					<span style="color:#3ed29a">
					<i class="glyphicon glyphicon-search"></i>
					ÿ���ձ��鿴
					</span>
					</a>
				</li>
			</ul>
		</li>
		<li class="" style="text-align: center;width: 18%;">
			<a  href="#" class=" dropdown-toggle"  data-toggle="dropdown">
			��ʷ����
			<span class="caret"></span>
			</a>
			<ul class="dropdown-menu dropdown-menu-usermenu pull-right">
			<li>
			<a href="dcallvcutds.asp" target="view_window">
			<span style="color:#3ed29a">
			<i class="glyphicon glyphicon-pencil"></i>
			�����������¿�
			</span>
			</a>
			</li>
			<li>
			<a href="dcallvcutdsgd.asp" target="view_window">
			<span style="color:#3ed29a">
			<i class="glyphicon glyphicon-search"></i>
			�㶫�������¿�
			</span>
			</a>
			</li>
			</ul>
		</li>
		</ul>
	</header>
</section>
<ul class="nav nav-tabs" role="tablist">
    <li role="presentation" class="active"><a href="#shangwubu" aria-controls="shangwubu" role="tab" data-toggle="tab">�������</a></li>
	<li role="presentation" class=""><a href="#zongjingli" aria-controls="zongjingli" role="tab" data-toggle="tab">�ܾ���Ч</a></li>
	<li role="presentation" class=""><a href="#wangzhan" aria-controls="wangzhan" role="tab" data-toggle="tab">��վ�������</a></li>
	<li role="presentation" class=""><a href="#dake" aria-controls="dake" role="tab" data-toggle="tab">��ͻ��Ŷӿ���</a></li>
</ul>
<div class="tab-content">
<div role="tabpanel" class="tab-pane active" id="shangwubu">
<div style="position:relative; width:100%; height:100%; overflow:auto">
<table class="tables" >
<tr align="left" bgcolor="#edffee" height=40>
	<td style="padding:5 3 5 3;line-height:140%;width:70%">
		���⣺�������񲿿����ƶ�
		&nbsp;<a href="zhidu.xlsx">����</a>
		<br />
		���ߣ�������&nbsp;&nbsp;<br />
		����ʱ�䣺2017-6-7 8:49:45<br />
		���¸���ʱ�䣺2021-12-15 15:55:45<br />
		<%
		'7.11--1.�������¿����ѽ������ɹҹ���2.������Ʒ��չʾ���Ʒ����ɣ������ҹ�
		'7.17--1.�޸���Ʒ��չʾ���Ʒ����ɣ������ҹ�
		'7.29--1.�޸������ܼಿ���µ����𷣵���
		'2020.0206--�ҹ����ѽ��
		'2020.3.24��������������׼
		'2021.9.20�����˷�cpc���Ʒ����������������
		
		%>
	</td>
</tr>
<tr>
<td>
<iframe frameborder="0" width="1180px" height="700px" src="dasouzhidu.asp?v=<%=now()%>"></iframe>
<!--a href="/worktable/dasou/dasouzhidu.png" target="_blank"><img src="/worktable/dasou/dasouzhidu.png?v=<%=now()%>"></a-->
</td>
</tr>
</table>
</div>
</div>

<div role="tabpanel" class="tab-pane" id="zongjingli">
<div style="position:relative; width:100%; height:100%; overflow:auto">
<table class="tables" >
<tr align="left" bgcolor="#edffee" height=40>
	<td style="padding:5 3 5 3;line-height:140%;width:70%">
		���⣺�ܾ���Ч�����ƶ�
		&nbsp;<a href="zongjingli.xlsx">����</a>
		<br />
		���ߣ�������&nbsp;&nbsp;<br />
		����ʱ�䣺2018-10-10 9:49:45<br />
		���¸���ʱ�䣺2021-04-14 10:26:45<br />
		<%
		'7.17--1.�����˽�������
			'2.������98%-100%��һ���Ľ���
		%>
	</td>
</tr>
<tr>
<td>
<a href="/worktable/dasou/zongjingli.png" target="_blank"><img src="/worktable/dasou/zongjingli.png?v=<%=now()%>"></a>
</td>
</tr>
</table>
</div>
</div>
<div role="tabpanel" class="tab-pane" id="wangzhan">
<div style="position:relative; width:100%; height:100%; overflow:auto">
<table class="tables" >
<tr align="left" bgcolor="#edffee" height=40>
	<td style="padding:5 3 5 3;line-height:140%;width:70%">
		���⣺��վ������ɿ����ƶ�
		&nbsp;<a href="wangzhan.xls">����</a>
		<br/>
		���ߣ�������&nbsp;&nbsp;<br />
		����ʱ�䣺2017-7-19 8:49:45<br />
		���¸���ʱ�䣺2020-06-24 18:14:45<br />
		<%
		'7.18--1.�����������������ɱ��������ֹ���
		'11.27--1.��������վ��ɹ������ֹ���
		%>
	</td>
</tr>
<tr>
<td>
<a href="/worktable/dasou/wangzhan.png" target="_blank"><img src="/worktable/dasou/wangzhan.png?v=<%=now()%>"></a>
</td>
</tr>
</table>
</div>
</div>

<div role="tabpanel" class="tab-pane" id="dake">
<div style="position:relative; width:100%; height:100%; overflow:auto">
<table class="tables" >
<tr align="left" bgcolor="#edffee" height=40>
	<td style="padding:5 3 5 3;line-height:140%;width:70%">
		���⣺��ͻ��Ŷӿ���ִ�б�׼
		&nbsp;<a href="dake.xlsx">����</a>
		<br/>
		���ߣ�������&nbsp;&nbsp;<br />
		����ʱ�䣺2021-4-6 13:46:45<br />
		���¸���ʱ�䣺2021-4-6 13:46:45<br />
		<%
		%>
	</td>
</tr>
<tr>
<td>
<a href="/worktable/dasou/dake.png" target="_blank"><img src="/worktable/dasou/dake.png?v=<%=now()%>"></a>
</td>
</tr>
</table>
</div>
</div>

</div>

</div>
</body>
</html>