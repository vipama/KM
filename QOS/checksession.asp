<%
			page=request("page")
			Select Case Page				

			Case "data_detail"%><!--#include file="data_detail.asp"--><%
			Case Else %><!--#include file="home.asp"--><%
		
			end select
%>