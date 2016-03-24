<%@include file="/libs/foundation/global.jsp"%>
<%@ page import="javax.jcr.Node,java.util.Calendar,javax.jcr.Property,java.text.SimpleDateFormat,java.util.Locale" %>
<%
    Node form = resourceResolver.getResource(currentNode.getPath()).adaptTo(Node.class);
    boolean hasFile = false;
    String date = "";
    if (form.hasNode("group-report.xlsx")) {
        Node file = form.getNode("group-report.xlsx");
        hasFile = true;
        if (file.hasProperty("jcr:created")) {
            Property dateProp = file.getProperty("jcr:created");
            date = dateProp.getString();
        }
    }
%>
<html>
<head>
    <title>Group Enrollment Report</title>
    <meta http-equiv="Content-Type" content="text/html; utf-8" />
    <script src="/libs/cq/ui/resources/cq-ui.js" type="text/javascript"></script>
    <cq:includeClientLib categories="jquery"/>
</head>
<body topmargin="0" leftmargin="0">
<h1>Group Enrollment Report</h1>
<h4>View existing or create new group enrollment report for all users in CQ.</h4>
<p>
<form action="/bin/sixdglobal/groupenrollment" id="report" name="report" method="POST">
    <input type="hidden" name="savePath" id="savePath" value="/apps/sixdglobal/reports/group-enrollment/run" />
    <input type="button" id="send" value="Generate Report" />
</form>
<%if(hasFile){%>
<br/>
<input type="button" id="download" value="Download Report" onclick="window.open('/apps/sixdglobal/reports/group-enrollment/run/group-report.xlsx')" />
<br/>
<div style="padding-left: 4px;">Last generated on: <%=date%></div>
<%}%>
</p>
<script>
    $("#send").click(function () {
        var formData = $("#report").serialize();
        $.ajax({
            type: 'POST',
            url:'/bin/sixdglobal/groupenrollment',
            data:formData,
            success: function(msg){
                alert(msg); //display the data returned by the servlet
                setTimeout(function(){
                    location.reload();
                }, 5000);
            }
        });
    });
</script>
</body>
</html>