<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>Simple Excel Exporter</title>
</head>
<body>
<h1>Simple Excel Exporter</h1>
<form action="${pageContext.request.contextPath}/export" method="post">
    <label> Row Count:
        <input type="text" name="rowCount" value="1000">
    </label>
    <input type="submit" name="Export Excel">
</form>
</body>
</html>