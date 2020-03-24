<%--
  Created by IntelliJ IDEA.
  User: XinjianLi
  Date: 2020/2/25
  Time: 12:03
  To change this template use File | Settings | File Templates.
--%>
<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<html>
<head>
    <title>Title</title>
</head>
<body>
<form action="importExcel" method="post" enctype="multipart/form-data">
    选择图片:<input type="file" name="file" /> <br>
    <input type="submit" value="上传">
</form>

<form action="exportExcel" method="post">
    <input type="submit" value="下载">
</form>
</body>
</html>
