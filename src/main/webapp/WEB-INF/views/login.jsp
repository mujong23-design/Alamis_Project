<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>로그인 - 재고관리</title>
    <link rel="stylesheet" href="${pageContext.request.contextPath}/css/style.css">
</head>
<body class="login-page">
    <div class="login-box">
        <h2>📦 재고관리 시스템</h2>
        <form action="${pageContext.request.contextPath}/login" method="post" autocomplete="off">
            <input type="text"     name="userId" placeholder="아이디"   required autofocus autocomplete="off">
            <input type="password" name="userPw" placeholder="비밀번호" required          autocomplete="new-password">
            <button type="submit">로그인</button>
            <c:if test="${not empty error}">
                <div class="error">${error}</div>
            </c:if>
        </form>
    </div>
</body>
</html>
