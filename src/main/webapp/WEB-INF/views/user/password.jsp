<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>비밀번호 변경</title>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>🔑 비밀번호 변경</h2>
</div>

<div class="form-box" style="max-width:480px;">
    <form action="${pageContext.request.contextPath}/user/password" method="post" autocomplete="off">
        <label>현재 비밀번호</label>
        <input type="password" name="currentPw" required autofocus autocomplete="current-password">

        <label>새 비밀번호 (4자 이상)</label>
        <input type="password" name="newPw" minlength="4" required autocomplete="new-password">

        <label>새 비밀번호 확인</label>
        <input type="password" name="newPwConfirm" minlength="4" required autocomplete="new-password">

        <div class="btns">
            <button type="submit" class="btn btn-save">변경</button>
            <a href="${pageContext.request.contextPath}/main" class="btn btn-back">취소</a>
        </div>

        <c:if test="${not empty error}"><div class="error">${error}</div></c:if>
        <c:if test="${not empty msg}"><div class="ok">${msg}</div></c:if>
    </form>
</div>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
