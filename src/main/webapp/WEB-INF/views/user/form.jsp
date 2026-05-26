<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>사용자 등록/수정</title>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>${isEdit ? '사용자 수정' : '사용자 등록'}</h2>
</div>

<div class="form-box">
    <form action="${pageContext.request.contextPath}/user/save" method="post" autocomplete="off">
        <input type="hidden" name="isEdit" value="${isEdit}">

        <label>아이디</label>
        <input type="text" name="userId" value="${user.userId}"
               <c:if test="${isEdit}">readonly</c:if> required>

        <label>이름</label>
        <input type="text" name="userName" value="${user.userName}" required>

        <label>권한</label>
        <select name="role">
            <option value="USER"  ${user.role == 'USER'  ? 'selected' : ''}>USER (일반사용자)</option>
            <option value="ADMIN" ${user.role == 'ADMIN' ? 'selected' : ''}>ADMIN (관리자)</option>
        </select>

        <label>비밀번호 ${isEdit ? '(변경 시에만 입력)' : ''}</label>
        <input type="password" name="rawPw" autocomplete="new-password" <c:if test="${!isEdit}">required</c:if>>
        <div class="hint">비밀번호는 BCrypt로 암호화되어 저장됩니다.</div>

        <div class="btns">
            <button type="submit" class="btn btn-save">저장</button>
            <a href="${pageContext.request.contextPath}/user/list" class="btn btn-back">목록</a>
        </div>
    </form>
</div>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
