<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>사용자 관리</title>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>👥 사용자 관리</h2>
    <a href="${pageContext.request.contextPath}/user/form" class="btn btn-primary">+ 사용자 등록</a>
</div>

<c:if test="${not empty msg}"><div class="flash">${msg}</div></c:if>

<div class="table-wrap">
<table>
    <thead>
        <tr>
            <th>아이디</th><th>이름</th><th>권한</th><th>등록일</th>
            <th style="width:180px;">관리</th>
        </tr>
    </thead>
    <tbody>
        <c:forEach var="u" items="${users}">
            <tr>
                <td class="center">${u.userId}</td>
                <td>${u.userName}</td>
                <td class="center">
                    <c:choose>
                        <c:when test="${u.role == 'ADMIN'}"><span class="badge badge-admin">ADMIN</span></c:when>
                        <c:otherwise><span class="badge badge-user">USER</span></c:otherwise>
                    </c:choose>
                </td>
                <td>${u.createdAt}</td>
                <td class="center">
                    <a href="${pageContext.request.contextPath}/user/form?userId=${u.userId}" class="btn btn-edit">수정</a>
                    <c:if test="${u.userId != sessionScope.loginUser.userId}">
                        <form action="${pageContext.request.contextPath}/user/delete" method="post" style="display:inline;"
                              onsubmit="return confirm('정말 삭제하시겠습니까?');">
                            <input type="hidden" name="userId" value="${u.userId}">
                            <button type="submit" class="btn btn-delete">삭제</button>
                        </form>
                    </c:if>
                </td>
            </tr>
        </c:forEach>
        <c:if test="${empty users}">
            <tr><td colspan="5" class="empty">등록된 사용자가 없습니다.</td></tr>
        </c:if>
    </tbody>
</table>
</div>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
