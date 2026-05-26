<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c"  uri="jakarta.tags.core" %>
<%@ taglib prefix="fn" uri="jakarta.tags.functions" %>
<c:set var="ctx" value="${pageContext.request.contextPath}"/>
<c:set var="uri" value="${pageContext.request.requestURI}"/>

<link rel="stylesheet" href="${ctx}/css/style.css">

<div class="app">
    <!-- ===== Sidebar ===== -->
    <aside class="sidebar">
        <div class="brand">📦 재고관리 시스템</div>
        <nav class="nav">
            <a href="${ctx}/main"
               class="nav-item ${fn:endsWith(uri, '/main') ? 'active' : ''}">🏠 홈</a>

            <div class="nav-section">메뉴</div>
            <a href="${ctx}/inbound/list"
               class="nav-item ${fn:contains(uri, '/inbound') ? 'active' : ''}">입고 관리</a>
            <a href="${ctx}/outbound/list"
               class="nav-item ${fn:contains(uri, '/outbound') ? 'active' : ''}">출고 관리</a>
            <a href="${ctx}/stock/list"
               class="nav-item ${fn:contains(uri, '/stock') ? 'active' : ''}">현재 재고 조회</a>

            <c:if test="${sessionScope.loginUser.role == 'ADMIN'}">
                <div class="nav-section">관리자</div>
                <a href="${ctx}/user/list"
                   class="nav-item ${fn:contains(uri, '/user/list') or fn:contains(uri, '/user/form') ? 'active' : ''}">사용자 관리</a>
            </c:if>
        </nav>
        <div class="sidebar-footer">v0.2.0</div>
    </aside>

    <!-- ===== Main Area ===== -->
    <div class="main-area">
        <div class="topbar">
            <div class="topbar-title">
                <c:choose>
                    <c:when test="${fn:endsWith(uri, '/main')}">대시보드</c:when>
                    <c:when test="${fn:contains(uri, '/inbound/list')}">입고 목록</c:when>
                    <c:when test="${fn:contains(uri, '/inbound/form')}">입고 등록 / 수정</c:when>
                    <c:when test="${fn:contains(uri, '/outbound/list')}">출고 목록</c:when>
                    <c:when test="${fn:contains(uri, '/outbound/form')}">출고 등록</c:when>
                    <c:when test="${fn:contains(uri, '/stock/list')}">현재 재고 조회</c:when>
                    <c:when test="${fn:contains(uri, '/user/list')}">사용자 목록</c:when>
                    <c:when test="${fn:contains(uri, '/user/form')}">사용자 등록 / 수정</c:when>
                    <c:when test="${fn:contains(uri, '/user/password')}">비밀번호 변경</c:when>
                </c:choose>
            </div>
            <div class="topbar-user">
                <c:if test="${not empty sessionScope.loginUser}">
                    <span>
                        ${sessionScope.loginUser.userName} 님
                        <c:if test="${sessionScope.loginUser.role == 'ADMIN'}">
                            <span class="badge-admin">ADMIN</span>
                        </c:if>
                    </span>
                    <a href="${ctx}/user/password">비밀번호 변경</a>
                    <a href="${ctx}/logout">로그아웃</a>
                </c:if>
            </div>
        </div>

        <main class="content">
