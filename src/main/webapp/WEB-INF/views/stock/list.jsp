<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<%@ taglib prefix="fmt" uri="jakarta.tags.fmt" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>현재 재고 조회</title>
    <style>
        /* 헤더 가운데 정렬 + 단위 부제 */
        .stock-table thead tr.unit-row th {
            font-size: 11px; color: var(--gray-500);
            font-weight: 500; letter-spacing: 0;
            border-top: 0;
            padding-top: 2px; padding-bottom: 8px;
        }
        .stock-table thead tr.title-row th {
            padding-bottom: 4px;
        }
        /* 매진(현재재고 0) 표시 */
        tr.sold-out { color: var(--gray-400); }
        tr.sold-out td.now-stock { color: var(--gray-400); }
        .now-stock { font-weight: 700; color: var(--color-primary); }
    </style>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>📊 현재 재고 조회</h2>
    <a href="${pageContext.request.contextPath}/stock/excel?onlyAvailable=${onlyAvailable}"
       class="btn btn-primary">📥 엑셀 다운로드</a>
</div>

<!-- 간단 필터: 현재재고 있는 것만 -->
<form action="${pageContext.request.contextPath}/stock/list" method="get" class="search-bar">
    <div class="filters">
        <div class="filter-group">
            <label class="checkbox-label">
                <input type="checkbox" name="onlyAvailable" value="true"
                       <c:if test="${onlyAvailable}">checked</c:if>>
                현재재고가 있는 것만
            </label>
        </div>
    </div>
    <div class="actions">
        <button type="submit" class="btn btn-search">조회</button>
    </div>
</form>

<div class="table-wrap">
<table class="stock-table">
    <thead>
        <tr class="title-row">
            <th style="width:60px;">No</th>
            <th style="width:200px;">입고번호</th>
            <th style="width:110px;">크기</th>
            <th style="width:110px;">초기재고</th>
            <th style="width:110px;">총입고량</th>
            <th style="width:110px;">총출고량</th>
            <th style="width:120px;">현재재고</th>
        </tr>
        <tr class="unit-row">
            <th></th>
            <th></th>
            <th>(μm)</th>
            <th>(${not empty stocks ? '' : ''}포장)</th>
            <th>(포장)</th>
            <th>(포장)</th>
            <th>(포장)</th>
        </tr>
    </thead>
    <tbody>
        <c:forEach var="s" items="${stocks}" varStatus="st">
            <tr class="${s.remainingQty == 0 ? 'sold-out' : ''}">
                <td class="center">${st.index + 1}</td>
                <td class="center">${s.inboundId}</td>
                <td class="num">
                    <c:if test="${not empty s.sizeUm}"><fmt:formatNumber value="${s.sizeUm}" pattern="0.000"/></c:if>
                </td>
                <td class="num"><fmt:formatNumber value="${s.packageCount}" pattern="#,##0"/></td>
                <td class="num"><fmt:formatNumber value="${s.packageCount}" pattern="#,##0"/></td>
                <td class="num"><fmt:formatNumber value="${s.outboundQty}" pattern="#,##0"/></td>
                <td class="num now-stock"><fmt:formatNumber value="${s.remainingQty}" pattern="#,##0"/></td>
            </tr>
        </c:forEach>
        <c:if test="${empty stocks}">
            <tr><td colspan="7" class="empty">조회된 재고가 없습니다.</td></tr>
        </c:if>
    </tbody>
</table>
</div>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
