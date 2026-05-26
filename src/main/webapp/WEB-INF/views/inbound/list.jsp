<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<%@ taglib prefix="fmt" uri="jakarta.tags.fmt" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>입고 목록</title>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>📥 입고 목록</h2>
    <div style="display:flex; gap:8px;">
        <button type="button" class="btn btn-secondary" id="btnLabel">🏷️ 선택 라벨 인쇄</button>
        <a href="${pageContext.request.contextPath}/inbound/form" class="btn btn-primary">+ 입고 등록</a>
    </div>
</div>

<c:if test="${not empty msg}"><div class="flash">${msg}</div></c:if>
<c:if test="${not empty error}">
    <div class="flash" style="background:var(--color-danger-light); border-left-color:var(--color-danger); color:var(--color-danger);">${error}</div>
</c:if>

<!-- 검색 바 -->
<form action="${pageContext.request.contextPath}/inbound/list" method="get" class="search-bar">
    <div class="filters">
        <div class="filter-group">
            <label class="checkbox-label">
                <input type="checkbox" name="useDate" value="true" id="useDateCheck"
                       <c:if test="${useDate}">checked</c:if>>
                생산일자
            </label>
            <input type="date" id="startDate" name="startDate"
                   value="${not empty startDate ? startDate : today}"
                   <c:if test="${!useDate}">disabled</c:if>>
            <span class="date-separator">~</span>
            <input type="date" id="endDate" name="endDate"
                   value="${not empty endDate ? endDate : today}"
                   <c:if test="${!useDate}">disabled</c:if>>
        </div>
    </div>
    <div class="actions">
        <c:if test="${useDate}">
            <a href="${pageContext.request.contextPath}/inbound/list" class="btn btn-reset">초기화</a>
        </c:if>
        <button type="submit" class="btn btn-search">조회</button>
    </div>
</form>

<div class="table-wrap">
<table>
    <thead>
        <tr>
            <th style="width:40px;"><input type="checkbox" id="selectAll" title="전체 선택"></th>
            <th style="width:50px;">#</th>
            <th style="width:180px;">제품번호</th>
            <th style="width:120px;">생산일자</th>
            <th style="width:110px;">포장단위(kg)</th>
            <th style="width:90px;">포장개수</th>
            <th style="width:90px;">출고량</th>
            <th style="width:90px;">잔여</th>
            <th style="width:90px;">크기(μm)</th>
            <th style="width:100px;">출력</th>
            <th style="width:200px;">비고</th>
            <th style="width:160px;">관리</th>
        </tr>
    </thead>
    <tbody>
        <c:forEach var="i" items="${inbounds}" varStatus="st">
            <tr>
                <td class="center">
                    <input type="checkbox" name="inboundIds" value="${i.inboundId}"
                           data-count="${i.packageCount}"
                           data-printed="${i.printedYn}">
                </td>
                <td class="center">${st.index + 1}</td>
                <td class="center">${i.inboundId}</td>
                <td class="center">${i.productionDate}</td>
                <td class="num"><fmt:formatNumber value="${i.packageUnitKg}" pattern="0.###"/></td>
                <td class="num"><fmt:formatNumber value="${i.packageCount}" pattern="#,##0"/></td>
                <td class="num"><fmt:formatNumber value="${i.outboundQty}" pattern="#,##0"/></td>
                <td class="num" style="font-weight:600;color:var(--color-primary);">
                    <fmt:formatNumber value="${i.remainingQty}" pattern="#,##0"/>
                </td>
                <td class="num">
                    <c:if test="${not empty i.sizeUm}"><fmt:formatNumber value="${i.sizeUm}" pattern="0.###"/></c:if>
                </td>
                <td class="center">
                    <c:choose>
                        <c:when test="${i.printedYn eq 'Y'}">
                            <span style="display:inline-block; white-space:nowrap; background:var(--color-success); color:#fff; padding:2px 10px; border-radius:10px; font-size:11px; font-weight:600;">출력완료</span>
                        </c:when>
                        <c:otherwise>
                            <span style="display:inline-block; white-space:nowrap; background:var(--gray-200); color:var(--gray-600); padding:2px 10px; border-radius:10px; font-size:11px;">미출력</span>
                        </c:otherwise>
                    </c:choose>
                </td>
                <td>${i.remark}</td>
                <td class="center">
                    <a href="${pageContext.request.contextPath}/inbound/form?id=${i.inboundId}" class="btn btn-edit">수정</a>
                    <c:if test="${i.outboundQty == 0}">
                        <form action="${pageContext.request.contextPath}/inbound/delete" method="post" style="display:inline;"
                              onsubmit="return confirm('정말 삭제하시겠습니까?');">
                            <input type="hidden" name="inboundId" value="${i.inboundId}">
                            <button type="submit" class="btn btn-delete">삭제</button>
                        </form>
                    </c:if>
                </td>
            </tr>
        </c:forEach>
        <c:if test="${empty inbounds}">
            <tr><td colspan="12" class="empty">입고 내역이 없습니다.</td></tr>
        </c:if>
    </tbody>
</table>
</div>

<script>
    const ctx = '${pageContext.request.contextPath}';

    const dateCheck = document.getElementById('useDateCheck');
    const startDate = document.getElementById('startDate');
    const endDate   = document.getElementById('endDate');
    if (dateCheck) {
        dateCheck.addEventListener('change', () => {
            startDate.disabled = !dateCheck.checked;
            endDate.disabled   = !dateCheck.checked;
        });
    }

    // 전체 선택
    const selectAll = document.getElementById('selectAll');
    if (selectAll) {
        selectAll.addEventListener('change', e => {
            document.querySelectorAll('input[name="inboundIds"]').forEach(cb => {
                cb.checked = e.target.checked;
            });
        });
    }

    // 라벨 인쇄
    document.getElementById('btnLabel').addEventListener('click', () => {
        const checked = document.querySelectorAll('input[name="inboundIds"]:checked');
        if (checked.length === 0) {
            alert('라벨 인쇄할 입고건을 1개 이상 선택해주세요.');
            return;
        }

        // 1) 이미 출력된 적이 있는 입고건이 있는지 검사 → 재출력 컨펌
        const alreadyPrinted = Array.from(checked).filter(cb => cb.dataset.printed === 'Y');
        if (alreadyPrinted.length > 0) {
            const ids = alreadyPrinted.map(cb => cb.value).join('\n');
            if (!confirm('이미 출력했던 이력이 있습니다.\n\n' + ids + '\n\n그래도 출력하시겠습니까?')) {
                return;
            }
        }

        // 2) 총 라벨 수 안내
        let totalLabels = 0;
        checked.forEach(cb => {
            const cnt = parseInt(cb.dataset.count || '0');
            if (!isNaN(cnt)) totalLabels += cnt;
        });
        if (!confirm('선택한 입고 ' + checked.length + '건 → 총 ' + totalLabels + '개 라벨이 생성됩니다. 계속하시겠습니까?')) {
            return;
        }

        const idCsv = Array.from(checked).map(cb => cb.value).join(',');
        window.open(ctx + '/inbound/label?ids=' + encodeURIComponent(idCsv), '_blank', 'width=900,height=900');
    });
</script>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
