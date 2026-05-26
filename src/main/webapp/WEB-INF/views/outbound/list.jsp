<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<%@ taglib prefix="fmt" uri="jakarta.tags.fmt" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>출고 목록</title>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>📤 출고 목록</h2>
    <a href="${pageContext.request.contextPath}/outbound/form" class="btn btn-primary">+ 출고 등록</a>
</div>

<c:if test="${not empty msg}"><div class="flash">${msg}</div></c:if>
<c:if test="${not empty error}">
    <div class="flash" style="background:var(--color-danger-light); border-left-color:var(--color-danger); color:var(--color-danger);">${error}</div>
</c:if>

<!-- 검색 -->
<form action="${pageContext.request.contextPath}/outbound/list" method="get" class="search-bar">
    <div class="filters">
        <div class="filter-group">
            <label class="checkbox-label">
                <input type="checkbox" name="useDate" value="true" id="useDateCheck"
                       <c:if test="${useDate}">checked</c:if>>
                출고일자
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
            <a href="${pageContext.request.contextPath}/outbound/list" class="btn btn-reset">초기화</a>
        </c:if>
        <button type="submit" class="btn btn-search">조회</button>
    </div>
</form>

<div class="list-hint">💡 행을 클릭하면 출고 디테일(어떤 입고에서 몇 개씩 나갔는지) 을 볼 수 있습니다.</div>

<div class="table-wrap">
<table class="clickable">
    <thead>
        <tr>
            <th style="width:60px;">#</th>
            <th style="width:180px;">출고번호</th>
            <th style="width:120px;">출고일자</th>
            <th style="width:180px;">COA</th>
            <th>출고장소</th>
            <th style="width:90px;">건수</th>
            <th style="width:90px;">총 수량</th>
            <th>비고</th>
            <th style="width:100px;">관리</th>
        </tr>
    </thead>
    <tbody>
        <c:forEach var="o" items="${outbounds}" varStatus="st">
            <tr onclick="openDetail('${o.outboundId}')">
                <td class="center">${st.index + 1}</td>
                <td class="center">${o.outboundId}</td>
                <td class="center">${o.outboundDate}</td>
                <td>${o.coaCode}</td>
                <td class="center">${o.outboundLocation}</td>
                <td class="num">${o.itemCount}</td>
                <td class="num" style="font-weight:600;color:var(--color-primary);">
                    <fmt:formatNumber value="${o.totalQty}" pattern="#,##0"/>
                </td>
                <td>${o.remark}</td>
                <td class="center" onclick="event.stopPropagation()">
                    <form action="${pageContext.request.contextPath}/outbound/delete" method="post" style="display:inline;"
                          onsubmit="return confirm('정말 삭제하시겠습니까? 관련 입고의 출고량이 원복됩니다.');">
                        <input type="hidden" name="outboundId" value="${o.outboundId}">
                        <button type="submit" class="btn btn-delete">삭제</button>
                    </form>
                </td>
            </tr>
        </c:forEach>
        <c:if test="${empty outbounds}">
            <tr><td colspan="9" class="empty">출고 내역이 없습니다.</td></tr>
        </c:if>
    </tbody>
</table>
</div>

<!-- ===== 상세 모달 ===== -->
<div id="detailModal" class="modal-overlay" onclick="if(event.target===this) closeDetail()">
    <div class="modal" style="width:700px;">
        <div class="modal-header">
            <h3>📤 출고 디테일</h3>
            <button type="button" class="modal-close" onclick="closeDetail()">&times;</button>
        </div>
        <div class="modal-body">
            <div id="m-info" class="readonly-info"></div>

            <div style="margin-top:14px;">
                <label style="margin-top:0;">🧾 출고된 입고 제품 리스트</label>
                <div id="m-items" style="background:var(--gray-50); border-radius:var(--radius-sm); padding:10px 12px; font-size:13px;">
                    <span style="color:var(--gray-500);">불러오는 중...</span>
                </div>
            </div>
        </div>
        <div class="modal-footer">
            <button type="button" class="btn btn-cancel" onclick="closeDetail()">닫기</button>
        </div>
    </div>
</div>

<script>
    const ctx = '${pageContext.request.contextPath}';

    function openDetail(id) {
        document.getElementById('m-info').textContent = '출고번호: ' + id;
        const box = document.getElementById('m-items');
        box.innerHTML = '<span style="color:var(--gray-500);">불러오는 중...</span>';
        document.getElementById('detailModal').classList.add('show');

        fetch(ctx + '/outbound/' + encodeURIComponent(id) + '/items')
            .then(r => r.json())
            .then(items => {
                if (!items.length) {
                    box.innerHTML = '<span style="color:var(--gray-500);">디테일 없음</span>';
                    return;
                }
                let html = '<table style="width:100%;font-size:13px;">';
                html += '<thead><tr style="text-align:left;color:var(--gray-600);">' +
                        '<th>제품번호</th><th>생산일자</th><th>포장단위</th><th>크기(μm)</th><th style="text-align:right;">출고수량</th>' +
                        '</tr></thead><tbody>';
                items.forEach(it => {
                    html += '<tr style="border-top:1px solid var(--gray-200);">' +
                            '<td>' + it.inboundId + '</td>' +
                            '<td>' + it.productionDate + '</td>' +
                            '<td>' + it.packageUnitKg + ' kg</td>' +
                            '<td>' + (it.sizeUm != null ? it.sizeUm : '-') + '</td>' +
                            '<td style="text-align:right;font-weight:600;">' + it.quantity + '</td>' +
                            '</tr>';
                });
                html += '</tbody></table>';
                box.innerHTML = html;
            })
            .catch(() => {
                box.innerHTML = '<span style="color:var(--color-danger);">조회 실패</span>';
            });
    }
    function closeDetail() {
        document.getElementById('detailModal').classList.remove('show');
    }
    document.addEventListener('keydown', e => {
        if (e.key === 'Escape') closeDetail();
    });

    // 일자 필터 토글
    const dateCheck = document.getElementById('useDateCheck');
    const startDate = document.getElementById('startDate');
    const endDate   = document.getElementById('endDate');
    if (dateCheck) {
        dateCheck.addEventListener('change', () => {
            startDate.disabled = !dateCheck.checked;
            endDate.disabled   = !dateCheck.checked;
        });
    }
</script>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
