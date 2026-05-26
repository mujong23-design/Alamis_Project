<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<%@ taglib prefix="fmt" uri="jakarta.tags.fmt" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>출고 등록</title>
    <style>
        .step-box {
            background: #fff; padding: 18px 22px;
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            margin-bottom: 14px;
        }
        .step-title {
            font-size: 15px; font-weight: 700; color: var(--gray-900);
            margin: 0 0 12px;
            display: flex; align-items: center; gap: 8px;
        }
        .step-num {
            display: inline-flex; align-items: center; justify-content: center;
            width: 22px; height: 22px;
            background: var(--color-primary); color: #fff;
            border-radius: 50%; font-size: 12px; font-weight: 700;
        }
        .selectable tbody tr.selected { background: var(--color-primary-light); }
        .selectable input[type="checkbox"] {
            width: 16px; height: 16px;
            accent-color: var(--color-primary);
            cursor: pointer;
        }
        .qty-input {
            width: 100px; padding: 6px 8px;
            border: 1px solid var(--gray-200);
            border-radius: var(--radius-sm);
            font-size: 13px;
            margin: 0;
        }
        .qty-input:disabled { background: var(--gray-100); color: var(--gray-400); }
        .qty-input.invalid {
            border-color: var(--color-danger);
            background: var(--color-danger-light);
        }
        .selected-info {
            background: var(--color-primary-light);
            color: var(--color-primary);
            padding: 12px 14px; border-radius: var(--radius-sm);
            font-size: 14px; margin-bottom: 14px; font-weight: 600;
            display: none;
        }
        .selected-info.show { display: flex; justify-content: space-between; }
        .selected-info .count { font-size: 18px; }

        .required-mark { color: var(--color-danger); margin-left: 2px; }
    </style>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>📤 출고 등록</h2>
</div>

<c:if test="${not empty error}">
    <div class="flash" style="background:var(--color-danger-light); border-left-color:var(--color-danger); color:var(--color-danger);">${error}</div>
</c:if>

<form action="${pageContext.request.contextPath}/outbound/save" method="post" id="outboundForm">

    <!-- 1단계: 입고 라인 선택 -->
    <div class="step-box">
        <h3 class="step-title"><span class="step-num">1</span>출고할 입고 제품 선택 (여러 개 가능)</h3>

        <div class="selected-info" id="selectedInfo">
            <span>✅ 선택된 라인 <span class="count" id="selCount">0</span>건</span>
            <span>총 출고 수량 <span class="count" id="selTotalQty">0</span></span>
        </div>

        <div class="table-wrap">
            <table class="selectable">
                <thead>
                    <tr>
                        <th style="width:50px;"><input type="checkbox" id="selectAll" title="전체 선택"></th>
                        <th style="width:180px;">제품번호</th>
                        <th style="width:120px;">생산일자</th>
                        <th style="width:100px;">포장단위(kg)</th>
                        <th style="width:90px;">포장개수</th>
                        <th style="width:90px;">잔여</th>
                        <th style="width:90px;">크기(μm)</th>
                        <th style="width:140px;">출고 수량 입력</th>
                    </tr>
                </thead>
                <tbody id="rowsBody">
                    <c:forEach var="i" items="${availableInbounds}">
                        <tr data-remaining="${i.remainingQty}">
                            <td class="center">
                                <input type="checkbox" name="inboundIds" value="${i.inboundId}">
                            </td>
                            <td class="center">${i.inboundId}</td>
                            <td class="center">${i.productionDate}</td>
                            <td class="num"><fmt:formatNumber value="${i.packageUnitKg}" pattern="0.###"/></td>
                            <td class="num">${i.packageCount}</td>
                            <td class="num" style="font-weight:600;color:var(--color-primary);">${i.remainingQty}</td>
                            <td class="num">
                                <c:if test="${not empty i.sizeUm}"><fmt:formatNumber value="${i.sizeUm}" pattern="0.000"/></c:if>
                            </td>
                            <td class="center">
                                <input type="number" class="qty-input" name="quantities"
                                       min="1" max="${i.remainingQty}" step="1"
                                       placeholder="0 / ${i.remainingQty}" disabled>
                            </td>
                        </tr>
                    </c:forEach>
                    <c:if test="${empty availableInbounds}">
                        <tr><td colspan="8" class="empty">출고 가능한 입고 제품이 없습니다.</td></tr>
                    </c:if>
                </tbody>
            </table>
        </div>
    </div>

    <!-- 2단계: 출고 정보 -->
    <div class="step-box">
        <h3 class="step-title"><span class="step-num">2</span>출고 정보 입력</h3>

        <div class="flex-row">
            <div>
                <label>출고일자</label>
                <input type="date" name="outboundDate" value="${today}" required>
            </div>
            <div>
                <label>COA (성적서 코드)<span class="required-mark">*</span></label>
                <input type="text" name="coaCode" required placeholder="필수 입력">
            </div>
        </div>

        <label>출고장소</label>
        <input type="text" name="outboundLocation" placeholder="예: ㈜OO화장품 본사 창고">

        <label>비고</label>
        <textarea name="remark" rows="3" placeholder="선택 사항"></textarea>

        <div class="btns">
            <button type="submit" class="btn btn-save" id="btnSubmit">출고 등록</button>
            <a href="${pageContext.request.contextPath}/outbound/list" class="btn btn-back">목록</a>
        </div>
    </div>
</form>

<script>
    const rowsBody = document.getElementById('rowsBody');
    const selectAll = document.getElementById('selectAll');
    const selectedInfo = document.getElementById('selectedInfo');
    const selCount = document.getElementById('selCount');
    const selTotalQty = document.getElementById('selTotalQty');
    const form = document.getElementById('outboundForm');
    const btnSubmit = document.getElementById('btnSubmit');

    // 전체 선택
    selectAll.addEventListener('change', () => {
        rowsBody.querySelectorAll('input[name="inboundIds"]').forEach(cb => {
            cb.checked = selectAll.checked;
            toggleRow(cb);
        });
        updateSummary();
    });

    // 행 클릭 / 체크박스 / 수량 입력
    rowsBody.addEventListener('change', e => {
        if (e.target.name === 'inboundIds') {
            toggleRow(e.target);
            updateSummary();
        }
    });
    rowsBody.addEventListener('input', e => {
        if (e.target.name === 'quantities') {
            validateQty(e.target);
            updateSummary();
        }
    });
    rowsBody.addEventListener('click', e => {
        // 행 자체를 클릭해도 체크박스 토글 (단, 입력 필드 클릭은 무시)
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'BUTTON') return;
        const row = e.target.closest('tr');
        if (!row || !row.dataset.remaining) return;
        const cb = row.querySelector('input[name="inboundIds"]');
        if (cb) {
            cb.checked = !cb.checked;
            toggleRow(cb);
            updateSummary();
        }
    });

    function toggleRow(cb) {
        const row = cb.closest('tr');
        const qty = row.querySelector('input[name="quantities"]');
        if (cb.checked) {
            row.classList.add('selected');
            qty.disabled = false;
            if (!qty.value) qty.value = '';   // 비워두고 사용자 입력 받음
            qty.focus();
        } else {
            row.classList.remove('selected');
            qty.disabled = true;
            qty.value = '';
            qty.classList.remove('invalid');
        }
        updateSelectAllState();
    }

    function validateQty(input) {
        const max = parseInt(input.max);
        const val = parseInt(input.value);
        if (val > max || val < 1) {
            input.classList.add('invalid');
        } else {
            input.classList.remove('invalid');
        }
    }

    function updateSelectAllState() {
        const all = rowsBody.querySelectorAll('input[name="inboundIds"]');
        const checked = rowsBody.querySelectorAll('input[name="inboundIds"]:checked');
        selectAll.checked = all.length > 0 && checked.length === all.length;
        selectAll.indeterminate = checked.length > 0 && checked.length < all.length;
    }

    function updateSummary() {
        const checked = rowsBody.querySelectorAll('input[name="inboundIds"]:checked');
        let total = 0;
        checked.forEach(cb => {
            const qty = parseInt(cb.closest('tr').querySelector('input[name="quantities"]').value || '0');
            if (!isNaN(qty)) total += qty;
        });
        selCount.textContent = checked.length;
        selTotalQty.textContent = total;
        if (checked.length > 0) selectedInfo.classList.add('show');
        else selectedInfo.classList.remove('show');
    }

    // 제출 검증
    form.addEventListener('submit', e => {
        const checked = rowsBody.querySelectorAll('input[name="inboundIds"]:checked');
        if (checked.length === 0) {
            e.preventDefault();
            alert('출고할 입고 제품을 1개 이상 선택해주세요.');
            return;
        }
        for (const cb of checked) {
            const qtyInput = cb.closest('tr').querySelector('input[name="quantities"]');
            const v = parseInt(qtyInput.value || '0');
            const max = parseInt(qtyInput.max);
            if (v < 1) {
                e.preventDefault();
                alert(cb.value + ': 출고 수량을 1 이상 입력해주세요.');
                qtyInput.focus();
                return;
            }
            if (v > max) {
                e.preventDefault();
                alert(cb.value + ': 출고 수량이 잔여재고(' + max + ')를 초과했습니다.');
                qtyInput.focus();
                return;
            }
        }
        const coa = form.querySelector('[name="coaCode"]').value.trim();
        if (!coa) {
            e.preventDefault();
            alert('COA(성적서 코드)는 필수입니다.');
            return;
        }
        if (!confirm(checked.length + '건 출고를 등록하시겠습니까?')) {
            e.preventDefault();
        }
    });
</script>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
