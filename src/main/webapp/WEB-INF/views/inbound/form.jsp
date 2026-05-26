<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>입고 등록 / 수정</title>
    <style>
        .preview-box {
            background: var(--color-primary-light);
            border-radius: var(--radius-sm);
            padding: 12px 16px;
            margin-top: 8px;
            font-size: 13px;
            color: var(--color-primary);
            font-weight: 500;
        }
    </style>
</head>
<body>
<jsp:include page="../common/header.jsp"/>

<div class="page-header">
    <h2>📥 ${isEdit ? '입고 수정' : '입고 등록'}</h2>
</div>

<c:if test="${not empty error}">
    <div class="flash" style="background:var(--color-danger-light); border-left-color:var(--color-danger); color:var(--color-danger);">${error}</div>
</c:if>

<div class="form-box">
    <form action="${pageContext.request.contextPath}/inbound/save" method="post">

        <c:if test="${isEdit}">
            <label>제품번호</label>
            <input type="text" name="inboundId" value="${inbound.inboundId}" readonly>
        </c:if>
        <c:if test="${!isEdit}">
            <label>제품번호 <span class="inline-hint">저장 시 자동 채번 (P-YYYYMMDD-NNN)</span></label>
            <input type="text" value="자동 생성됩니다" readonly style="color:var(--gray-500);">
        </c:if>

        <label>생산일자</label>
        <input type="date" name="productionDate" required
               value="${not empty inbound.productionDate ? inbound.productionDate : today}">

        <div class="flex-row">
            <div>
                <label>포장 단위 (kg) <span class="inline-hint">기본 10</span></label>
                <input type="number" name="packageUnitKg" step="0.001" min="0.001"
                       value="${not empty inbound.packageUnitKg ? inbound.packageUnitKg : 10}" required>
            </div>
            <div>
                <label>포장 개수 <span class="inline-hint">정수</span></label>
                <input type="number" name="packageCount" step="1" min="1"
                       value="${inbound.packageCount}" required id="packageCount">
            </div>
        </div>

        <label>크기 (μm) <span class="inline-hint">정수 2자리 + 소수 3자리 (최대 99.999)</span></label>
        <input type="number" name="sizeUm" step="0.001" min="0" max="99.999"
               value="${inbound.sizeUm}">

        <c:if test="${isEdit}">
            <div class="preview-box">
                ℹ️ 현재 출고량: <strong>${inbound.outboundQty}</strong> /
                포장개수: <strong>${inbound.packageCount}</strong>
                → 잔여 재고: <strong>${inbound.remainingQty}</strong>
            </div>
        </c:if>

        <label>비고</label>
        <textarea name="remark" rows="3">${inbound.remark}</textarea>

        <div class="btns">
            <button type="submit" class="btn btn-save">${isEdit ? '저장' : '등록'}</button>
            <a href="${pageContext.request.contextPath}/inbound/list" class="btn btn-back">목록</a>
        </div>
    </form>
</div>

<jsp:include page="../common/footer.jsp"/>
</body>
</html>
