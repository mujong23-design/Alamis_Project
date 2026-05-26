<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c"   uri="jakarta.tags.core" %>
<%@ taglib prefix="fmt" uri="jakarta.tags.fmt" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>포장 라벨 인쇄</title>
    <link rel="stylesheet" href="${pageContext.request.contextPath}/css/style.css">
    <style>
        @page { size: A4; margin: 10mm; }
        body { background: #f0f0f0; padding: 20px; }

        .toolbar {
            max-width: 190mm; margin: 0 auto 16px;
            display: flex; justify-content: flex-end; gap: 8px;
        }

        /* A4 한 페이지 = 2 × 5 = 10 라벨 (각 약 90mm × 50mm) */
        .label-sheet {
            max-width: 190mm; margin: 0 auto;
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 4mm;
            background: #fff;
            padding: 4mm;
        }

        .label {
            border: 1px dashed #bbb;
            padding: 6mm 8mm;
            min-height: 48mm;
            display: flex; flex-direction: column;
            justify-content: space-between;
            page-break-inside: avoid;
        }
        .label .brand {
            font-size: 11px; color: #666; letter-spacing: 1px;
            border-bottom: 1px solid #333; padding-bottom: 3mm;
            display: flex; justify-content: space-between;
        }
        .label .brand strong { font-size: 14px; color: #000; letter-spacing: 0; }
        .label .id-line {
            font-size: 18px; font-weight: 700; text-align: center;
            font-family: 'Consolas', monospace;
            padding: 4mm 0;
            border-bottom: 1px solid #ddd;
        }
        .label .meta {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 2mm 8mm;
            font-size: 12px;
            padding-top: 3mm;
        }
        .label .meta .k { color: #666; }
        .label .meta .v { color: #000; font-weight: 600; text-align: right; }
        .label .seq {
            margin-top: 2mm;
            font-size: 10px;
            color: #999;
            text-align: right;
        }

        @media print {
            body { background: #fff; padding: 0; }
            .toolbar { display: none; }
            .label-sheet { padding: 0; max-width: none; }
            .label { border: 1px dashed #ccc; }
        }
    </style>
</head>
<body>

<div class="toolbar">
    <button type="button" class="btn btn-primary" onclick="window.print()">🖨️ 인쇄</button>
    <button type="button" class="btn btn-back" onclick="window.close()">닫기</button>
</div>

<div class="label-sheet">
    <c:set var="totalCnt" value="0"/>
    <c:forEach var="i" items="${inbounds}">
        <c:set var="cnt" value="${i.packageCount}"/>
        <c:forEach begin="1" end="${cnt}" var="seq">
            <div class="label">
                <div class="brand">
                    <strong>아라미스 펄</strong>
                    <span>PEARL</span>
                </div>

                <div class="id-line">${i.inboundId}</div>

                <div class="meta">
                    <div class="k">생산일자</div>
                    <div class="v">${i.productionDate}</div>

                    <div class="k">포장량</div>
                    <div class="v"><fmt:formatNumber value="${i.packageUnitKg}" pattern="0.###"/> kg</div>

                    <div class="k">크기 (μm)</div>
                    <div class="v">
                        <c:choose>
                            <c:when test="${not empty i.sizeUm}">
                                <fmt:formatNumber value="${i.sizeUm}" pattern="0.000"/>
                            </c:when>
                            <c:otherwise>-</c:otherwise>
                        </c:choose>
                    </div>
                </div>

                <div class="seq">#${seq} / ${cnt}</div>
            </div>
            <c:set var="totalCnt" value="${totalCnt + 1}"/>
        </c:forEach>
    </c:forEach>

    <c:if test="${empty inbounds}">
        <div style="grid-column:1/-1; text-align:center; padding:40px; color:#999;">
            인쇄할 라벨이 없습니다.
        </div>
    </c:if>
</div>

<script>
    window.addEventListener('load', () => {
        setTimeout(() => window.print(), 300);
    });
</script>
</body>
</html>
