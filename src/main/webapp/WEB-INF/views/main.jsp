<%@ page contentType="text/html;charset=UTF-8" language="java" %>
<%@ taglib prefix="c" uri="jakarta.tags.core" %>
<%@ taglib prefix="fmt" uri="jakarta.tags.fmt" %>
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>대시보드 - 재고관리</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
    <style>
        /* 대시보드 — 한 화면 가득 채움 (스크롤 X) */
        html { height: 100%; overflow: hidden; }
        body.dashboard { height: 100vh; overflow: hidden; margin: 0; }

        body.dashboard .content {
            height: calc(100vh - 64px);
            display: flex;
            flex-direction: column;
            gap: 12px;
            padding: 18px 24px;
            box-sizing: border-box;
        }
        body.dashboard .welcome,
        body.dashboard .stat-cards { flex-shrink: 0; margin-bottom: 0; }

        body.dashboard .welcome h1 { margin: 0 0 4px; font-size: 22px; }
        body.dashboard .welcome p  { color: var(--gray-500); font-size: 13px; }

        body.dashboard .stat-cards {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 12px;
        }
        body.dashboard .stat-card {
            background: #fff;
            padding: 16px 20px;
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            transition: all .15s ease;
        }
        body.dashboard .stat-card:hover { box-shadow: var(--shadow-md); transform: translateY(-2px); }
        body.dashboard .stat-card .label { color: var(--gray-500); font-size: 13px; margin-bottom: 4px; font-weight: 500; }
        body.dashboard .stat-card .value { font-size: 24px; font-weight: 700; color: var(--gray-900); line-height: 1.1; }
        body.dashboard .stat-card .unit  { font-size: 13px; color: var(--gray-500); font-weight: 500; margin-left: 4px; }
        body.dashboard .stat-card.primary .value { color: var(--color-primary); }
        body.dashboard .stat-card.danger  .value { color: var(--color-danger);  }
        body.dashboard .stat-card.success .value { color: var(--color-success); }
        body.dashboard .stat-card.warning .value { color: var(--color-warning); }

        /* 차트 2개가 남은 공간 1:1 로 나눔 */
        body.dashboard .chart-card {
            flex: 1;
            display: flex;
            flex-direction: column;
            min-height: 180px;
            background: #fff;
            padding: 14px 20px;
            border-radius: var(--radius-lg);
            box-shadow: var(--shadow-sm);
            box-sizing: border-box;
        }
        body.dashboard .chart-card h3 {
            font-size: 14px; font-weight: 700; color: var(--gray-800);
            margin: 0 0 10px;
            flex-shrink: 0;
        }
        body.dashboard .chart-canvas-wrap {
            flex: 1;
            min-height: 0;
            position: relative;
        }
    </style>
</head>
<body class="dashboard">
<jsp:include page="common/header.jsp"/>

<div class="welcome">
    <h1>안녕하세요, ${sessionScope.loginUser.userName}님 👋</h1>
    <p>재고관리 시스템 대시보드입니다.</p>
</div>

<!-- ===== 통계 카드 4개 ===== -->
<div class="stat-cards">
    <div class="stat-card">
        <div class="label">전체 입고건수</div>
        <div class="value">
            <fmt:formatNumber value="${totalInbounds}" pattern="#,##0"/><span class="unit">건</span>
        </div>
    </div>
    <div class="stat-card primary">
        <div class="label">현재 재고</div>
        <div class="value">
            <fmt:formatNumber value="${currentStock}" pattern="#,##0"/><span class="unit">포장</span>
        </div>
    </div>
    <div class="stat-card success">
        <div class="label">금일 입고</div>
        <div class="value">
            <fmt:formatNumber value="${todayInbounds}" pattern="#,##0"/><span class="unit">건</span>
        </div>
    </div>
    <div class="stat-card danger">
        <div class="label">금일 출고</div>
        <div class="value">
            <fmt:formatNumber value="${todayOutbounds}" pattern="#,##0"/><span class="unit">건</span>
            <span style="font-size:12px; color:var(--gray-500); font-weight:500; margin-left:6px;">
                (<fmt:formatNumber value="${todayOutboundQty}" pattern="#,##0"/>포장)
            </span>
        </div>
    </div>
</div>

<!-- ===== 차트 1: 월별 입출고 추이 ===== -->
<div class="chart-card">
    <h3>📈 월별 입출고 추이 (최근 6개월 · 포장 수 기준)</h3>
    <div class="chart-canvas-wrap">
        <canvas id="monthlyChart"></canvas>
    </div>
</div>

<!-- ===== 차트 2: 제품별 현재 재고 TOP 10 ===== -->
<div class="chart-card">
    <h3>📦 제품별 현재 재고 TOP 10</h3>
    <div class="chart-canvas-wrap">
        <canvas id="stockChart"></canvas>
    </div>
</div>

<script>
    const ctx = '${pageContext.request.contextPath}';
    const C_PRIMARY = '#3182F6';
    const C_DANGER  = '#F04452';
    const C_SUCCESS = '#00C896';

    Chart.defaults.font.family = "'Pretendard Variable', Pretendard, sans-serif";
    Chart.defaults.color = '#4E5968';

    fetch(ctx + '/dashboard/charts')
        .then(r => r.json())
        .then(data => {
            renderMonthlyChart(data.monthlyData);
            renderStockChart(data.topStocks);
        })
        .catch(err => console.error('차트 데이터 로딩 실패:', err));

    function renderMonthlyChart(monthly) {
        new Chart(document.getElementById('monthlyChart'), {
            type: 'line',
            data: {
                labels: monthly.map(m => m.month),
                datasets: [
                    {
                        label: '입고',
                        data: monthly.map(m => m.inbound),
                        borderColor: C_PRIMARY,
                        backgroundColor: 'rgba(49,130,246,0.1)',
                        borderWidth: 2,
                        tension: 0.3,
                        fill: true,
                        pointBackgroundColor: C_PRIMARY,
                        pointRadius: 4, pointHoverRadius: 6
                    },
                    {
                        label: '출고',
                        data: monthly.map(m => m.outbound),
                        borderColor: C_DANGER,
                        backgroundColor: 'rgba(240,68,82,0.1)',
                        borderWidth: 2,
                        tension: 0.3,
                        fill: true,
                        pointBackgroundColor: C_DANGER,
                        pointRadius: 4, pointHoverRadius: 6
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { intersect: false, mode: 'index' },
                plugins: {
                    legend: { position: 'bottom', labels: { usePointStyle: true, padding: 14, boxHeight: 8 } }
                },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: '#F2F4F6' } },
                    x: { grid: { display: false } }
                }
            }
        });
    }

    function renderStockChart(stocks) {
        if (!stocks || !stocks.length) {
            const canvas = document.getElementById('stockChart');
            canvas.parentNode.innerHTML =
                '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#95a5a6;font-size:13px;">' +
                '현재 재고가 있는 제품이 없습니다.</div>';
            return;
        }
        new Chart(document.getElementById('stockChart'), {
            type: 'bar',
            data: {
                labels: stocks.map(s => s.inboundId),
                datasets: [{
                    label: '잔여재고',
                    data: stocks.map(s => Number(s.remaining)),
                    backgroundColor: C_SUCCESS,
                    borderRadius: 6,
                    maxBarThickness: 40
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { callbacks: { label: ctx => ctx.parsed.y + ' 포장' } }
                },
                scales: {
                    y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: '#F2F4F6' } },
                    x: { grid: { display: false }, ticks: { font: { size: 10 } } }
                }
            }
        });
    }
</script>

<jsp:include page="common/footer.jsp"/>
</body>
</html>
