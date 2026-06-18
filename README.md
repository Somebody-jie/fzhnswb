# fzhnswb
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.5, user-scalable=yes" />
    <title>华南2026年竣工项目 · 数据看板</title>
    <!-- ECharts -->
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js">
    </script>
    <!-- SheetJS (XLSX) -->
    <script src="https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js">
    </script>
    <style>
        /* ===== 全局重置 & 字体 ===== */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            background: linear-gradient(145deg, #f0f6fe 0%, #f8faff 100%);
            color: #1a2639;
            padding: 16px;
            min-height: 100vh;
        }
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #eef4fa;
            border-radius: 8px;
        }
        ::-webkit-scrollbar-thumb {
            background: #b9cddf;
            border-radius: 8px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #9bb4cc;
        }

        .dashboard {
            max-width: 1440px;
            margin: 0 auto;
        }

        /* ===== 顶部栏 ===== */
        .header {
            display: flex;
            flex-wrap: wrap;
            align-items: center;
            justify-content: space-between;
            background: rgba(255, 255, 255, 0.75);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            border-radius: 24px;
            padding: 16px 20px;
            margin-bottom: 24px;
            box-shadow: 0 8px 32px rgba(26, 58, 92, 0.08);
            border: 1px solid rgba(255, 255, 255, 0.6);
            transition: box-shadow 0.3s;
            gap: 12px;
        }
        .header:hover {
            box-shadow: 0 12px 40px rgba(26, 58, 92, 0.12);
        }

        .header-left {
            display: flex;
            align-items: center;
            gap: 12px;
            flex: 1 1 auto;
            min-width: 0;
        }
        .header-left .logo-icon {
            width: 40px;
            height: 40px;
            background: linear-gradient(135deg, #3b7bb0, #5c9ed1);
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #fff;
            font-weight: 700;
            font-size: 18px;
            flex-shrink: 0;
            box-shadow: 0 4px 12px rgba(59, 123, 176, 0.3);
        }
        .header-title {
            min-width: 0;
        }
        .header-title h1 {
            font-size: clamp(16px, 4vw, 22px);
            font-weight: 700;
            letter-spacing: -0.3px;
            color: #0b2a44;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .header-title .sub {
            font-size: clamp(11px, 2.5vw, 13px);
            color: #6f8ba5;
            font-weight: 400;
            margin-top: 2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .header-right {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-wrap: wrap;
            flex: 0 0 auto;
        }
        .header-right .update-time {
            font-size: 12px;
            color: #6f8ba5;
            background: rgba(240, 246, 254, 0.7);
            padding: 4px 12px;
            border-radius: 40px;
            white-space: nowrap;
            backdrop-filter: blur(4px);
            border: 1px solid rgba(255, 255, 255, 0.6);
            display: none; /* 在手机上隐藏，节省空间 */
        }
        .btn-group {
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
        }
        .btn {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            padding: 6px 14px;
            border: none;
            border-radius: 40px;
            font-size: 13px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.25s ease;
            background: #ffffff;
            color: #1a2639;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.02);
            border: 1px solid rgba(200, 215, 230, 0.3);
            white-space: nowrap;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(26, 58, 92, 0.08);
            background: #f8fbff;
        }
        .btn-primary {
            background: linear-gradient(135deg, #3b7bb0, #5c9ed1);
            color: #fff;
            border: none;
            box-shadow: 0 4px 14px rgba(59, 123, 176, 0.25);
        }
        .btn-primary:hover {
            background: linear-gradient(135deg, #2f6a9a, #4a8dbd);
            box-shadow: 0 8px 24px rgba(59, 123, 176, 0.35);
            transform: translateY(-2px);
        }
        .btn-success {
            background: #e8f5e9;
            color: #2e7d32;
            border: 1px solid rgba(46, 125, 50, 0.15);
        }
        .btn-success:hover {
            background: #c8e6c9;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(46, 125, 50, 0.15);
        }
        .btn-share {
            background: #e3f0fa;
            color: #1a4b6b;
            border: 1px solid rgba(59, 123, 176, 0.15);
        }
        .btn-share:hover {
            background: #c8e2f5;
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(59, 123, 176, 0.15);
        }
        .btn svg {
            width: 16px;
            height: 16px;
            fill: currentColor;
        }
        .upload-btn-wrapper {
            position: relative;
            overflow: hidden;
            display: inline-block;
        }
        .upload-btn-wrapper input[type="file"] {
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
            width: 100%;
            height: 100%;
            cursor: pointer;
            font-size: 0;
        }

        /* ===== KPI 卡片 ===== */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .kpi-card {
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(6px);
            -webkit-backdrop-filter: blur(6px);
            border-radius: 20px;
            padding: 18px 18px 16px;
            box-shadow: 0 4px 24px rgba(26, 58, 92, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.7);
            transition: all 0.3s ease;
        }
        .kpi-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 12px 40px rgba(26, 58, 92, 0.10);
            background: rgba(255, 255, 255, 0.95);
        }
        .kpi-card .label {
            font-size: 12px;
            font-weight: 500;
            color: #6f8ba5;
            letter-spacing: 0.3px;
            text-transform: uppercase;
            margin-bottom: 4px;
        }
        .kpi-card .value {
            font-size: clamp(24px, 6vw, 30px);
            font-weight: 700;
            color: #0b2a44;
            letter-spacing: -0.5px;
        }
        .kpi-card .value .unit {
            font-size: 13px;
            font-weight: 500;
            color: #6f8ba5;
            margin-left: 4px;
        }
        .kpi-card .trend {
            font-size: 12px;
            color: #6f8ba5;
            margin-top: 4px;
            display: flex;
            align-items: center;
            gap: 4px;
            flex-wrap: wrap;
        }
        .kpi-card.accent-1 .value {
            color: #2b5a8a;
        }
        .kpi-card.accent-2 .value {
            color: #3b7bb0;
        }
        .kpi-card.accent-3 .value {
            color: #2e7d32;
        }
        .kpi-card.accent-4 .value {
            color: #c62828;
        }

        /* ===== 图表区域 ===== */
        .charts-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 24px;
        }
        .chart-card {
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(6px);
            -webkit-backdrop-filter: blur(6px);
            border-radius: 20px;
            padding: 16px 16px 4px;
            box-shadow: 0 4px 24px rgba(26, 58, 92, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.7);
            transition: box-shadow 0.3s;
        }
        .chart-card:hover {
            box-shadow: 0 8px 32px rgba(26, 58, 92, 0.10);
        }
        .chart-card .chart-title {
            font-size: clamp(14px, 3vw, 16px);
            font-weight: 600;
            color: #0b2a44;
            margin-bottom: 2px;
            padding-left: 4px;
        }
        .chart-card .chart-sub {
            font-size: 11px;
            color: #8aa5bf;
            padding-left: 4px;
            margin-bottom: 6px;
        }
        .chart-container {
            width: 100%;
            height: 260px;
        }
        .chart-container.tall {
            height: 300px;
        }

        /* ===== 表格区域 ===== */
        .table-card {
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(6px);
            -webkit-backdrop-filter: blur(6px);
            border-radius: 20px;
            padding: 16px 16px 4px;
            box-shadow: 0 4px 24px rgba(26, 58, 92, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.7);
            margin-bottom: 24px;
            overflow: hidden;
            transition: box-shadow 0.3s;
        }
        .table-card:hover {
            box-shadow: 0 8px 32px rgba(26, 58, 92, 0.10);
        }
        .table-card .table-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 10px;
            padding: 0 4px;
            gap: 8px;
        }
        .table-card .table-header .title {
            font-size: clamp(14px, 3vw, 16px);
            font-weight: 600;
            color: #0b2a44;
        }
        .table-card .table-header .count {
            font-size: 12px;
            color: #6f8ba5;
        }
        .table-wrapper {
            overflow-x: auto;
            border-radius: 16px;
            border: 1px solid rgba(200, 215, 230, 0.3);
            -webkit-overflow-scrolling: touch;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            min-width: 800px;
        }
        table thead {
            background: #f6fafd;
        }
        table th {
            padding: 10px 12px;
            text-align: left;
            font-weight: 600;
            color: #2b4a66;
            border-bottom: 2px solid #dce6ef;
            white-space: nowrap;
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        table td {
            padding: 8px 12px;
            border-bottom: 1px solid #eef4fa;
            color: #1a2639;
            white-space: nowrap;
            min-width: 50px;
            font-size: 12px;
        }
        table td[contenteditable="true"] {
            background-color: #f9fcff;
            cursor: text;
            transition: background 0.2s;
        }
        table td[contenteditable="true"]:hover {
            background-color: #edf5fe;
        }
        table td[contenteditable="true"]:focus {
            outline: 2px solid #5c9ed1;
            outline-offset: -1px;
            background-color: #ffffff;
        }
        table tbody tr {
            transition: background 0.2s;
        }
        table tbody tr:hover {
            background: #f0f7fe;
        }
        table tbody tr.row-internal {
            background-color: #eef6fe;
        }
        table tbody tr.row-internal:hover {
            background-color: #e2eefa;
        }

        table .sector-badge {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 40px;
            font-size: 10px;
            font-weight: 600;
            background: #e9eef2;
            color: #2b3f54;
            white-space: nowrap;
        }
        .sector-badge.total {
            background: #dbe7f7;
            color: #1a3a5c;
        }
        .sector-badge.infrastructure {
            background: #dff0e5;
            color: #1e6b34;
        }
        .sector-badge.mechanical {
            background: #fdf0e0;
            color: #a85d1a;
        }

        .rate-positive {
            color: #2e7d32;
            font-weight: 600;
        }
        .rate-negative {
            color: #c62828;
            font-weight: 600;
        }
        .rate-neutral {
            color: #6f8ba5;
        }
        .overdue-high {
            color: #c62828;
            font-weight: 600;
        }
        .overdue-mid {
            color: #b25d1a;
            font-weight: 500;
        }
        .overdue-low {
            color: #6f8ba5;
        }

        .internal-tag {
            display: inline-block;
            background: #5c9ed1;
            color: #fff;
            font-size: 9px;
            font-weight: 600;
            padding: 1px 8px;
            border-radius: 40px;
            margin-left: 6px;
            letter-spacing: 0.3px;
            vertical-align: middle;
        }

        .edit-hint {
            font-size: 11px;
            color: #8aa5bf;
            margin-left: 8px;
            font-weight: 400;
            display: inline-block;
        }

        /* ===== Toast ===== */
        .toast {
            position: fixed;
            bottom: 20px;
            left: 50%;
            transform: translateX(-50%);
            background: #1a2f44;
            color: #fff;
            padding: 10px 24px;
            border-radius: 60px;
            font-size: 13px;
            font-weight: 500;
            box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
            opacity: 0;
            transition: opacity 0.4s ease;
            pointer-events: none;
            z-index: 999;
            backdrop-filter: blur(4px);
            border: 1px solid rgba(255, 255, 255, 0.1);
            max-width: 90%;
            text-align: center;
        }
        .toast.show {
            opacity: 1;
        }
        .toast.success {
            background: #1e6b34;
        }
        .toast.error {
            background: #a52828;
        }

        /* ===== 响应式 - 平板及手机 ===== */
        @media (max-width: 900px) {
            .charts-grid {
                grid-template-columns: 1fr;
                gap: 16px;
            }
            .chart-container {
                height: 240px;
            }
            .chart-container.tall {
                height: 270px;
            }
            .header-right .update-time {
                display: inline-block; /* 平板上显示 */
            }
        }

        @media (max-width: 700px) {
            body {
                padding: 12px;
            }
            .header {
                padding: 12px 14px;
                border-radius: 20px;
                gap: 10px;
            }
            .header-left .logo-icon {
                width: 32px;
                height: 32px;
                font-size: 16px;
                border-radius: 10px;
            }
            .header-title h1 {
                font-size: clamp(14px, 3.5vw, 18px);
            }
            .header-title .sub {
                font-size: 11px;
            }
            .header-right .update-time {
                display: none; /* 手机上隐藏，节省空间 */
            }
            .btn {
                font-size: 12px;
                padding: 5px 10px;
                gap: 3px;
            }
            .btn svg {
                width: 14px;
                height: 14px;
            }
            .kpi-grid {
                grid-template-columns: repeat(2, 1fr);
                gap: 12px;
            }
            .kpi-card {
                padding: 14px 14px 12px;
                border-radius: 16px;
            }
            .kpi-card .value {
                font-size: clamp(20px, 5vw, 26px);
            }
            .kpi-card .label {
                font-size: 11px;
            }
            .kpi-card .trend {
                font-size: 11px;
            }
            .chart-card {
                padding: 12px 12px 0px;
                border-radius: 16px;
            }
            .chart-container {
                height: 200px;
            }
            .chart-container.tall {
                height: 230px;
            }
            .table-card {
                padding: 12px 12px 0px;
                border-radius: 16px;
            }
            table {
                font-size: 11px;
                min-width: 700px;
            }
            table th {
                padding: 8px 8px;
                font-size: 10px;
            }
            table td {
                padding: 6px 8px;
                font-size: 11px;
                min-width: 40px;
            }
            .internal-tag {
                font-size: 8px;
                padding: 0 6px;
                margin-left: 4px;
            }
            .edit-hint {
                font-size: 10px;
                margin-left: 4px;
            }
            .toast {
                font-size: 12px;
                padding: 8px 18px;
                bottom: 16px;
            }
        }

        @media (max-width: 450px) {
            .kpi-grid {
                grid-template-columns: 1fr 1fr;
                gap: 10px;
            }
            .kpi-card .value {
                font-size: 20px;
            }
            .btn-group {
                gap: 4px;
            }
            .btn {
                font-size: 11px;
                padding: 4px 8px;
            }
            .btn svg {
                width: 12px;
                height: 12px;
            }
            .header-title .sub {
                font-size: 10px;
            }
            table {
                font-size: 10px;
                min-width: 600px;
            }
            table th,
            table td {
                padding: 4px 6px;
            }
            .chart-container {
                height: 180px;
            }
            .chart-container.tall {
                height: 200px;
            }
        }
    </style>
</head>
<body>

    <div class="dashboard" id="app">
        <!-- ===== 顶部栏 ===== -->
        <header class="header">
            <div class="header-left">
                <div class="logo-icon">📊</div>
                <div class="header-title">
                    <h1>华南 2026 竣工项目</h1>
                    <div class="sub">共 <span id="headerCount">12</span> 个项目</div>
                </div>
            </div>
            <div class="header-right">
                <span class="update-time" id="updateTime">📅 2026.3.16</span>
                <div class="btn-group">
                    <button class="btn btn-success" id="exportBtn" title="导出CSV">
                        <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zM6 20V4h7v5h5v11H6z"/><path d="M8 12l4-4 4 4-1.5 1.5L13 12v4h-2v-4l-1.5 1.5L8 12z"/></svg>
                        <span class="btn-label">CSV</span>
                    </button>
                    <button class="btn btn-share" id="shareBtn" title="导出完整HTML">
                        <svg viewBox="0 0 24 24"><path d="M18 16.08c-.76 0-1.44.3-1.96.77L8.91 12.7c.05-.23.09-.46.09-.7s-.04-.47-.09-.7l7.05-4.11c.54.5 1.25.81 2.04.81 1.66 0 3-1.34 3-3s-1.34-3-3-3-3 1.34-3 3c0 .24.04.47.09.7L8.04 9.81C7.5 9.31 6.79 9 6 9c-1.66 0-3 1.34-3 3s1.34 3 3 3c.79 0 1.5-.31 2.04-.81l7.12 4.16c-.05.21-.08.43-.08.65 0 1.61 1.31 2.92 2.92 2.92 1.61 0 2.92-1.31 2.92-2.92s-1.31-2.92-2.92-2.92z"/></svg>
                        <span class="btn-label">分享</span>
                    </button>
                    <div class="upload-btn-wrapper">
                        <button class="btn btn-primary">
                            <svg viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zM6 20V4h7v5h5v11H6z"/><path d="M8 12l4-4 4 4-1.5 1.5L13 12v4h-2v-4l-1.5 1.5L8 12z"/></svg>
                            <span class="btn-label">上传</span>
                        </button>
                        <input type="file" id="fileInput" accept=".xlsx,.xls" />
                    </div>
                </div>
            </div>
        </header>

        <!-- ===== KPI ===== -->
        <section class="kpi-grid" id="kpiGrid">
            <div class="kpi-card accent-1">
                <div class="label">📦 项目总数</div>
                <div class="value" id="kpiTotal">0</div>
                <div class="trend" id="kpiTotalSub">竣工 0 · 内控 0</div>
            </div>
            <div class="kpi-card accent-2">
                <div class="label">💰 合计结算总额</div>
                <div class="value" id="kpiAmount">0 <span class="unit">万元</span></div>
                <div class="trend">自施 <span id="kpiSelfAmount">0</span> 万元</div>
            </div>
            <div class="kpi-card accent-3">
                <div class="label">📈 平均收益率</div>
                <div class="value" id="kpiRate">0%</div>
                <div class="trend">自施核定 <span id="kpiSelfRate">0%</span></div>
            </div>
            <div class="kpi-card accent-4">
                <div class="label">⏰ 超期项目</div>
                <div class="value" id="kpiOverdue">0</div>
                <div class="trend">最长 <span id="kpiMaxOverdue">0</span> 个月</div>
            </div>
        </section>

        <!-- ===== 图表 ===== -->
        <section class="charts-grid">
            <div class="chart-card">
                <div class="chart-title">🏗️ 按版块分布</div>
                <div class="chart-sub">全部项目 · 含内控</div>
                <div class="chart-container" id="chartSector"></div>
            </div>
            <div class="chart-card">
                <div class="chart-title">📊 项目收益率排行</div>
                <div class="chart-sub">全部项目 · 含内控</div>
                <div class="chart-container" id="chartRate"></div>
            </div>
            <div class="chart-card" style="grid-column: 1 / -1;">
                <div class="chart-title">⏳ 超期月数排行</div>
                <div class="chart-sub">全部项目 · 含内控</div>
                <div class="chart-container tall" id="chartOverdue"></div>
            </div>
        </section>

        <!-- ===== 表格 ===== -->
        <section class="table-card">
            <div class="table-header">
                <span class="title">📋 项目明细 <span class="edit-hint">（点击编辑）</span></span>
                <span class="count" id="tableCount">共 0 个项目</span>
            </div>
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>序号</th>
                            <th>项目名称</th>
                            <th>项目经理</th>
                            <th>商务经理</th>
                            <th>版块</th>
                            <th>自施总额</th>
                            <th>合计总额</th>
                            <th>自施收益率</th>
                            <th>自施收益</th>
                            <th>结算收益</th>
                            <th>结算收益率</th>
                            <th>完工日期</th>
                            <th>超期月数</th>
                        </tr>
                    </thead>
                    <tbody id="tableBody"></tbody>
                </table>
            </div>
        </section>

        <!-- ===== Toast ===== -->
        <div class="toast" id="toast"></div>
    </div>

    <script>
        // ============================================================
        //  默认示例数据（含内控）
        // ============================================================
        const DEFAULT_DATA = [{
            id: 1,
            name: '深圳新皇岗口岸综合业务楼',
            manager: '李颖辉',
            bizManager: '周坚',
            sector: '总包',
            selfTotal: 68439.891207,
            total: 85500.015942,
            selfRate: 0,
            selfProfit: 4536.83056035,
            totalProfit: 5813.92,
            rate: 0.0680,
            completionDate: '2025-12-22',
            overdue: 0.3,
            isInternal: false
        }, {
            id: 2,
            name: '深圳南山中国电子科技园',
            manager: '王博',
            bizManager: '桑海鹏',
            sector: '总包',
            selfTotal: 70000,
            total: 86000,
            selfRate: 0.055,
            selfProfit: 3850,
            totalProfit: -1711.63,
            rate: -0.0199,
            completionDate: '2025-03-25',
            overdue: 9.37,
            isInternal: false
        }, {
            id: 3,
            name: '深圳市坪地北地块场平工程',
            manager: '王岩',
            bizManager: '周文',
            sector: '基础设施',
            selfTotal: 13500,
            total: 13500,
            selfRate: 0.08,
            selfProfit: 1080,
            totalProfit: 42.38,
            rate: 0.00314,
            completionDate: '2025-04-01',
            overdue: 9.13,
            isInternal: false
        }, {
            id: 4,
            name: '珠海华发广场',
            manager: '程玉琦',
            bizManager: '无人',
            sector: '总包',
            selfTotal: 22800,
            total: 22800,
            selfRate: 0.0561,
            selfProfit: 1279.08,
            totalProfit: -5720,
            rate: -0.2509,
            completionDate: '2022-09-30',
            overdue: 39.6,
            isInternal: false
        }, {
            id: 5,
            name: '江门华发四季',
            manager: '姜超',
            bizManager: '刘子轩',
            sector: '总包',
            selfTotal: 38300,
            total: 38300,
            selfRate: 0.0563,
            selfProfit: 2156.29,
            totalProfit: -5769.29,
            rate: -0.1506,
            completionDate: '2021-10-29',
            overdue: 50.8,
            isInternal: false
        }, {
            id: 6,
            name: '东莞OPPO滨海湾数据中心机电',
            manager: '刘飞',
            bizManager: '张旭柱',
            sector: '机电',
            selfTotal: 14067.53,
            total: 15582.16,
            selfRate: 0.065,
            selfProfit: 892.5345,
            totalProfit: 50,
            rate: 0.00321,
            completionDate: '2022-11-30',
            overdue: 37.57,
            isInternal: false
        }, {
            id: 7,
            name: '深圳礼鼎高端集成电路载板及先进封装基地',
            manager: '辛昕',
            bizManager: '龙郑浩',
            sector: '总包',
            selfTotal: 132000,
            total: 132000,
            selfRate: 0.0566,
            selfProfit: 7492.5,
            totalProfit: 5893.58,
            rate: 0.04465,
            completionDate: '2023-10-30',
            overdue: 26.43,
            isInternal: false
        }, {
            id: 8,
            name: '深圳前海深港青年梦工场北区(二期)',
            manager: '辛昕',
            bizManager: '龙郑浩',
            sector: '总包',
            selfTotal: 14500,
            total: 28400,
            selfRate: 0.0548,
            selfProfit: 794.6,
            totalProfit: -833.57,
            rate: -0.02935,
            completionDate: '2023-09-25',
            overdue: 27.6,
            isInternal: false
        }, {
            id: 9,
            name: '广州保利三元里',
            manager: '杨丽鹏',
            bizManager: '张培培',
            sector: '总包',
            selfTotal: 50786.59,
            total: 51500,
            selfRate: 0.0601,
            selfProfit: 3052.274059,
            totalProfit: -6300,
            rate: -0.1223,
            completionDate: '2023-04-30',
            overdue: 32.53,
            isInternal: false
        }, {
            id: 10,
            name: '深圳嘉里商务中心(三期)',
            manager: '张建锋',
            bizManager: '沈志斌',
            sector: '总包',
            selfTotal: 41225.065484,
            total: 88693.888769,
            selfRate: 0.0506,
            selfProfit: 2085.988313,
            totalProfit: 1225.4,
            rate: 0.01382,
            completionDate: '2024-11-30',
            overdue: 13.2,
            isInternal: false
        }, {
            id: 11,
            name: '深圳光明区金融街',
            manager: '张秋学',
            bizManager: '石福生',
            sector: '总包',
            selfTotal: 36300,
            total: 36300,
            selfRate: 0.057,
            selfProfit: 2069.1,
            totalProfit: -5030.15,
            rate: -0.1386,
            completionDate: '2023-04-30',
            overdue: 43.7,
            isInternal: false
        }, {
            id: 12,
            name: '深圳中山大学附属七院二期',
            manager: '林志松',
            bizManager: '许亚男',
            sector: '总包',
            selfTotal: 105478,
            total: 105478,
            selfRate: 0.0534,
            selfProfit: 5642.2645,
            totalProfit: 5734.64,
            rate: 0.05437,
            completionDate: '2025-06-30',
            overdue: 6.13,
            isInternal: false
        }, {
            id: 13,
            name: '龙城街道龙飞学校新建工程(二期)',
            manager: '王啸飞',
            bizManager: '唐清鉴',
            sector: '总包',
            selfTotal: 15446.97,
            total: 30155.9141466,
            selfRate: 0.0526,
            selfProfit: 812.51,
            totalProfit: 1180.05,
            rate: 0.03913,
            completionDate: '2026-12-30',
            overdue: 0,
            isInternal: true
        }, {
            id: 14,
            name: '深圳机场教育基地项目',
            manager: '张建锋',
            bizManager: '王洪艳',
            sector: '总包',
            selfTotal: 26224.4,
            total: 44383.68,
            selfRate: 0.0528,
            selfProfit: 1384.65,
            totalProfit: 1384.65,
            rate: 0.0312,
            completionDate: '2026-04-30',
            overdue: 0,
            isInternal: true
        }, {
            id: 15,
            name: '怡翠',
            manager: '',
            bizManager: '',
            sector: '总包',
            selfTotal: 18640,
            total: 31140,
            selfRate: null,
            selfProfit: null,
            totalProfit: null,
            rate: null,
            completionDate: '',
            overdue: null,
            isInternal: true
        }];

        // ============================================================
        //  工具函数
        // ============================================================
        let currentData = [];
        let chartInstances = {};
        let editTimer = null;

        function formatNum(v, digits = 2) {
            if (v === undefined || v === null || isNaN(v)) return '—';
            return Number(v).toFixed(digits);
        }

        function formatRate(v) {
            if (v === undefined || v === null || isNaN(v)) return '—';
            return (v * 100).toFixed(2) + '%';
        }

        function formatDate(d) {
            if (!d) return '—';
            if (d instanceof Date) {
                return d.toISOString().slice(0, 10);
            }
            const s = String(d);
            if (s.includes('-')) {
                const parts = s.split(' ');
                return parts[0];
            }
            return s;
        }

        function safeNum(v) {
            if (v === undefined || v === null || v === '') return null;
            if (typeof v === 'string' && v.includes('%')) {
                v = v.replace(/%/g, '').trim();
            }
            const n = parseFloat(v);
            return isNaN(n) ? null : n;
        }

        function getSectorClass(sector) {
            const map = { '总包': 'total', '基础设施': 'infrastructure', '机电': 'mechanical' };
            return map[sector] || '';
        }

        let toastTimer = null;

        function showToast(msg, type = 'info') {
            const el = document.getElementById('toast');
            el.textContent = msg;
            el.className = 'toast ' + type;
            clearTimeout(toastTimer);
            void el.offsetWidth;
            el.classList.add('show');
            toastTimer = setTimeout(() => {
                el.classList.remove('show');
            }, 3500);
        }

        // ============================================================
        //  渲染看板 (主入口)
        // ============================================================
        function renderDashboard(data) {
            const allProjects = data;
            const mainProjects = data.filter(d => !d.isInternal);
            const totalCount = allProjects.length;
            const totalMain = mainProjects.length;
            const internalCount = totalCount - totalMain;

            const sumTotal = mainProjects.reduce((s, d) => s + d.total, 0);
            const sumSelf = mainProjects.reduce((s, d) => s + d.selfTotal, 0);
            const validRate = mainProjects.filter(d => d.rate !== null && !isNaN(d.rate));
            const avgRate = validRate.length ? validRate.reduce((s, d) => s + d.rate, 0) / validRate.length : 0;
            const validSelfRate = mainProjects.filter(d => d.selfRate !== null && !isNaN(d.selfRate));
            const avgSelfRate = validSelfRate.length ? validSelfRate.reduce((s, d) => s + d.selfRate, 0) / validSelfRate
                .length : 0;
            const overdueProjects = mainProjects.filter(d => d.overdue !== null && d.overdue > 0);
            const maxOverdue = overdueProjects.length ? Math.max(...overdueProjects.map(d => d.overdue)) : 0;

            document.getElementById('kpiTotal').textContent = totalCount;
            document.getElementById('kpiTotalSub').textContent = `竣工 ${totalMain} · 内控 ${internalCount}`;
            document.getElementById('kpiAmount').innerHTML = formatNum(sumTotal, 1) + ' <span class="unit">万元</span>';
            document.getElementById('kpiSelfAmount').textContent = formatNum(sumSelf, 1);
            document.getElementById('kpiRate').textContent = formatRate(avgRate);
            document.getElementById('kpiSelfRate').textContent = formatRate(avgSelfRate);
            document.getElementById('kpiOverdue').textContent = overdueProjects.length;
            document.getElementById('kpiMaxOverdue').textContent = formatNum(maxOverdue, 1);

            document.getElementById('headerCount').textContent = totalCount;
            document.getElementById('tableCount').textContent = '共 ' + allProjects.length + ' 个项目';

            renderTable(allProjects);
            renderSectorChart(allProjects);
            renderRateChart(allProjects);
            renderOverdueChart(allProjects);

            currentData = data;
        }

        // ============================================================
        //  表格渲染
        // ============================================================
        function renderTable(data) {
            const tbody = document.getElementById('tableBody');
            if (!data.length) {
                tbody.innerHTML =
                    `<tr><td colspan="13" style="text-align:center;padding:30px;color:#6f8ba5;">暂无数据</td></tr>`;
                return;
            }
            let html = '';
            data.forEach((d, index) => {
                const isInternal = d.isInternal;
                const rowClass = isInternal ? 'row-internal' : '';
                const internalTag = isInternal ? '<span class="internal-tag">内控</span>' : '';
                const sectorCls = getSectorClass(d.sector);
                const rateCls = (d.rate !== null && !isNaN(d.rate)) ? (d.rate >= 0 ? 'rate-positive' : 'rate-negative') :
                    'rate-neutral';
                const overdueCls = (d.overdue !== null && !isNaN(d.overdue)) ?
                    (d.overdue > 24 ? 'overdue-high' : d.overdue > 12 ? 'overdue-mid' : 'overdue-low') :
                    'rate-neutral';
                const editableAttr = (field) => `contenteditable="true" data-field="${field}" data-row="${index}"`;

                html += `<tr class="${rowClass}">`;
                html += `<td>${d.id}</td>`;
                html += `<td><strong>${d.name}</strong>${internalTag}</td>`;
                html += `<td ${editableAttr('manager')}>${d.manager || ''}</td>`;
                html += `<td ${editableAttr('bizManager')}>${d.bizManager || ''}</td>`;
                html +=
                `<td ${editableAttr('sector')}><span class="sector-badge ${sectorCls}">${d.sector || ''}</span></td>`;
                html +=
                    `<td ${editableAttr('selfTotal')}>${d.selfTotal !== null && !isNaN(d.selfTotal) ? d.selfTotal : ''}</td>`;
                html +=
                    `<td ${editableAttr('total')}>${d.total !== null && !isNaN(d.total) ? d.total : ''}</td>`;
                html +=
                    `<td ${editableAttr('selfRate')}>${d.selfRate !== null && !isNaN(d.selfRate) ? d.selfRate : ''}</td>`;
                html +=
                    `<td ${editableAttr('selfProfit')}>${d.selfProfit !== null && !isNaN(d.selfProfit) ? d.selfProfit : ''}</td>`;
                html +=
                    `<td ${editableAttr('totalProfit')}>${d.totalProfit !== null && !isNaN(d.totalProfit) ? d.totalProfit : ''}</td>`;
                html +=
                    `<td ${editableAttr('rate')} class="${rateCls}">${d.rate !== null && !isNaN(d.rate) ? d.rate : ''}</td>`;
                html += `<td ${editableAttr('completionDate')}>${formatDate(d.completionDate)}</td>`;
                html +=
                    `<td ${editableAttr('overdue')} class="${overdueCls}">${d.overdue !== null && !isNaN(d.overdue) ? d.overdue : ''}</td>`;
                html += `</tr>`;
            });
            tbody.innerHTML = html;

            tbody.removeEventListener('input', onTableEdit);
            tbody.addEventListener('input', onTableEdit);
        }

        function onTableEdit(e) {
            const cell = e.target;
            if (!cell.hasAttribute('contenteditable')) return;
            const field = cell.dataset.field;
            const rowIdx = parseInt(cell.dataset.row);
            if (isNaN(rowIdx) || !field) return;
            let newValue = cell.innerText.trim();
            const numericFields = ['selfTotal', 'total', 'selfRate', 'selfProfit', 'totalProfit', 'rate', 'overdue'];
            let parsedValue;
            if (numericFields.includes(field)) {
                parsedValue = safeNum(newValue);
                if (parsedValue === null && newValue !== '') {
                    showToast('请输入有效数字', 'error');
                    renderTable(currentData);
                    return;
                }
            } else {
                parsedValue = newValue;
            }
            if (rowIdx < currentData.length) {
                const oldVal = currentData[rowIdx][field];
                if (parsedValue === oldVal) return;
                currentData[rowIdx][field] = parsedValue;
                clearTimeout(editTimer);
                editTimer = setTimeout(() => {
                    renderDashboard(currentData);
                    showToast('数据已更新，图表同步刷新', 'success');
                }, 300);
            }
        }

        // ============================================================
        //  图表渲染
        // ============================================================
        function renderSectorChart(projects) {
            const sectorMap = {};
            projects.forEach(d => {
                const sec = d.sector || '未分类';
                if (!sectorMap[sec]) sectorMap[sec] = { count: 0, total: 0 };
                sectorMap[sec].count += 1;
                sectorMap[sec].total += d.total || 0;
            });
            const sectorNames = Object.keys(sectorMap);
            const sectorColors = {
                '总包': '#3b7bb0',
                '基础设施': '#4caf84',
                '机电': '#f5a35b',
                '未分类': '#a0b8cc'
            };
            const pieData = sectorNames.map(name => ({ name, value: sectorMap[name].total, count: sectorMap[name].count }));
            const chart = initChart('chartSector');
            chart.setOption({
                tooltip: { trigger: 'item', formatter: function(params) { const p = params.data; return `<strong>${p.name}</strong><br/>结算总额: ${p.value.toFixed(1)} 万元<br/>项目数: ${p.count} 个<br/>占比: ${p.percent}%`; } },
                legend: { orient: 'vertical', right: '5%', top: 'center', textStyle: { fontSize: 12, color: '#2b4a66' },
                    formatter: function(name) { const item = sectorMap[name]; return name + '  (' + item.count + '个)'; } },
                series: [{
                    type: 'pie',
                    radius: ['42%', '72%'],
                    center: ['42%', '50%'],
                    avoidLabelOverlap: true,
                    itemStyle: { borderRadius: 8, borderColor: '#ffffff', borderWidth: 2 },
                    label: { show: true, formatter: '{b}\n{d}%', fontSize: 11, color: '#1a2639' },
                    labelLine: { length: 12, length2: 16 },
                    emphasis: { scale: true, scaleSize: 8 },
                    data: pieData.map(d => ({ ...d, itemStyle: { color: sectorColors[d.name] || '#a0b8cc' } }))
                }]
            });
            chartInstances['chartSector'] = chart;
        }

        function renderRateChart(projects) {
            const rateData = projects.filter(d => d.rate !== null && !isNaN(d.rate)).sort((a, b) => b.rate - a.rate);
            const chart = initChart('chartRate');
            if (!rateData.length) {
                chart.setOption({ title: { text: '暂无收益率数据', left: 'center', top: 'center', textStyle: { color: '#8aa5bf',
                            fontSize: 14 } } });
                return;
            }
            const names = rateData.map(d => d.name.length > 12 ? d.name.slice(0, 12) + '…' : d.name);
            const values = rateData.map(d => +(d.rate * 100).toFixed(2));
            const colors = values.map(v => v >= 0 ? '#3b7bb0' : '#e57373');
            chart.setOption({
                tooltip: { trigger: 'axis', formatter: function(params) { const p = params[0]; const idx = p.dataIndex; const d =
                            rateData[idx]; return `<strong>${d.name}</strong><br/>结算收益率: ${p.value.toFixed(2)}%<br/>结算收益: ${d.totalProfit !== null && !isNaN(d.totalProfit) ? d.totalProfit.toFixed(2) : '—'} 万元`; } },
                grid: { left: '6%', right: '6%', top: '10%', bottom: '18%' },
                xAxis: { type: 'category', data: names, axisLabel: { rotate: 35, fontSize: 10, color: '#4a6a8a',
                        interval: 0 }, axisLine: { lineStyle: { color: '#dce6ef' } } },
                yAxis: { type: 'value', name: '收益率 (%)', nameTextStyle: { fontSize: 11, color: '#6f8ba5' },
                    axisLabel: { fontSize: 11, color: '#4a6a8a', formatter: '{value}%' },
                    splitLine: { lineStyle: { color: '#eef4fa', type: 'dashed' } } },
                series: [{
                    type: 'bar',
                    data: values.map((v, i) => ({ value: v, itemStyle: { color: colors[i], borderRadius: [6, 6, 0,
                                0] } })),
                    barWidth: '46%',
                    label: { show: true, position: 'top', formatter: function(params) { return params.value.toFixed(
                                1) + '%'; }, fontSize: 10, color: '#1a2639', fontWeight: 500 }
                }]
            });
            chartInstances['chartRate'] = chart;
        }

        function renderOverdueChart(projects) {
            const overdueData = projects.filter(d => d.overdue !== null && !isNaN(d.overdue) && d.overdue > 0).sort((a,
            b) => b.overdue - a.overdue);
            const chart = initChart('chartOverdue');
            if (!overdueData.length) {
                chart.setOption({ title: { text: '暂无超期项目', left: 'center', top: 'center', textStyle: { color: '#8aa5bf',
                            fontSize: 14 } } });
                return;
            }
            const names = overdueData.map(d => d.name.length > 14 ? d.name.slice(0, 14) + '…' : d.name);
            const values = overdueData.map(d => +(d.overdue).toFixed(1));
            const colors = values.map(v => v > 36 ? '#e57373' : v > 18 ? '#f5a35b' : '#3b7bb0');
            chart.setOption({
                tooltip: { trigger: 'axis', formatter: function(params) { const p = params[0]; const idx = p.dataIndex; const d =
                            overdueData[idx]; return `<strong>${d.name}</strong><br/>超期月数: ${p.value.toFixed(1)} 个月<br/>考核完工: ${formatDate(d.completionDate)}`; } },
                grid: { left: '6%', right: '6%', top: '10%', bottom: '18%' },
                xAxis: { type: 'category', data: names, axisLabel: { rotate: 35, fontSize: 10, color: '#4a6a8a',
                        interval: 0 }, axisLine: { lineStyle: { color: '#dce6ef' } } },
                yAxis: { type: 'value', name: '超期月数', nameTextStyle: { fontSize: 11, color: '#6f8ba5' },
                    axisLabel: { fontSize: 11, color: '#4a6a8a', formatter: '{value}月' },
                    splitLine: { lineStyle: { color: '#eef4fa', type: 'dashed' } } },
                series: [{
                    type: 'bar',
                    data: values.map((v, i) => ({ value: v, itemStyle: { color: colors[i], borderRadius: [6, 6, 0,
                                0] } })),
                    barWidth: '46%',
                    label: { show: true, position: 'top', formatter: function(params) { return params.value.toFixed(
                                1) + '月'; }, fontSize: 10, color: '#1a2639', fontWeight: 500 }
                }]
            });
            chartInstances['chartOverdue'] = chart;
        }

        function initChart(domId) {
            const dom = document.getElementById(domId);
            if (!dom) return null;
            if (chartInstances[domId]) chartInstances[domId].dispose();
            const instance = echarts.init(dom);
            window.addEventListener('resize', () => { instance.resize(); });
            return instance;
        }

        // ============================================================
        //  导出 CSV
        // ============================================================
        document.getElementById('exportBtn').addEventListener('click', function() {
            const data = currentData;
            if (!data || !data.length) { showToast('没有数据可导出', 'error'); return; }
            const headers = ['序号', '项目名称', '项目经理', '商务经理', '版块', '自施总额(万元)', '合计总额(万元)',
                '自施核定收益率', '自施收益(万元)', '结算收益(万元)', '结算收益率', '完工日期', '超期月数', '是否内控'
            ];
            const rows = data.map(d => [
                d.id, d.name, d.manager || '', d.bizManager || '', d.sector || '',
                d.selfTotal ?? '', d.total ?? '',
                d.selfRate !== null && !isNaN(d.selfRate) ? (d.selfRate * 100).toFixed(2) + '%' : '',
                d.selfProfit ?? '', d.totalProfit ?? '',
                d.rate !== null && !isNaN(d.rate) ? (d.rate * 100).toFixed(2) + '%' : '',
                formatDate(d.completionDate), d.overdue ?? '', d.isInternal ? '是' : '否'
            ]);
            let csvContent = headers.join(',') + '\n';
            rows.forEach(row => {
                const escaped = row.map(cell => {
                    if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"') || cell.includes(
                            '\n'))) return `"${cell.replace(/"/g, '""')}"`;
                    return cell;
                });
                csvContent += escaped.join(',') + '\n';
            });
            const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `华南2026竣工项目_${new Date().toISOString().slice(0,10)}.csv`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
            showToast('✅ 数据导出成功', 'success');
        });

        // ============================================================
        //  分享功能：导出完整 HTML 文件（含当前数据）
        // ============================================================
        document.getElementById('shareBtn').addEventListener('click', function() {
            const data = currentData;
            if (!data || !data.length) {
                showToast('没有数据可分享', 'error');
                return;
            }
            // 获取当前页面的完整 HTML，并替换 DEFAULT_DATA
            let fullHtml = document.documentElement.outerHTML;
            const dataJson = JSON.stringify(data, (key, value) => {
                if (typeof value === 'number' && !isFinite(value)) return null;
                return value;
            }, 2);
            const regex = /const\s+DEFAULT_DATA\s*=\s*\[[\s\S]*?\];/;
            const newDataStr = `const DEFAULT_DATA = ${dataJson};`;
            fullHtml = fullHtml.replace(regex, newDataStr);
            // 移除分享按钮的点击事件（避免干扰），但保留其他功能
            const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `华南2026竣工看板_${new Date().toISOString().slice(0,10)}.html`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
            showToast('✅ 完整看板已导出为 HTML 文件，可分享', 'success');
        });

        // ============================================================
        //  解析 Excel (智能查找表头行)
        // ============================================================
        function parseExcelFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                        let targetSheet = null;
                        let targetSheetName = '';
                        for (let name of workbook.SheetNames) {
                            const sheet = workbook.Sheets[name];
                            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 });
                            if (rows.length === 0) continue;
                            let found = false;
                            for (let r = 0; r < Math.min(rows.length, 30); r++) {
                                const row = rows[r];
                                if (row.some(cell => /项目名称|项目名|工程名称/.test(String(cell).trim()))) {
                                    found = true;
                                    break;
                                }
                            }
                            if (found) {
                                targetSheet = sheet;
                                targetSheetName = name;
                                break;
                            }
                        }
                        if (!targetSheet) {
                            targetSheetName = workbook.SheetNames[0];
                            targetSheet = workbook.Sheets[targetSheetName];
                        }
                        if (!targetSheet) {
                            reject(new Error('未找到有效数据表'));
                            return;
                        }

                        const rawRows = XLSX.utils.sheet_to_json(targetSheet, { defval: '', header: 1 });
                        if (rawRows.length < 2) {
                            reject(new Error('数据表为空或格式不正确'));
                            return;
                        }

                        let headerRowIndex = -1;
                        for (let r = 0; r < Math.min(rawRows.length, 30); r++) {
                            const row = rawRows[r];
                            if (row.some(cell => /项目名称|项目名|工程名称/.test(String(cell).trim()))) {
                                headerRowIndex = r;
                                break;
                            }
                        }
                        if (headerRowIndex === -1) {
                            headerRowIndex = 0;
                        }

                        const headerRow = rawRows[headerRowIndex].map(cell => String(cell).trim());
                        const colIndex = (keywords) => {
                            for (let kw of keywords) {
                                for (let i = 0; i < headerRow.length; i++) {
                                    if (headerRow[i].includes(kw)) return i;
                                }
                            }
                            return -1;
                        };

                        const idx = {
                            id: colIndex(['序号', '序号', '编号']),
                            name: colIndex(['项目名称', '项目名', '工程名称']),
                            manager: colIndex(['项目经理', '项目负责人']),
                            bizManager: colIndex(['商务经理', '商务负责人']),
                            sector: colIndex(['所属版块', '版块', '板块']),
                            selfTotal: colIndex(['自施预计结算总额合理值', '自施预计结算总额', '自施总额']),
                            total: colIndex(['合计预计结算总额合理值', '合计预计结算总额', '合计总额']),
                            selfRate: colIndex(['自施核定收益率', '自施收益率']),
                            selfProfit: colIndex(['自施工核定收益额', '自施收益', '自施工核定收益额（万元）']),
                            totalProfit: colIndex(['合计预计结算收益合理值', '合计结算收益', '结算收益']),
                            rate: colIndex(['预计结算收益率', '结算收益率']),
                            completionDate: colIndex(['考核完工日期', '完工日期', '考核完工']),
                            overdue: colIndex(['超完工日期月数', '超期月数', '超期'])
                        };

                        if (idx.name === -1) {
                            reject(new Error('未找到“项目名称”列，请检查表头'));
                            return;
                        }

                        const projects = [];
                        let isInternalSection = false;

                        for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
                            const row = rawRows[r];
                            const firstCell = String(row[0] || '').trim();
                            if (firstCell.includes('内控结算指标')) {
                                isInternalSection = true;
                                continue;
                            }

                            const nameVal = (idx.name !== -1) ? String(row[idx.name] || '').trim() : '';
                            if (!nameVal) continue;
                            if (['华南', '合计', '总计', '小计'].includes(nameVal)) continue;

                            let idVal = (idx.id !== -1) ? row[idx.id] : '';
                            if (idVal === undefined || idVal === null) idVal = '';
                            const idNum = parseInt(idVal);
                            if (isNaN(idNum)) continue;

                            const manager = (idx.manager !== -1) ? String(row[idx.manager] || '').trim() : '';
                            const bizManager = (idx.bizManager !== -1) ? String(row[idx.bizManager] || '').trim() : '';
                            const sector = (idx.sector !== -1) ? String(row[idx.sector] || '').trim() : '总包';

                            const selfTotal = (idx.selfTotal !== -1) ? safeNum(row[idx.selfTotal]) : 0;
                            const total = (idx.total !== -1) ? safeNum(row[idx.total]) : 0;
                            const selfRate = (idx.selfRate !== -1) ? safeNum(row[idx.selfRate]) : null;
                            const selfProfit = (idx.selfProfit !== -1) ? safeNum(row[idx.selfProfit]) : null;
                            const totalProfit = (idx.totalProfit !== -1) ? safeNum(row[idx.totalProfit]) : null;
                            const rate = (idx.rate !== -1) ? safeNum(row[idx.rate]) : null;
                            let completionDate = (idx.completionDate !== -1) ? row[idx.completionDate] : '';
                            if (completionDate instanceof Date) {
                                completionDate = completionDate.toISOString().slice(0, 10);
                            } else if (typeof completionDate === 'number') {
                                const d = new Date((completionDate - 25569) * 86400 * 1000);
                                if (!isNaN(d.getTime())) completionDate = d.toISOString().slice(0, 10);
                            } else {
                                completionDate = formatDate(completionDate);
                            }
                            const overdue = (idx.overdue !== -1) ? safeNum(row[idx.overdue]) : null;

                            const isInternal = isInternalSection || nameVal.includes('内控');

                            projects.push({
                                id: idNum,
                                name: nameVal,
                                manager: manager,
                                bizManager: bizManager,
                                sector: sector,
                                selfTotal: selfTotal !== null ? selfTotal : 0,
                                total: total !== null ? total : 0,
                                selfRate: selfRate,
                                selfProfit: selfProfit,
                                totalProfit: totalProfit,
                                rate: rate,
                                completionDate: completionDate,
                                overdue: overdue,
                                isInternal: isInternal
                            });
                        }

                        if (projects.length === 0) {
                            reject(new Error('未解析到任何项目数据，请检查表格中是否有“项目名称”列且包含有效数据'));
                            return;
                        }

                        projects.sort((a, b) => a.id - b.id);
                        resolve(projects);
                    } catch (err) {
                        reject(new Error('解析出错: ' + err.message));
                    }
                };
                reader.onerror = function() { reject(new Error('文件读取失败')); };
                reader.readAsArrayBuffer(file);
            });
        }

        // ============================================================
        //  文件上传
        // ============================================================
        document.getElementById('fileInput').addEventListener('change', function(e) {
            const file = this.files[0];
            if (!file) return;
            const ext = file.name.split('.').pop().toLowerCase();
            if (!['xlsx', 'xls'].includes(ext)) {
                showToast('请上传 .xlsx 或 .xls 格式的 Excel 文件', 'error');
                this.value = '';
                return;
            }
            showToast('正在解析数据…', 'info');
            parseExcelFile(file)
                .then(data => {
                    renderDashboard(data);
                    showToast('✅ 数据更新成功！共 ' + data.length + ' 个项目', 'success');
                    document.getElementById('updateTime').textContent = '📅 ' + file.name;
                    this.value = '';
                })
                .catch(err => {
                    showToast('❌ 解析失败: ' + err.message, 'error');
                    this.value = '';
                });
        });

        // ============================================================
        //  初始化
        // ============================================================
        renderDashboard(DEFAULT_DATA);
        showToast('📊 已加载示例数据，自适应电脑/手机', 'info');
        console.log('📊 华南2026竣工项目看板 (响应式适配)');
    </script>

</body>
</html>
