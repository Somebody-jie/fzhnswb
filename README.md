<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.5, user-scalable=yes" />
    <title>华南区域商务部指标看板</title>
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
            background: #f0f6fe;
            color: #1a2639;
            min-height: 100vh;
            display: flex;
        }

        /* ===== 侧边栏 ===== */
        .sidebar {
            width: 220px;
            background: rgba(255, 255, 255, 0.85);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-right: 1px solid rgba(200, 215, 230, 0.3);
            height: 100vh;
            position: sticky;
            top: 0;
            overflow-y: auto;
            padding: 20px 0 30px;
            flex-shrink: 0;
            box-shadow: 2px 0 20px rgba(26, 58, 92, 0.06);
            z-index: 10;
            transition: width 0.3s;
        }
        .sidebar .logo {
            padding: 0 20px 20px;
            font-size: 18px;
            font-weight: 700;
            color: #0b2a44;
            display: flex;
            align-items: center;
            gap: 10px;
            border-bottom: 1px solid rgba(200, 215, 230, 0.3);
            margin-bottom: 16px;
        }
        .sidebar .logo span {
            background: linear-gradient(135deg, #3b7bb0, #5c9ed1);
            color: #fff;
            width: 36px;
            height: 36px;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 18px;
            flex-shrink: 0;
        }
        .sidebar .logo .text {
            font-size: 15px;
        }
        .sidebar .menu-group {
            margin-bottom: 12px;
        }
        .sidebar .menu-group .group-title {
            font-size: 11px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #8aa5bf;
            padding: 8px 20px 4px;
            font-weight: 600;
        }
        .sidebar .menu-item {
            display: block;
            padding: 10px 20px;
            color: #2b4a66;
            text-decoration: none;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            border-left: 3px solid transparent;
        }
        .sidebar .menu-item:hover {
            background: rgba(59, 123, 176, 0.08);
            color: #1a3a5c;
        }
        .sidebar .menu-item.active {
            background: rgba(59, 123, 176, 0.12);
            border-left-color: #3b7bb0;
            color: #0b2a44;
            font-weight: 600;
        }
        .sidebar .menu-item.sub {
            padding-left: 32px;
            font-size: 13px;
            font-weight: 400;
        }
        .sidebar .menu-item.sub.active {
            font-weight: 500;
        }
        .sidebar .menu-item .badge {
            float: right;
            background: #dbe7f7;
            color: #1a3a5c;
            font-size: 10px;
            padding: 0 8px;
            border-radius: 20px;
            line-height: 18px;
        }

        /* ===== 主内容 ===== */
        .main-content {
            flex: 1;
            padding: 20px 24px 40px;
            overflow-y: auto;
            max-height: 100vh;
        }
        .main-content .page-title {
            font-size: 22px;
            font-weight: 700;
            color: #0b2a44;
            margin-bottom: 6px;
        }
        .main-content .page-desc {
            color: #6f8ba5;
            font-size: 14px;
            margin-bottom: 20px;
            border-bottom: 1px solid rgba(200, 215, 230, 0.3);
            padding-bottom: 12px;
        }

        /* ===== 板块容器 ===== */
        .section {
            display: none;
        }
        .section.active {
            display: block;
        }

        /* ===== 卡片 ===== */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .kpi-card {
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(6px);
            border-radius: 16px;
            padding: 18px 18px 14px;
            box-shadow: 0 4px 20px rgba(26, 58, 92, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.7);
            transition: all 0.3s;
        }
        .kpi-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 8px 28px rgba(26, 58, 92, 0.10);
        }
        .kpi-card .label {
            font-size: 12px;
            font-weight: 500;
            color: #6f8ba5;
            text-transform: uppercase;
            letter-spacing: 0.3px;
            margin-bottom: 4px;
        }
        .kpi-card .value {
            font-size: 26px;
            font-weight: 700;
            color: #0b2a44;
            letter-spacing: -0.3px;
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
        }

        /* ===== 图表容器 ===== */
        .charts-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 24px;
        }
        .chart-card {
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(6px);
            border-radius: 16px;
            padding: 16px 16px 4px;
            box-shadow: 0 4px 20px rgba(26, 58, 92, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.7);
        }
        .chart-card .chart-title {
            font-size: 14px;
            font-weight: 600;
            color: #0b2a44;
            padding-left: 4px;
            margin-bottom: 2px;
        }
        .chart-card .chart-sub {
            font-size: 11px;
            color: #8aa5bf;
            padding-left: 4px;
            margin-bottom: 6px;
        }
        .chart-container {
            width: 100%;
            height: 250px;
        }

        /* ===== 表格 ===== */
        .table-card {
            background: rgba(255, 255, 255, 0.8);
            backdrop-filter: blur(6px);
            border-radius: 16px;
            padding: 16px 16px 4px;
            box-shadow: 0 4px 20px rgba(26, 58, 92, 0.06);
            border: 1px solid rgba(255, 255, 255, 0.7);
            margin-bottom: 24px;
            overflow: hidden;
        }
        .table-card .table-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            margin-bottom: 10px;
            padding: 0 4px;
        }
        .table-card .table-header .title {
            font-size: 15px;
            font-weight: 600;
            color: #0b2a44;
        }
        .table-card .table-header .count {
            font-size: 12px;
            color: #6f8ba5;
        }
        .table-wrapper {
            overflow-x: auto;
            border-radius: 12px;
            border: 1px solid rgba(200, 215, 230, 0.3);
            -webkit-overflow-scrolling: touch;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 12px;
            min-width: 700px;
        }
        table thead {
            background: #f6fafd;
        }
        table th {
            padding: 8px 10px;
            text-align: left;
            font-weight: 600;
            color: #2b4a66;
            border-bottom: 2px solid #dce6ef;
            white-space: nowrap;
            font-size: 10px;
            text-transform: uppercase;
            letter-spacing: 0.3px;
        }
        table td {
            padding: 6px 10px;
            border-bottom: 1px solid #eef4fa;
            color: #1a2639;
            white-space: nowrap;
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
        table tbody tr:hover {
            background: #f0f7fe;
        }
        .status-badge {
            display: inline-block;
            padding: 1px 10px;
            border-radius: 20px;
            font-size: 10px;
            font-weight: 600;
        }
        .status-badge.win {
            background: #dff0e5;
            color: #1e6b34;
        }
        .status-badge.lose {
            background: #fde8e8;
            color: #b71c1c;
        }
        .status-badge.pending {
            background: #fdf0e0;
            color: #a85d1a;
        }

        /* ===== 上传区域 ===== */
        .upload-area {
            background: rgba(255, 255, 255, 0.6);
            border: 2px dashed rgba(59, 123, 176, 0.3);
            border-radius: 16px;
            padding: 20px;
            text-align: center;
            margin-bottom: 24px;
            cursor: pointer;
            transition: all 0.3s;
        }
        .upload-area:hover {
            background: rgba(255, 255, 255, 0.9);
            border-color: #3b7bb0;
        }
        .upload-area input[type="file"] {
            display: none;
        }
        .upload-area .label {
            font-size: 14px;
            color: #4a6a8a;
        }
        .upload-area .label strong {
            color: #3b7bb0;
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
            transition: opacity 0.4s;
            pointer-events: none;
            z-index: 999;
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

        /* ===== 响应式 ===== */
        @media (max-width: 768px) {
            .sidebar {
                width: 60px;
                padding: 12px 0;
            }
            .sidebar .logo span {
                margin: 0 auto;
            }
            .sidebar .logo .text {
                display: none;
            }
            .sidebar .menu-item {
                padding: 10px 12px;
                font-size: 12px;
                text-align: center;
                border-left: none;
                border-bottom: 2px solid transparent;
            }
            .sidebar .menu-item .badge {
                display: none;
            }
            .sidebar .menu-item.sub {
                padding-left: 12px;
                font-size: 11px;
            }
            .sidebar .menu-group .group-title {
                text-align: center;
                font-size: 9px;
                padding: 4px 0;
            }
            .main-content {
                padding: 12px 16px;
                max-height: none;
            }
            .charts-grid {
                grid-template-columns: 1fr;
            }
            .kpi-grid {
                grid-template-columns: 1fr 1fr;
                gap: 12px;
            }
            .kpi-card .value {
                font-size: 22px;
            }
        }
        @media (max-width: 480px) {
            .sidebar {
                width: 50px;
            }
            .sidebar .menu-item {
                padding: 8px 6px;
                font-size: 10px;
            }
            .main-content {
                padding: 8px 10px;
            }
            .kpi-grid {
                grid-template-columns: 1fr 1fr;
                gap: 8px;
            }
            .kpi-card {
                padding: 12px;
            }
            .kpi-card .value {
                font-size: 18px;
            }
        }
        /* 内控标签样式 */
        .internal-tag {
            background: #5c9ed1;
            color: #fff;
            font-size: 9px;
            padding: 1px 8px;
            border-radius: 40px;
            margin-left: 6px;
        }
        .sector-badge {
            display: inline-block;
            padding: 2px 10px;
            border-radius: 40px;
            font-size: 10px;
            font-weight: 600;
            background: #e9eef2;
            color: #2b3f54;
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
        .row-internal {
            background-color: #eef6fe;
        }
        .row-internal:hover {
            background-color: #e2eefa;
        }
        .edit-hint {
            font-size: 11px;
            color: #8aa5bf;
            font-weight: 400;
            margin-left: 8px;
        }
    </style>
</head>
<body>

    <!-- ===== 侧边栏 ===== -->
    <nav class="sidebar" id="sidebar">
        <div class="logo">
            <span>📊</span>
            <span class="text">综合看板</span>
        </div>
        <div class="menu-group">
            <div class="group-title">📌 主模块</div>
            <div class="menu-item active" data-section="jieSuan">竣工结算</div>
            <div class="menu-item" data-section="touBiao">投标台账</div>
        </div>
        <div class="menu-group" id="tenderSubMenu" style="display:none;">
            <div class="group-title">📂 投标台账</div>
            <div class="menu-item sub active" data-sub="2026">2026年</div>
            <div class="menu-item sub" data-sub="2025">2025年</div>
            <div class="menu-item sub" data-sub="2024">2024年</div>
        </div>
    </nav>

    <!-- ===== 主内容 ===== -->
    <div class="main-content" id="mainContent">
        <!-- 上传区域 -->
        <div class="upload-area" id="uploadArea">
            <div class="label">📁 点击或拖拽上传 <strong>Excel 文件</strong>（支持 .xlsx / .xls）<br />
                <span style="font-size:12px;color:#8aa5bf;">上传后自动解析所有工作表，更新看板数据</span>
            </div>
            <input type="file" id="fileInput" accept=".xlsx,.xls" />
        </div>

        <!-- 页面主标题 -->
        <div class="page-title" id="pageTitle">竣工项目结算看板</div>
        <div class="page-desc" id="pageDesc">数据统计截止 2026.3.10 · 共 <span id="headerCount">0</span> 个项目</div>

        <!-- ===== 板块：竣工结算 ===== -->
        <div class="section active" id="section-jieSuan">
            <!-- KPI -->
            <div class="kpi-grid" id="kpiJieSuan">
                <div class="kpi-card accent-1"><div class="label">📦 项目总数</div><div class="value" id="jsTotal">0</div><div class="trend" id="jsTotalSub">竣工 0 · 内控 0</div></div>
                <div class="kpi-card accent-2"><div class="label">💰 合计结算总额</div><div class="value" id="jsAmount">0 <span class="unit">万元</span></div><div class="trend">自施 <span id="jsSelfAmount">0</span> 万元</div></div>
                <div class="kpi-card accent-3"><div class="label">📈 平均收益率</div><div class="value" id="jsRate">0%</div><div class="trend">自施核定 <span id="jsSelfRate">0%</span></div></div>
                <div class="kpi-card accent-4"><div class="label">⏰ 超期项目</div><div class="value" id="jsOverdue">0</div><div class="trend">最长 <span id="jsMaxOverdue">0</span> 个月</div></div>
            </div>
            <!-- 图表 -->
            <div class="charts-grid">
                <div class="chart-card"><div class="chart-title">🏗️ 按版块分布</div><div class="chart-sub">全部项目 · 含内控</div><div class="chart-container" id="chartSector"></div></div>
                <div class="chart-card"><div class="chart-title">📊 项目收益率排行</div><div class="chart-sub">全部项目 · 含内控</div><div class="chart-container" id="chartRate"></div></div>
                <div class="chart-card" style="grid-column:1/-1;"><div class="chart-title">⏳ 超期月数排行</div><div class="chart-sub">全部项目 · 含内控</div><div class="chart-container tall" id="chartOverdue" style="height:280px;"></div></div>
            </div>
            <!-- 表格 -->
            <div class="table-card">
                <div class="table-header"><span class="title">📋 项目明细 <span class="edit-hint">（点击单元格可编辑，修改后自动更新指标）</span></span><span class="count" id="jsTableCount">共 0 个项目</span></div>
                <div class="table-wrapper"><table><thead><tr><th>序号</th><th>项目名称</th><th>项目经理</th><th>商务经理</th><th>版块</th><th>自施总额(万元)</th><th>合计总额(万元)</th><th>自施收益率</th><th>自施收益</th><th>结算收益</th><th>结算收益率</th><th>完工日期</th><th>超期月数</th></tr></thead><tbody id="jsTableBody"></tbody></table></div>
            </div>
        </div>

        <!-- ===== 板块：投标台账 ===== -->
        <div class="section" id="section-touBiao">
            <div id="tenderContent"></div>
        </div>
    </div>

    <!-- ===== Toast ===== -->
    <div class="toast" id="toast"></div>

    <script>
        // ============================================================
        //  1. 默认示例数据（竣工项目）
        // ============================================================
        const DEFAULT_JIESUAN = [{
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

        // 默认投标台账示例
        const DEFAULT_TENDER = {
            '2026': [],
            '2025': [],
            '2024': []
        };

        // ============================================================
        //  2. 全局状态
        // ============================================================
        let currentData = {
            jieSuan: DEFAULT_JIESUAN,
            tender: DEFAULT_TENDER
        };
        let chartInstances = {};
        let currentSection = 'jieSuan';
        let currentTenderSub = '2026';

        // ============================================================
        //  3. 工具函数
        // ============================================================
        function formatNum(v, d = 2) {
            if (v === undefined || v === null || isNaN(v)) return '—';
            return Number(v).toFixed(d);
        }

        function formatRate(v) {
            if (v === undefined || v === null || isNaN(v)) return '—';
            return (v * 100).toFixed(2) + '%';
        }

        function formatDate(d) {
            if (!d) return '—';
            if (d instanceof Date) return d.toISOString().slice(0, 10);
            const s = String(d);
            if (s.includes('-')) {
                const parts = s.split(' ');
                return parts[0];
            }
            return s;
        }

        function safeNum(v) {
            if (v === undefined || v === null || v === '') return null;
            if (typeof v === 'string' && v.includes('%')) v = v.replace(/%/g, '').trim();
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
            toastTimer = setTimeout(() => { el.classList.remove('show'); }, 3500);
        }

        // ============================================================
        //  4. 侧边栏切换逻辑
        // ============================================================
        document.querySelectorAll('.menu-item[data-section]').forEach(item => {
            item.addEventListener('click', function() {
                const section = this.dataset.section;
                document.querySelectorAll('.menu-item[data-section]').forEach(el => el.classList.remove('active'));
                this.classList.add('active');
                document.querySelectorAll('.section').forEach(el => el.classList.remove('active'));
                document.getElementById('section-' + section).classList.add('active');
                const subMenu = document.getElementById('tenderSubMenu');
                if (section === 'touBiao') {
                    subMenu.style.display = 'block';
                    document.querySelectorAll('#tenderSubMenu .menu-item.sub').forEach(el => el.classList.remove('active'));
                    document.querySelector('#tenderSubMenu .menu-item.sub[data-sub="2026"]').classList.add('active');
                    currentTenderSub = '2026';
                    renderTenderSub(currentTenderSub);
                } else {
                    subMenu.style.display = 'none';
                }
                currentSection = section;
                document.getElementById('pageTitle').textContent = section === 'jieSuan' ? '竣工项目结算看板' : '投标台账看板';
                setTimeout(() => { Object.values(chartInstances).forEach(c => { if (c && c.resize) c.resize(); }); }, 100);
            });
        });

        document.querySelectorAll('#tenderSubMenu .menu-item.sub').forEach(item => {
            item.addEventListener('click', function() {
                document.querySelectorAll('#tenderSubMenu .menu-item.sub').forEach(el => el.classList.remove('active'));
                this.classList.add('active');
                currentTenderSub = this.dataset.sub;
                renderTenderSub(currentTenderSub);
            });
        });

        // ============================================================
        //  5. 渲染竣工结算板块
        // ============================================================
        function renderJieSuan(data) {
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

            document.getElementById('jsTotal').textContent = totalCount;
            document.getElementById('jsTotalSub').textContent = `竣工 ${totalMain} · 内控 ${internalCount}`;
            document.getElementById('jsAmount').innerHTML = formatNum(sumTotal, 1) + ' <span class="unit">万元</span>';
            document.getElementById('jsSelfAmount').textContent = formatNum(sumSelf, 1);
            document.getElementById('jsRate').textContent = formatRate(avgRate);
            document.getElementById('jsSelfRate').textContent = formatRate(avgSelfRate);
            document.getElementById('jsOverdue').textContent = overdueProjects.length;
            document.getElementById('jsMaxOverdue').textContent = formatNum(maxOverdue, 1);

            document.getElementById('headerCount').textContent = totalCount;
            document.getElementById('jsTableCount').textContent = '共 ' + allProjects.length + ' 个项目';

            renderJieSuanTable(allProjects);
            renderJieSuanCharts(allProjects);
        }

        function renderJieSuanTable(data) {
            const tbody = document.getElementById('jsTableBody');
            if (!data.length) { tbody.innerHTML =
                '<tr><td colspan="13" style="text-align:center;padding:20px;color:#6f8ba5;">暂无数据</td></tr>'; return; }
            let html = '';
            data.forEach((d, index) => {
                const internalTag = d.isInternal ? '<span class="internal-tag">内控</span>' : '';
                const sectorCls = getSectorClass(d.sector);
                const rateCls = (d.rate !== null && !isNaN(d.rate)) ? (d.rate >= 0 ? 'rate-positive' : 'rate-negative') :
                    'rate-neutral';
                const overdueCls = (d.overdue !== null && !isNaN(d.overdue)) ?
                    (d.overdue > 24 ? 'overdue-high' : d.overdue > 12 ? 'overdue-mid' : 'overdue-low') :
                    'rate-neutral';
                const editableAttr = (field) => `contenteditable="true" data-field="${field}" data-row="${index}"`;
                html += `<tr class="${d.isInternal ? 'row-internal' : ''}">
                            <td>${d.id}</td>
                            <td><strong>${d.name}</strong>${internalTag}</td>
                            <td ${editableAttr('manager')}>${d.manager || ''}</td>
                            <td ${editableAttr('bizManager')}>${d.bizManager || ''}</td>
                            <td ${editableAttr('sector')}><span class="sector-badge ${sectorCls}">${d.sector || ''}</span></td>
                            <td ${editableAttr('selfTotal')}>${d.selfTotal !== null && !isNaN(d.selfTotal) ? d.selfTotal : ''}</td>
                            <td ${editableAttr('total')}>${d.total !== null && !isNaN(d.total) ? d.total : ''}</td>
                            <td ${editableAttr('selfRate')}>${d.selfRate !== null && !isNaN(d.selfRate) ? d.selfRate : ''}</td>
                            <td ${editableAttr('selfProfit')}>${d.selfProfit !== null && !isNaN(d.selfProfit) ? d.selfProfit : ''}</td>
                            <td ${editableAttr('totalProfit')}>${d.totalProfit !== null && !isNaN(d.totalProfit) ? d.totalProfit : ''}</td>
                            <td ${editableAttr('rate')} class="${rateCls}">${d.rate !== null && !isNaN(d.rate) ? d.rate : ''}</td>
                            <td ${editableAttr('completionDate')}>${formatDate(d.completionDate)}</td>
                            <td ${editableAttr('overdue')} class="${overdueCls}">${d.overdue !== null && !isNaN(d.overdue) ? d.overdue : ''}</td>
                        </tr>`;
            });
            tbody.innerHTML = html;

            // 绑定编辑事件
            tbody.removeEventListener('input', onJieSuanEdit);
            tbody.addEventListener('input', onJieSuanEdit);
        }

        // 竣工结算表格编辑事件
        function onJieSuanEdit(e) {
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
                    renderJieSuanTable(currentData.jieSuan);
                    return;
                }
            } else {
                parsedValue = newValue;
            }
            if (rowIdx < currentData.jieSuan.length) {
                const oldVal = currentData.jieSuan[rowIdx][field];
                if (parsedValue === oldVal) return;
                currentData.jieSuan[rowIdx][field] = parsedValue;
                clearTimeout(window._editTimer);
                window._editTimer = setTimeout(() => {
                    renderJieSuan(currentData.jieSuan);
                    showToast('数据已更新，指标同步刷新', 'success');
                }, 300);
            }
        }

        function renderJieSuanCharts(projects) {
            // 版块饼图
            const sectorMap = {};
            projects.forEach(d => {
                const sec = d.sector || '未分类';
                if (!sectorMap[sec]) sectorMap[sec] = { count: 0, total: 0 };
                sectorMap[sec].count += 1;
                sectorMap[sec].total += d.total || 0;
            });
            const sectorColors = { '总包': '#3b7bb0', '基础设施': '#4caf84', '机电': '#f5a35b', '未分类': '#a0b8cc' };
            const pieData = Object.keys(sectorMap).map(name => ({ name, value: sectorMap[name].total, count: sectorMap[name]
                    .count }));
            const chart1 = initChart('chartSector');
            chart1.setOption({
                tooltip: { trigger: 'item', formatter: function(p) { const d = p.data; return `<strong>${d.name}</strong><br/>结算总额: ${d.value.toFixed(1)} 万元<br/>项目数: ${d.count} 个<br/>占比: ${p.percent}%`; } },
                legend: { orient: 'vertical', right: '5%', top: 'center', textStyle: { fontSize: 12, color: '#2b4a66' },
                    formatter: function(name) { const item = sectorMap[name]; return name + '  (' + item.count + '个)'; } },
                series: [{
                    type: 'pie',
                    radius: ['42%', '72%'],
                    center: ['42%', '50%'],
                    avoidLabelOverlap: true,
                    itemStyle: { borderRadius: 8, borderColor: '#fff', borderWidth: 2 },
                    label: { show: true, formatter: '{b}\n{d}%', fontSize: 11, color: '#1a2639' },
                    labelLine: { length: 12, length2: 16 },
                    emphasis: { scale: true, scaleSize: 8 },
                    data: pieData.map(d => ({ ...d, itemStyle: { color: sectorColors[d.name] || '#a0b8cc' } }))
                }]
            });
            chartInstances['chartSector'] = chart1;

            // 收益率排行
            const rateData = projects.filter(d => d.rate !== null && !isNaN(d.rate)).sort((a, b) => b.rate - a.rate);
            const chart2 = initChart('chartRate');
            if (!rateData.length) {
                chart2.setOption({ title: { text: '暂无收益率数据', left: 'center', top: 'center', textStyle: { color: '#8aa5bf',
                            fontSize: 14 } } });
            } else {
                const names = rateData.map(d => d.name.length > 12 ? d.name.slice(0, 12) + '…' : d.name);
                const values = rateData.map(d => +(d.rate * 100).toFixed(2));
                const colors = values.map(v => v >= 0 ? '#3b7bb0' : '#e57373');
                chart2.setOption({
                    tooltip: { trigger: 'axis', formatter: function(params) { const p = params[0]; const idx = p.dataIndex;
                            const d = rateData[idx]; return `<strong>${d.name}</strong><br/>结算收益率: ${p.value.toFixed(2)}%<br/>结算收益: ${d.totalProfit !== null && !isNaN(d.totalProfit) ? d.totalProfit.toFixed(2) : '—'} 万元`; } },
                    grid: { left: '6%', right: '6%', top: '10%', bottom: '18%' },
                    xAxis: { type: 'category', data: names, axisLabel: { rotate: 35, fontSize: 10, color: '#4a6a8a',
                            interval: 0 }, axisLine: { lineStyle: { color: '#dce6ef' } } },
                    yAxis: { type: 'value', name: '收益率 (%)', nameTextStyle: { fontSize: 11, color: '#6f8ba5' },
                        axisLabel: { fontSize: 11, color: '#4a6a8a', formatter: '{value}%' },
                        splitLine: { lineStyle: { color: '#eef4fa', type: 'dashed' } } },
                    series: [{
                        type: 'bar',
                        data: values.map((v, i) => ({ value: v, itemStyle: { color: colors[i], borderRadius: [6,
                                    6, 0, 0] } })),
                        barWidth: '46%',
                        label: { show: true, position: 'top', formatter: function(p) { return p.value.toFixed(1) +
                                '%'; }, fontSize: 10, color: '#1a2639', fontWeight: 500 }
                    }]
                });
            }
            chartInstances['chartRate'] = chart2;

            // 超期排行
            const overdueData = projects.filter(d => d.overdue !== null && !isNaN(d.overdue) && d.overdue > 0).sort((a, b) => b
                .overdue - a.overdue);
            const chart3 = initChart('chartOverdue');
            if (!overdueData.length) {
                chart3.setOption({ title: { text: '暂无超期项目', left: 'center', top: 'center', textStyle: { color: '#8aa5bf',
                            fontSize: 14 } } });
            } else {
                const names = overdueData.map(d => d.name.length > 14 ? d.name.slice(0, 14) + '…' : d.name);
                const values = overdueData.map(d => +(d.overdue).toFixed(1));
                const colors = values.map(v => v > 36 ? '#e57373' : v > 18 ? '#f5a35b' : '#3b7bb0');
                chart3.setOption({
                    tooltip: { trigger: 'axis', formatter: function(params) { const p = params[0]; const idx = p.dataIndex;
                            const d = overdueData[idx]; return `<strong>${d.name}</strong><br/>超期月数: ${p.value.toFixed(1)} 个月<br/>考核完工: ${formatDate(d.completionDate)}`; } },
                    grid: { left: '6%', right: '6%', top: '10%', bottom: '18%' },
                    xAxis: { type: 'category', data: names, axisLabel: { rotate: 35, fontSize: 10, color: '#4a6a8a',
                            interval: 0 }, axisLine: { lineStyle: { color: '#dce6ef' } } },
                    yAxis: { type: 'value', name: '超期月数', nameTextStyle: { fontSize: 11, color: '#6f8ba5' },
                        axisLabel: { fontSize: 11, color: '#4a6a8a', formatter: '{value}月' },
                        splitLine: { lineStyle: { color: '#eef4fa', type: 'dashed' } } },
                    series: [{
                        type: 'bar',
                        data: values.map((v, i) => ({ value: v, itemStyle: { color: colors[i], borderRadius: [6,
                                    6, 0, 0] } })),
                        barWidth: '46%',
                        label: { show: true, position: 'top', formatter: function(p) { return p.value.toFixed(1) +
                                '月'; }, fontSize: 10, color: '#1a2639', fontWeight: 500 }
                    }]
                });
            }
            chartInstances['chartOverdue'] = chart3;
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
        //  6. 渲染投标台账子板块（仅年份） + 可编辑表格
        // ============================================================
        function renderTenderSub(subKey) {
            const container = document.getElementById('tenderContent');
            const data = currentData.tender[subKey] || [];

            if (!data.length) {
                container.innerHTML =
                    `<div class="table-card"><div class="table-header"><span class="title">${subKey}年投标台账 <span class="edit-hint">（点击单元格可编辑，修改后自动更新指标）</span></span><span class="count">暂无数据</span></div></div>`;
                return;
            }

            const total = data.length;
            const win = data.filter(d => d.status && d.status.includes('中标')).length;
            const lose = data.filter(d => d.status && d.status.includes('未中标')).length;
            const pending = total - win - lose;
            let rates = data.map(d => safeNum(d.rate)).filter(r => r !== null && !isNaN(r));
            let avgRate = rates.length ? rates.reduce((a, b) => a + b, 0) / rates.length : 0;

            let html = `
                        <div class="kpi-grid">
                            <div class="kpi-card"><div class="label">📋 项目数</div><div class="value">${total}</div><div class="trend">已投项目</div></div>
                            <div class="kpi-card"><div class="label">🏆 中标</div><div class="value">${win}</div><div class="trend">中标率 ${total? (win/total*100).toFixed(1) : 0}%</div></div>
                            <div class="kpi-card"><div class="label">📉 未中标</div><div class="value">${lose}</div><div class="trend">待定 ${pending}</div></div>
                            <div class="kpi-card"><div class="label">📈 平均收益率</div><div class="value">${formatRate(avgRate)}</div><div class="trend">基于可用数据</div></div>
                        </div>
                        <div class="charts-grid">
                            <div class="chart-card"><div class="chart-title">🏷️ 中标状态分布</div><div class="chart-sub">${subKey}年</div><div class="chart-container" id="tenderStatusChart" style="height:220px;"></div></div>
                            <div class="chart-card"><div class="chart-title">📊 项目类别分布</div><div class="chart-sub">${subKey}年</div><div class="chart-container" id="tenderCategoryChart" style="height:220px;"></div></div>
                        </div>
                        <div class="table-card">
                            <div class="table-header"><span class="title">📋 项目明细 <span class="edit-hint">（点击单元格可编辑，修改后自动更新指标）</span></span><span class="count">共 ${total} 个项目</span></div>
                            <div class="table-wrapper"><table><thead><tr><th>序号</th><th>项目名称</th><th>招标单位</th><th>估算金额(万元)</th><th>下浮率</th><th>收益率</th><th>中标情况</th><th>备注</th></tr></thead><tbody>`;
            data.forEach((d, i) => {
                const statusClass = d.status && d.status.includes('中标') ? 'win' : (d.status && d.status.includes('未中标') ?
                    'lose' : 'pending');
                const statusLabel = d.status || '待定';
                // 可编辑属性
                const editableAttr = (field) => `contenteditable="true" data-field="${field}" data-row="${i}"`;
                html += `<tr>
                            <td>${i+1}</td>
                            <td ${editableAttr('name')}>${d.name||''}</td>
                            <td ${editableAttr('unit')}>${d.unit||''}</td>
                            <td ${editableAttr('estimate')}>${d.estimate||''}</td>
                            <td ${editableAttr('rateDown')}>${d.rateDown||''}</td>
                            <td ${editableAttr('rate')}>${d.rate||''}</td>
                            <td ${editableAttr('status')}><span class="status-badge ${statusClass}">${statusLabel}</span></td>
                            <td ${editableAttr('remark')}>${d.remark||''}</td>
                        </tr>`;
            });
            html += `</tbody></table></div></div>`;
            container.innerHTML = html;

            // 绑定编辑事件
            const tbody = container.querySelector('table tbody');
            if (tbody) {
                tbody.removeEventListener('input', onTenderEdit);
                tbody.addEventListener('input', onTenderEdit);
            }

            // 渲染图表
            setTimeout(() => {
                // 状态饼图
                const statusChart = echarts.init(document.getElementById('tenderStatusChart'));
                statusChart.setOption({
                    tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
                    legend: { orient: 'vertical', right: '5%', top: 'center', textStyle: { fontSize: 11 } },
                    series: [{
                        type: 'pie',
                        radius: ['40%', '70%'],
                        center: ['40%', '50%'],
                        avoidLabelOverlap: true,
                        itemStyle: { borderRadius: 6, borderColor: '#fff', borderWidth: 2 },
                        label: { show: true, formatter: '{b}\n{d}%', fontSize: 10 },
                        labelLine: { length: 10, length2: 12 },
                        data: [
                            { value: win, name: '中标', itemStyle: { color: '#4caf84' } },
                            { value: lose, name: '未中标', itemStyle: { color: '#e57373' } },
                            { value: pending, name: '待定', itemStyle: { color: '#f5a35b' } }
                        ]
                    }]
                });
                // 类别分布柱状图
                const catMap = {};
                data.forEach(d => {
                    const cat = d.category || '其他';
                    catMap[cat] = (catMap[cat] || 0) + 1;
                });
                const catNames = Object.keys(catMap);
                const catValues = catNames.map(k => catMap[k]);
                const catChart = echarts.init(document.getElementById('tenderCategoryChart'));
                catChart.setOption({
                    tooltip: { trigger: 'axis', formatter: function(params) { const p = params[0]; return p.name +
                                ': ' + p.value + ' 个'; } },
                    grid: { left: '8%', right: '8%', top: '10%', bottom: '15%' },
                    xAxis: { type: 'category', data: catNames, axisLabel: { rotate: 25, fontSize: 10,
                            interval: 0 }, axisLine: { lineStyle: { color: '#dce6ef' } } },
                    yAxis: { type: 'value', name: '项目数', nameTextStyle: { fontSize: 11, color: '#6f8ba5' },
                        splitLine: { lineStyle: { color: '#eef4fa', type: 'dashed' } } },
                    series: [{
                        type: 'bar',
                        data: catValues.map(v => ({ value: v, itemStyle: { color: '#3b7bb0',
                                borderRadius: [4, 4, 0, 0] } })),
                        barWidth: '40%',
                        label: { show: true, position: 'top', fontSize: 10 }
                    }]
                });
                chartInstances['tenderStatus'] = statusChart;
                chartInstances['tenderCategory'] = catChart;
                window.addEventListener('resize', () => {
                    if (statusChart) statusChart.resize();
                    if (catChart) catChart.resize();
                });
            }, 100);
        }

        // 投标台账表格编辑事件
        function onTenderEdit(e) {
            const cell = e.target;
            if (!cell.hasAttribute('contenteditable')) return;
            const field = cell.dataset.field;
            const rowIdx = parseInt(cell.dataset.row);
            if (isNaN(rowIdx) || !field) return;
            let newValue = cell.innerText.trim();
            // 对数值字段做类型转换
            const numericFields = ['estimate', 'rate'];
            let parsedValue;
            if (numericFields.includes(field)) {
                parsedValue = safeNum(newValue);
                if (parsedValue === null && newValue !== '') {
                    showToast('请输入有效数字', 'error');
                    renderTenderSub(currentTenderSub);
                    return;
                }
            } else {
                parsedValue = newValue;
            }
            const year = currentTenderSub;
            const data = currentData.tender[year];
            if (rowIdx < data.length) {
                const oldVal = data[rowIdx][field];
                if (parsedValue === oldVal) return;
                data[rowIdx][field] = parsedValue;
                clearTimeout(window._tenderEditTimer);
                window._tenderEditTimer = setTimeout(() => {
                    renderTenderSub(year);
                    showToast('投标数据已更新，指标同步刷新', 'success');
                }, 300);
            }
        }

        // ============================================================
        //  7. 解析Excel（综合）
        // ============================================================
        function parseExcelFile(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = function(e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                        const result = {
                            jieSuan: [],
                            tender: { '2026': [], '2025': [], '2024': [] }
                        };

                        function isSkipTitle(name) {
                            if (!name) return true;
                            const trimmed = name.trim();
                            const exactMatch = ['已投项目', '在投项目', '投标项目', '配合项目', '弃标', '弃投', '二', '三', '四', '五', '六', '七', '八', '九',
                                '十'
                            ];
                            if (exactMatch.includes(trimmed)) return true;
                            if (/^[一二三四五六七八九十]+[、.．\s]/.test(trimmed)) return true;
                            if (/已投项目|在投项目/.test(trimmed) && trimmed.length <= 10) return true;
                            return false;
                        }

                        workbook.SheetNames.forEach(sheetName => {
                            const sheet = workbook.Sheets[sheetName];
                            const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', header: 1 });
                            if (rows.length < 2) return;

                            let type = '';
                            if (sheetName.includes('竣工') || sheetName.includes('结算')) type = 'jiesuan';
                            else if (sheetName.includes('2026')) type = 'tender2026';
                            else if (sheetName.includes('2025')) type = 'tender2025';
                            else if (sheetName.includes('2024')) type = 'tender2024';
                            else {
                                const firstRow = rows[0].map(c => String(c).trim());
                                if (firstRow.some(h => /项目名称/.test(h))) {
                                    if (firstRow.some(h => /竣工|结算/.test(h))) type = 'jiesuan';
                                    else type = 'tender2026';
                                }
                            }

                            if (type === 'jiesuan') {
                                let headerIdx = -1;
                                for (let r = 0; r < Math.min(rows.length, 30); r++) {
                                    if (rows[r].some(c => /项目名称/.test(String(c).trim()))) { headerIdx = r; break; }
                                }
                                if (headerIdx === -1) return;
                                const header = rows[headerIdx].map(c => String(c).trim());
                                const col = (kw) => {
                                    for (let i = 0; i < header.length; i++) {
                                        if (header[i].includes(kw)) return i;
                                    }
                                    return -1;
                                };
                                const idx = {
                                    id: col('序号'),
                                    name: col('项目名称'),
                                    manager: col('项目经理'),
                                    biz: col('商务经理'),
                                    sector: col('版块'),
                                    selfTotal: col('自施预计结算总额合理值'),
                                    total: col('合计预计结算总额合理值'),
                                    selfRate: col('自施核定收益率'),
                                    selfProfit: col('自施工核定收益额'),
                                    totalProfit: col('合计预计结算收益合理值'),
                                    rate: col('预计结算收益率'),
                                    completion: col('考核完工日期'),
                                    overdue: col('超完工日期月数')
                                };
                                if (idx.name === -1) return;
                                for (let r = headerIdx + 1; r < rows.length; r++) {
                                    const row = rows[r];
                                    const name = (idx.name !== -1) ? String(row[idx.name] || '').trim() : '';
                                    if (!name || ['华南', '合计', '总计'].includes(name)) continue;
                                    if (isSkipTitle(name)) continue;
                                    const idNum = (idx.id !== -1) ? parseInt(row[idx.id]) : (r - headerIdx);
                                    if (isNaN(idNum)) continue;
                                    const isInternal = name.includes('内控') || sheetName.includes('内控');
                                    const p = {
                                        id: idNum,
                                        name: name,
                                        manager: (idx.manager !== -1) ? String(row[idx.manager] || '').trim() : '',
                                        bizManager: (idx.biz !== -1) ? String(row[idx.biz] || '').trim() : '',
                                        sector: (idx.sector !== -1) ? String(row[idx.sector] || '').trim() : '总包',
                                        selfTotal: (idx.selfTotal !== -1) ? safeNum(row[idx.selfTotal]) || 0 : 0,
                                        total: (idx.total !== -1) ? safeNum(row[idx.total]) || 0 : 0,
                                        selfRate: (idx.selfRate !== -1) ? safeNum(row[idx.selfRate]) : null,
                                        selfProfit: (idx.selfProfit !== -1) ? safeNum(row[idx.selfProfit]) : null,
                                        totalProfit: (idx.totalProfit !== -1) ? safeNum(row[idx.totalProfit]) :
                                            null,
                                        rate: (idx.rate !== -1) ? safeNum(row[idx.rate]) : null,
                                        completionDate: (idx.completion !== -1) ? formatDate(row[idx.completion]) :
                                            '',
                                        overdue: (idx.overdue !== -1) ? safeNum(row[idx.overdue]) : null,
                                        isInternal: isInternal
                                    };
                                    result.jieSuan.push(p);
                                }
                                return;
                            }

                            if (type.startsWith('tender')) {
                                let year = type.replace('tender', '');
                                let headerIdx = -1;
                                for (let r = 0; r < Math.min(rows.length, 30); r++) {
                                    if (rows[r].some(c => /项目名称/.test(String(c).trim()))) { headerIdx = r; break; }
                                }
                                if (headerIdx === -1) return;
                                const header = rows[headerIdx].map(c => String(c).trim());
                                const col = (kw) => {
                                    for (let i = 0; i < header.length; i++) {
                                        if (header[i].includes(kw)) return i;
                                    }
                                    return -1;
                                };
                                const idx = {
                                    name: col('项目名称'),
                                    unit: col('招标单位'),
                                    estimate: col('招标估算'),
                                    rateDown: col('投标净下浮率'),
                                    rate: col('标前收益率'),
                                    status: col('中标情况'),
                                    remark: col('备注'),
                                    category: col('项目类别')
                                };
                                if (idx.name === -1) return;
                                for (let r = headerIdx + 1; r < rows.length; r++) {
                                    const row = rows[r];
                                    const name = (idx.name !== -1) ? String(row[idx.name] || '').trim() : '';
                                    if (!name || ['合计', '总计', '一', '二'].includes(name)) continue;
                                    if (isSkipTitle(name)) continue;
                                    const p = {
                                        name: name,
                                        unit: (idx.unit !== -1) ? String(row[idx.unit] || '').trim() : '',
                                        estimate: (idx.estimate !== -1) ? safeNum(row[idx.estimate]) : '',
                                        rateDown: (idx.rateDown !== -1) ? row[idx.rateDown] : '',
                                        rate: (idx.rate !== -1) ? safeNum(row[idx.rate]) : '',
                                        status: (idx.status !== -1) ? String(row[idx.status] || '').trim() : '',
                                        remark: (idx.remark !== -1) ? String(row[idx.remark] || '').trim() : '',
                                        category: (idx.category !== -1) ? String(row[idx.category] || '').trim() : ''
                                    };
                                    result.tender[year].push(p);
                                }
                                return;
                            }
                        });

                        resolve(result);
                    } catch (err) {
                        reject(err);
                    }
                };
                reader.onerror = function() { reject(new Error('文件读取失败')); };
                reader.readAsArrayBuffer(file);
            });
        }

        // ============================================================
        //  8. 文件上传处理
        // ============================================================
        document.getElementById('uploadArea').addEventListener('click', function() {
            document.getElementById('fileInput').click();
        });
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
                .then(result => {
                    if (result.jieSuan.length) currentData.jieSuan = result.jieSuan;
                    else currentData.jieSuan = DEFAULT_JIESUAN;
                    for (let key of ['2026', '2025', '2024']) {
                        if (result.tender[key] && result.tender[key].length) {
                            currentData.tender[key] = result.tender[key];
                        }
                    }
                    if (currentSection === 'jieSuan') {
                        renderJieSuan(currentData.jieSuan);
                    } else {
                        renderTenderSub(currentTenderSub);
                    }
                    showToast('✅ 数据更新成功！', 'success');
                    document.getElementById('pageDesc').textContent = '数据来源: ' + file.name;
                    this.value = '';
                })
                .catch(err => {
                    showToast('❌ 解析失败: ' + err.message, 'error');
                    this.value = '';
                });
        });

        // ============================================================
        //  9. 初始化
        // ============================================================
        document.title = '华南区域商务部指标看板';

        // 植入示例投标数据
        currentData.tender['2026'] = [
            { name: '深圳市未成年人救助保护中心修缮项目', unit: '深圳市建筑工务署', estimate: '1890', rateDown: '10.1%', rate: '0.12',
                status: '待定', category: '房屋建筑' },
            { name: '汕尾市智慧物流园项目', unit: '汕尾市智慧物流投资有限公司', estimate: '28718.76', rateDown: '0.142%', rate: '0.08',
                status: '待定', category: '房屋建筑' },
            { name: '创维创客工业园一期二标段', unit: '深圳创维创客发展有限公司', estimate: '61516.57', rateDown: '—', rate: '0.05',
                status: '待定', category: '房屋建筑' },
            { name: '抖音深圳二期项目土护降及桩基', unit: '深圳今日头条', estimate: '18598', rateDown: '—', rate: '0.10', status: '中标',
                category: '专业工程' }
        ];
        currentData.tender['2025'] = [
            { name: '新布新路市政工程', unit: '深圳市龙岗区建筑工务署', estimate: '13000', rateDown: '16.98%', rate: '0.06',
                status: '未中标', category: '基础设施' },
            { name: '中海壳牌惠州三期乙烯项目', unit: '中海壳牌', estimate: '9795', rateDown: '3%', rate: '0.0506', status: '未中标',
                category: '房屋建筑' },
            { name: '深圳湾超级总部基地片区2标', unit: '深圳市建筑工务署', estimate: '80353.78', rateDown: '26.1%', rate: '0.04',
                status: '未中标', category: '基础设施' }
        ];
        currentData.tender['2024'] = [
            { name: '南方医科大学深圳医院二期', unit: '深圳市建筑工务署', estimate: '146300', rateDown: '23.99%', rate: '0.05',
                status: '未中标', category: '房屋建筑' },
            { name: '深圳职业技术大学西丽校区修缮', unit: '深圳市建筑工务署', estimate: '11761.53', rateDown: '11.12%', rate: '0.08',
                status: '中标', category: '房屋建筑' },
            { name: '香港中文大学（深圳）医学院Ⅱ标', unit: '深圳市建筑工务署', estimate: '108631.59', rateDown: '24.05%', rate: '0.07',
                status: '中标', category: '房屋建筑' }
        ];

        renderJieSuan(DEFAULT_JIESUAN);
        showToast('📊 华南区域商务部指标看板已加载，所有表格均可编辑', 'info');
        console.log('华南区域商务部指标看板启动（投标台账可编辑）');
    </script>
</body>
</html>
