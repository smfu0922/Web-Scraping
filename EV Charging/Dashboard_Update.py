import json
import os
import pandas as pd

# 1. 讀取清洗後的 EV 數據
file_path = r"C:\Users\User\Downloads\ev_scraping_full.xlsx"
sheet = "Cleaned EV Data"

if not os.path.exists(file_path):
    print(f"錯誤：找不到檔案 {file_path}，請確保檔案在同一個目錄下！")
    exit()

df = pd.read_excel(file_path, sheet_name=sheet)

# 2. 數據過濾：完全排除廣告與促銷的內容
initial_count = len(df)
df = df[~df["Purpose"].str.contains("廣告|促銷", na=False, case=False)]
df = df[~df["Theme"].str.contains("廣告|促銷", na=False, case=False)]
print(
    f"數據過濾完成：已排除廣告。從 {initial_count} 筆過濾至 {len(df)} 筆有效評論。"
)

# 填補空值
df["Emotion"] = df["Emotion"].fillna("中性").str.strip()
df["Emotion_Detail"] = df["Emotion_Detail"].fillna("陳述").str.strip()
df["Theme"] = df["Theme"].fillna("其他").str.strip()
df["Location"] = df["Location"].fillna("null").str.strip()
df["EV_Model"] = df["EV_Model"].fillna("null").str.strip()
df["Competitor"] = df["Competitor"].fillna("null").str.strip()
df["Charging_Type"] = df["Charging_Type"].fillna("null").str.strip()
df["Text"] = df["Text"].fillna("").str.strip()
df["User"] = df["User"].fillna("匿名用戶").str.strip()
df["Likes"] = (
    pd.to_numeric(df["Likes"], errors="coerce").fillna(0).astype(int)
)


# 強大防呆：將日期強制清洗為 YYYY-MM-DD
def clean_to_date_str(val):
    if pd.isna(val):
        return "未提及"
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")

    val_str = str(val).strip()
    if val_str.lower() in ["nan", "nat", "none", ""]:
        return "未提及"

    if val_str.isdigit() or (
            val_str.replace(".", "", 1).isdigit()
            and float(val_str).is_integer()
    ):
        try:
            dt = pd.to_datetime(
                int(float(val_str)), unit="D", origin="1899-12-30"
            )
            return dt.strftime("%Y-%m-%d")
        except:
            pass

    try:
        dt = pd.to_datetime(val_str, errors="coerce")
        if not pd.isna(dt):
            return dt.strftime("%Y-%m-%d")
    except:
        pass

    return (
        val_str[:10]
        if len(val_str) >= 10 and val_str[:4].isdigit()
        else val_str
    )


df["Formatted_Time"] = df["Time"].apply(clean_to_date_str)

# 3. 建立香港地區座標對照字典
geo_dict = {
    "荃灣": [22.3686, 114.1114],
    "荃灣京匯中心": [22.3703, 114.1122],
    "海之戀": [22.3670, 114.1110],
    "荃灣馬角街18號": [22.3655, 114.1171],
    "九龍灣": [22.3210, 114.2100],
    "九龍灣大昌行": [22.3242, 114.2091],
    "九龍灣大昌": [22.3242, 114.2091],
    "將軍澳": [22.3120, 114.2560],
    "將軍澳尚德邨": [22.3135, 114.2601],
    "觀塘": [22.3130, 114.2250],
    "鴻圖道6號": [22.3151, 114.2215],
    "觀塘鴻圖道6號": [22.3151, 114.2215],
    "偉業街": [22.3180, 114.2190],
    "大角咀": [22.3220, 114.1620],
    "大角咀 One Bedford Place": [22.3236, 114.1610],
    "奧海城": [22.3175, 114.1602],
    "佐敦": [22.3050, 114.1710],
    "佐敦高臨停車場": [22.3045, 114.1705],
    "佐敦西貢街": [22.3061, 114.1708],
    "尖沙咀": [22.2988, 114.1722],
    "紅磡": [22.3075, 114.1831],
    "深水埗": [22.3300, 114.1620],
    "元朗": [22.4450, 114.0220],
    "元朗新潭路": [22.4720, 114.0530],
    "錦上路站": [22.4340, 114.0620],
    "錦上路": [22.4340, 114.0620],
    "大棠": [22.4170, 114.0220],
    "屯門": [22.3950, 113.9740],
    "火炭": [22.3980, 114.1950],
    "沙田": [22.3830, 114.1830],
    "馬鞍山": [22.4230, 114.2300],
    "馬鞍山頌安停車場": [22.4255, 114.2285],
    "大埔": [22.4500, 114.1650],
    "大埔東昌街": [22.4468, 114.1691],
    "大埔東昌街體育館停車場": [22.4468, 114.1691],
    "粉嶺": [22.4920, 114.1380],
    "聯和墟": [22.4990, 114.1420],
    "香園圍": [22.5530, 114.1560],
    "灣仔": [22.2790, 114.1720],
    "灣仔尚翹峰灣仔街市": [22.2754, 114.1742],
    "金鐘": [22.2780, 114.1650],
    "中環": [22.2820, 114.1580],
    "鰂魚涌": [22.2840, 114.2120],
    "K11 ATELIER King’s Road Carpark": [22.2851, 114.2110],
    "筲箕灣": [22.2780, 114.2290],
    "筲箕灣天悅廣場": [22.2788, 114.2278],
    "數碼港": [22.2610, 114.1300],
    "海洋公園": [22.2467, 114.1757],
    "香港科技大學": [22.3364, 114.2654],
    "機場": [22.3080, 113.9185],
    "西貢": [22.3830, 114.2700],
    "龍蝦灣": [22.3160, 114.2900],
}

HK_CENTER = [22.3527, 114.1272]

records = []
for idx, row in df.iterrows():
    loc_name = row["Location"]
    coords = geo_dict.get(
        loc_name,
        (
            HK_CENTER
            if loc_name != "未提及" and loc_name != "香港"
            else None
        ),
    )

    records.append(
        {
            "id": int(row["ID"]),
            "source": row["Source"],
            "time": row["Formatted_Time"],
            "text": row["Text"],
            "emotion_basic": row["Emotion"],
            "emotion_detailed": row["Emotion_Detail"],
            "theme": row["Theme"],
            "location": loc_name,
            "coords": coords,
            "model": row["EV_Model"],
            "competitor": row["Competitor"],
            "charging_type": row["Charging_Type"],
            "likes": int(row["Likes"]),
        }
    )

json_data = json.dumps(records, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香港電動車充電輿情智慧儀表板 (商業決策版)</title>
    <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css" />
    <script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js"></script>

    <style>
        :root {{
            --color-beige: #dcd8bc;
            --color-red: #85443f;
            --color-green: #367d79;
            --color-black: #58595b;
            --color-blue: #6fa8dc;
        }}
        body {{
            color: var(--color-black);
            background-color: #f7f6f0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        }}
        .bg-beige {{ background-color: var(--color-beige); }}
        .bg-red-custom {{ background-color: var(--color-red); }}
        .bg-green-custom {{ background-color: var(--color-green); }}
        .bg-black-custom {{ background-color: var(--color-black); }}
        .bg-blue-custom {{ background-color: var(--color-blue); }}

        .active-btn-green {{ background-color: var(--color-green) !important; color: white !important; }}
        .active-btn-red {{ background-color: var(--color-red) !important; color: white !important; }}
        .active-btn-black {{ background-color: var(--color-black) !important; color: white !important; }}
        .active-btn-blue {{ background-color: var(--color-blue) !important; color: white !important; }}

        ::-webkit-scrollbar {{ width: 6px; height: 6px; }}
        ::-webkit-scrollbar-track {{ background: #f1f1f1; }}
        ::-webkit-scrollbar-thumb {{ background: #c1c1c1; border-radius: 3px; }}

        .date-picker-input {{
            border: 1px solid #d1d5db;
            padding: 6px 12px;
            border-radius: 8px;
            background-color: #ffffff;
            font-size: 13px;
            outline: none;
        }}
    </style>
</head>
<body class="p-4 md:p-6">

    <header class="mb-6 p-4 rounded-xl shadow-sm text-white flex flex-col md:flex-row justify-between items-start md:items-center bg-[#58595b]">
        <div>
            <h1 class="text-2xl font-bold tracking-wide flex items-center gap-2">
                ⚡ 香港電動車充電商業決策與輿情洞察看板
            </h1>
            <p class="text-xs text-gray-300 mt-1">資深數據分析師重構版：打通車主畫像、競品痛點與區域情緒聯動漏斗</p>
        </div>
        <div class="text-xs text-right mt-2 md:mt-0 bg-black/20 px-3 py-2 rounded-lg">
            當前篩選命中數: <span id="total-count-badge" class="font-bold text-yellow-300">0</span> 筆
        </div>
    </header>

    <!-- 頂層歷史大盤 -->
    <div class="mb-6 grid grid-cols-1 lg:grid-cols-12 gap-4 border border-gray-200 p-4 rounded-xl bg-white shadow-xs">
        <div class="lg:col-span-8 border-r border-gray-100 pr-2">
            <div class="text-xs font-bold text-gray-500 mb-2">📊 歷史全期發文日期分佈趨勢</div>
            <div class="relative w-full h-32">
                <canvas id="staticTrendChart"></canvas>
            </div>
        </div>
        <div class="lg:col-span-4 flex flex-col justify-between">
            <div class="text-xs font-bold text-gray-500 mb-2">🍕 全期輿情情緒基本盤比例佔比</div>
            <div class="relative w-full h-28 flex justify-center items-center">
                <canvas id="staticPieChart"></canvas>
            </div>
        </div>
    </div>

    <!-- 核心多維交互控制區 -->
    <section class="mb-6 p-5 bg-white rounded-xl shadow-xs border border-gray-200">
        <h2 class="text-sm font-semibold text-gray-500 uppercase tracking-wider mb-4 flex items-center gap-1.5">
            🎛️ 多維度智能交叉過濾核心 (點擊圖表或按鈕可多重聯動底端表格)
        </h2>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div class="lg:col-span-5 border-r border-gray-100 pr-2">
                <label class="block text-sm font-bold text-gray-700 mb-2">1. 基本情感傾向 (Emotion)</label>
                <div class="flex gap-2" id="emotion-basic-container">
                    <button onclick="toggleBasicEmotion('全部')" id="eb-all" class="flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-300 transition-all cursor-pointer text-center bg-gray-100 text-gray-700 font-bold active-btn-black">全部</button>
                    <button onclick="toggleBasicEmotion('正面')" id="eb-pos" class="flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-200 transition-all cursor-pointer text-center bg-white text-gray-700 font-bold hover:bg-gray-50">🟢 正面</button>
                    <button onclick="toggleBasicEmotion('中性')" id="eb-neu" class="flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-200 transition-all cursor-pointer text-center bg-white text-gray-700 font-bold hover:bg-gray-50">⚪ 中性</button>
                    <button onclick="toggleBasicEmotion('負面')" id="eb-neg" class="flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-200 transition-all cursor-pointer text-center bg-white text-gray-700 font-bold hover:bg-gray-50">🔴 負面</button>
                </div>
            </div>

            <div class="lg:col-span-4 border-r border-gray-100 pr-2">
                <label class="block text-sm font-bold text-gray-700 mb-2">2. 行銷行銷突破口 (情緒摃桿)</label>
                <div class="flex gap-2">
                    <button onclick="toggleHighLikesFilter()" id="like-filter-btn" class="flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-300 transition-all cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50 flex items-center justify-center gap-1">
                        🔥 只看高讚真實意見 (Likes &gt;= 5)
                    </button>
                </div>
            </div>

            <div class="lg:col-span-3">
                <label class="block text-sm font-bold text-gray-700 mb-2">3. 操作快捷鍵</label>
                <button onclick="resetAllFilters()" class="w-full py-2 px-4 bg-gray-600 hover:bg-gray-700 text-white font-medium text-xs rounded-lg transition-all cursor-pointer flex items-center justify-center gap-1">
                    🔄 重置所有高亮與圖表過濾
                </button>
            </div>
        </div>

        <div class="mt-4 pt-4 border-t border-gray-100">
            <label class="block text-xs font-bold text-gray-500 mb-2">細分語氣情緒 (Emotion Detailed)</label>
            <div id="emotion-detailed-container" class="flex flex-wrap gap-1.5"></div>
        </div>
    </section>

    <!-- 中層主力圖表區：提及主題、地圖怨氣分佈 -->
    <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-6">
        <div class="lg:col-span-5 bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col">
            <div class="flex justify-between items-center mb-4">
                <h3 class="font-bold text-gray-800 text-sm flex items-center gap-2">
                    📋 核心提及主題分佈 (Theme)
                </h3>
                <span class="text-[10px] text-gray-400">點擊橫條可反向鎖定</span>
            </div>
            <div class="relative w-full h-64 flex-1">
                <canvas id="themeChart"></canvas>
            </div>
        </div>

        <div class="lg:col-span-7 bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col h-[360px] lg:h-auto">
            <div class="flex justify-between items-center mb-2">
                <h3 class="font-bold text-gray-800 text-sm flex items-center gap-1">
                    📍 全港充電熱點「體驗怨氣分佈圖」<span class="text-xs font-normal text-gray-400">(綠色: 好評多 | 紅色: 怨氣重)</span>
                </h3>
                <span class="text-[10px] text-gray-400">點擊針頭鎖定地區</span>
            </div>
            <div id="map" class="w-full flex-1 rounded-lg z-10 border border-gray-200"></div>
        </div>
    </div>

    <!-- 關鍵升級交互區：車款與競品聲量 (具備點擊過濾 row data 的功能) -->
    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
        <div class="bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col">
            <div class="flex justify-between items-center mb-2">
                <h3 class="font-bold text-gray-800 text-sm flex items-center gap-2">
                    🚗 提及車款型號聲量 (EV Model) — 點擊柱狀圖鎖定車主評論
                </h3>
                <span id="active-model-badge" class="text-xs bg-blue-100 text-blue-700 font-bold px-2 py-0.5 rounded hide"></span>
            </div>
            <div class="relative w-full h-56">
                <canvas id="modelChart"></canvas>
            </div>
        </div>

        <div class="bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col">
            <div class="flex justify-between items-center mb-2">
                <h3 class="font-bold text-gray-800 text-sm flex items-center gap-2">
                    🥊 提及競爭品牌與對手 (Competitor) — 點擊柱狀圖看競品痛點
                </h3>
                <span id="active-competitor-badge" class="text-xs bg-red-100 text-red-700 font-bold px-2 py-0.5 rounded hide"></span>
            </div>
            <div class="relative w-full h-56">
                <canvas id="competitorChart"></canvas>
            </div>
        </div>
    </div>

    <!-- 底層數據穿透探索器 -->
    <section class="bg-white rounded-xl shadow-xs border border-gray-200 p-5">
        <div class="border-b border-gray-100 pb-3 mb-4">
            <h3 class="text-base font-bold text-gray-800">
                🔍 穿透式原始發文與評論詳情探索器 (同步動態響應上方所有點擊)
            </h3>
        </div>

        <!-- 狀態看板 -->
        <div id="filter-status-bar" class="mb-4 flex flex-wrap gap-2 text-xs font-medium text-gray-600 bg-[#f7f6f0] p-3 rounded-lg border border-gray-200">
            <div class="font-bold text-gray-700">📌 當前篩選沙盒：</div>
            <div id="status-basic" class="bg-white px-2 py-0.5 rounded border border-gray-200">基本情緒: 全部</div>
            <div id="status-detail" class="bg-white px-2 py-0.5 rounded border border-gray-200">細分情緒: 全部</div>
            <div id="status-theme" class="bg-white px-2 py-0.5 rounded border border-gray-200">主題鎖定: 無</div>
            <div id="status-model" class="bg-white px-2 py-0.5 rounded border border-gray-200">車款鎖定: 無</div>
            <div id="status-competitor" class="bg-white px-2 py-0.5 rounded border border-gray-200">競品鎖定: 無</div>
            <div id="status-loc" class="bg-white px-2 py-0.5 rounded border border-gray-200">地點鎖定: 無</div>
            <div id="status-likes" class="bg-white px-2 py-0.5 rounded border border-gray-200">讚數過濾: 全部</div>
            <div id="status-time" class="bg-white px-2 py-0.5 rounded border border-gray-200">時間區間: 全期</div>
        </div>

        <div class="mb-4 p-3 bg-gray-50 rounded-xl border border-gray-200 flex flex-wrap items-center gap-3">
            <span class="text-xs font-bold text-gray-700">📅 歷史時間軸穿透過濾：</span>
            <div class="flex items-center gap-2 text-xs">
                <span>從</span>
                <input type="date" id="calendar-start" class="date-picker-input" onchange="handleDateFilterChange()" />
                <span>至</span>
                <input type="date" id="calendar-end" class="date-picker-input" onchange="handleDateFilterChange()" />
                <button onclick="clearCalendarFilter()" class="bg-gray-200 hover:bg-gray-300 text-gray-700 px-3 py-1.5 rounded-lg font-medium transition-all cursor-pointer">清除日期</button>
            </div>
        </div>

        <div class="overflow-x-auto max-h-[500px] overflow-y-auto border border-gray-200 rounded-lg">
            <table class="w-full text-left border-collapse">
                <thead>
                    <tr class="bg-gray-50 border-b border-gray-200 text-gray-700 text-xs font-bold uppercase tracking-wider sticky top-0 z-20">
                        <th class="p-3 w-16 text-center">編號</th>
                        <th class="p-3 w-32">時間</th>
                        <th class="p-3 w-32">發言來源</th>
                        <th class="p-3 w-40 text-center">車款/競品</th>
                        <th class="p-3 w-28 text-center">情緒傾向</th>
                        <th class="p-3 w-28 text-center">提及地點</th>
                        <th class="p-3">真實留言內文完整摘要</th>
                        <th class="p-3 w-20 text-center">讚數</th>
                    </tr>
                </thead>
                <tbody id="raw-data-table-body" class="text-sm divide-y divide-gray-100"></tbody>
            </table>
        </div>
    </section>

    <script>
        const rawData = {json_data};

        const PALETTE = {{
            beige: '#dcd8bc',
            red: '#85443f',
            green: '#367d79',
            black: '#58595b',
            blue: '#6fa8dc'
        }};

        // 多維交叉篩選狀態
        let currentBasicEmotion = '全部';
        let selectedDetailedEmotions = [];
        let clickedTheme = null;
        let clickedLocation = null;
        let clickedModel = null;        // 🚀 新增車款鎖定
        let clickedCompetitor = null;   // 🚀 新增競品鎖定
        let onlyHighLikes = false;      // 🚀 新增高讚意見過濾

        let filterStartDate = null;
        let filterEndDate = null;

        // 圖表實例
        let themeChartInstance = null;
        let modelChartInstance = null;
        let competitorChartInstance = null;
        let mapInstance = null;
        let markerLayerGroup = null;

        window.addEventListener('DOMContentLoaded', () => {{
            initDetailedEmotionFilters();
            initLeafletMap();
            buildStaticVisuals(); 
            applyCrossFilters();
        }});

        function buildStaticVisuals() {{
            const months = rawData.map(d => {{
                let t = d.time ? d.time.trim() : '';
                return (t.length >= 7) ? t.substring(0, 7) : null;
            }}).filter(m => m && m.match(/^\d{{4}}-\d{{2}}$/)).sort();

            if (months.length === 0) return;

            const uniqueMonths = [...new Set(months)];
            const monthCounts = {{}};
            months.forEach(m => {{ monthCounts[m] = (monthCounts[m] || 0) + 1; }});

            const trendCtx = document.getElementById('staticTrendChart').getContext('2d');
            new Chart(trendCtx, {{
                type: 'bar',
                data: {{
                    labels: uniqueMonths,
                    datasets: [{{
                        label: '發文數量趨勢',
                        data: uniqueMonths.map(m => monthCounts[m]),
                        backgroundColor: '#dcd8bc95',
                        borderColor: '#58595b',
                        borderWidth: 1,
                        maxBarThickness: 20
                    }}]
                }},
                options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ display: false }} }} }}
            }});

            let posCount = 0, neuCount = 0, negCount = 0;
            rawData.forEach(d => {{
                if (d.emotion_basic === '正面') posCount++;
                else if (d.emotion_basic === '負面') negCount++;
                else neuCount++;
            }});

            const pieCtx = document.getElementById('staticPieChart').getContext('2d');
            new Chart(pieCtx, {{
                type: 'pie',
                data: {{
                    labels: ['正面', '中性', '負面'],
                    datasets: [{{
                        data: [posCount, neuCount, negCount],
                        backgroundColor: [PALETTE.green, '#e5e7eb', PALETTE.red]
                    }}]
                }},
                options: {{ responsive: true, maintainAspectRatio: false, plugins: {{ legend: {{ position: 'right' }} }} }}
            }});
        }}

        function initDetailedEmotionFilters() {{
            const details = [...new Set(rawData.map(d => d.emotion_detailed))].filter(Boolean);
            const container = document.getElementById('emotion-detailed-container');
            container.innerHTML = '';
            details.forEach(det => {{
                const btn = document.createElement('button');
                btn.className = "py-1 px-2.5 rounded-full text-xs font-medium border border-gray-300 bg-white hover:bg-gray-50 cursor-pointer transition-all";
                btn.id = `ed-${{det}}`;
                btn.innerText = det;
                btn.onclick = () => toggleDetailedEmotion(det);
                container.appendChild(btn);
            }});
        }}

        function initLeafletMap() {{
            mapInstance = L.map('map').setView([22.3527, 114.1272], 11);
            L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; OpenStreetMap'
            }}).addTo(mapInstance);
            markerLayerGroup = L.layerGroup().addTo(mapInstance);
        }}

        function toggleBasicEmotion(type) {{
            currentBasicEmotion = type;
            ['all', 'pos', 'neu', 'neg'].forEach(id => {{
                document.getElementById(`eb-${{id}}`).className = "flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-200 cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50";
            }});
            if(type === '全部') document.getElementById('eb-all').classList.add('active-btn-black');
            if(type === '正面') document.getElementById('eb-pos').classList.add('active-btn-green');
            if(type === '中性') document.getElementById('eb-neu').classList.add('active-btn-black');
            if(type === '負面') document.getElementById('eb-neg').classList.add('active-btn-red');
            applyCrossFilters();
        }}

        function toggleDetailedEmotion(det) {{
            const idx = selectedDetailedEmotions.indexOf(det);
            const btn = document.getElementById(`ed-${{det}}`);
            if(idx > -1) {{
                selectedDetailedEmotions.splice(idx, 1);
                btn.className = "py-1 px-2.5 rounded-full text-xs font-medium border border-gray-300 bg-white hover:bg-gray-50 cursor-pointer transition-all";
            }} else {{
                selectedDetailedEmotions.push(det);
                btn.className = "py-1 px-2.5 rounded-full text-xs font-medium text-white border bg-[#6fa8dc] border-[#6fa8dc] cursor-pointer transition-all";
            }}
            applyCrossFilters();
        }}

        // 🚀 新增：切換高讚意見漏斗
        function toggleHighLikesFilter() {{
            onlyHighLikes = !onlyHighLikes;
            const btn = document.getElementById('like-filter-btn');
            if(onlyHighLikes) {{
                btn.className = "flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-amber-500 transition-all cursor-pointer text-center text-white bg-amber-500 flex items-center justify-center gap-1 font-bold shadow-xs";
            }} else {{
                btn.className = "flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-300 transition-all cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50 flex items-center justify-center gap-1";
            }}
            applyCrossFilters();
        }}

        function handleDateFilterChange() {{
            filterStartDate = document.getElementById('calendar-start').value || null;
            filterEndDate = document.getElementById('calendar-end').value || null;
            applyCrossFilters();
        }}

        function clearCalendarFilter() {{
            filterStartDate = null; filterEndDate = null;
            document.getElementById('calendar-start').value = '';
            document.getElementById('calendar-end').value = '';
            applyCrossFilters();
        }}

        function resetAllFilters() {{
            clickedTheme = null; clickedLocation = null;
            clickedModel = null; clickedCompetitor = null; onlyHighLikes = false;
            currentBasicEmotion = '全部'; selectedDetailedEmotions = [];
            filterStartDate = null; filterEndDate = null;

            document.getElementById('calendar-start').value = '';
            document.getElementById('calendar-end').value = '';

            const lbtn = document.getElementById('like-filter-btn');
            lbtn.className = "flex-1 py-2 px-3 rounded-lg font-medium text-xs border border-gray-300 transition-all cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50 flex items-center justify-center gap-1";

            toggleBasicEmotion('全部');
            initDetailedEmotionFilters();
        }}

        // 核心交叉過濾器漏斗
        function applyCrossFilters() {{
            let filtered = rawData;

            if (currentBasicEmotion !== '全部') filtered = filtered.filter(d => d.emotion_basic === currentBasicEmotion);
            if (selectedDetailedEmotions.length > 0) filtered = filtered.filter(d => selectedDetailedEmotions.includes(d.emotion_detailed));
            if (clickedTheme) filtered = filtered.filter(d => d.theme === clickedTheme);
            if (clickedLocation) filtered = filtered.filter(d => d.location === clickedLocation);
            if (clickedModel) filtered = filtered.filter(d => d.model && d.model.includes(clickedModel));
            if (clickedCompetitor) filtered = filtered.filter(d => d.competitor && d.competitor.includes(clickedCompetitor));
            if (onlyHighLikes) filtered = filtered.filter(d => d.likes >= 5);
            if (filterStartDate) filtered = filtered.filter(d => d.time && d.time >= filterStartDate);
            if (filterEndDate) filtered = filtered.filter(d => d.time && d.time <= filterEndDate);

            // 更新 UI 狀態資訊欄
            document.getElementById('total-count-badge').innerText = filtered.length;
            document.getElementById('status-basic').innerText = `基本情緒: ${{currentBasicEmotion}}`;
            document.getElementById('status-detail').innerText = `細分情緒: ${{selectedDetailedEmotions.length ? selectedDetailedEmotions.join(',') : '全部'}}`;
            document.getElementById('status-theme').innerText = `主題鎖定: ${{clickedTheme || '無'}}`;
            document.getElementById('status-model').innerText = `車款鎖定: ${{clickedModel || '無'}}`;
            document.getElementById('status-competitor').innerText = `競品鎖定: ${{clickedCompetitor || '無'}}`;
            document.getElementById('status-loc').innerText = `地點鎖定: ${{clickedLocation || '無'}}`;
            document.getElementById('status-likes').innerText = `讚數過濾: ${{onlyHighLikes ? '🔥 Likes>=5 高讚意見' : '全部'}}`;
            document.getElementById('status-time').innerText = `時間區間: ${{filterStartDate || '全期'}} ~ ${{filterEndDate || '全期'}}`;

            // 重新繪製受聯動影響的組件
            renderThemeChart(filtered);
            renderModelChart(filtered);
            renderCompetitorChart(filtered);
            renderMapMarkers(filtered);
            renderDataTable(filtered);
        }}

        function getCounts(array, key) {{
            const counts = {{}};
            array.forEach(item => {{
                let val = item[key];
                if (!val || val === 'null' || val === '未提及') return;
                let parts = val.split(',').map(p => p.trim());
                parts.forEach(p => {{ if(p && p !== 'null') counts[p] = (counts[p] || 0) + 1; }});
            }});
            return Object.entries(counts).sort((a,b) => b[1] - a[1]);
        }}

        function renderThemeChart(data) {{
            const sortedData = getCounts(data, 'theme');
            const labels = sortedData.map(d => d[0]);
            const counts = sortedData.map(d => Number(d[1] || 0));

            if (themeChartInstance) themeChartInstance.destroy();
            const ctx = document.getElementById('themeChart').getContext('2d');
            themeChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: '提及次數',
                        data: counts,
                        backgroundColor: clickedTheme ? PALETTE.black : PALETTE.blue,
                        borderRadius: 4
                    }}]
                }},
                options: {{
                    indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    onClick: (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            clickedTheme = (clickedTheme === labels[idx]) ? null : labels[idx];
                            applyCrossFilters();
                        }}
                    }}
                }}
            }});
        }}

        // 🚀 關鍵升級 1：車款長條圖支持交互聯動過濾器
        function renderModelChart(data) {{
            const sortedData = getCounts(data, 'model').slice(0, 8);
            const labels = sortedData.map(d => d[0]);
            const counts = sortedData.map(d => d[1]);

            const badge = document.getElementById('active-model-badge');
            if(clickedModel) {{
                badge.innerText = `已鎖定: ${{clickedModel}}`;
                badge.classList.remove('hide');
            }} else {{ badge.classList.add('hide'); }}

            if (modelChartInstance) modelChartInstance.destroy();
            const ctx = document.getElementById('modelChart').getContext('2d');
            modelChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: '提及數量',
                        data: counts,
                        // 如果有被點擊高亮，則該柱子變深色
                        backgroundColor: labels.map(l => l === clickedModel ? PALETTE.black : PALETTE.beige),
                        borderColor: PALETTE.black,
                        borderWidth: 1,
                        borderRadius: 4
                    }}]
                }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    onClick: (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            const targetLabel = labels[idx];
                            // 反覆點擊同一根柱子則取消鎖定
                            clickedModel = (clickedModel === targetLabel) ? null : targetLabel;
                            applyCrossFilters();
                        }}
                    }}
                }}
            }});
        }}

        // 🚀 關鍵升級 2：競品長條圖支持交互聯動過濾器
        function renderCompetitorChart(data) {{
            const sortedData = getCounts(data, 'competitor').slice(0, 8);
            const labels = sortedData.map(d => d[0]);
            const counts = sortedData.map(d => d[1]);

            const badge = document.getElementById('active-competitor-badge');
            if(clickedCompetitor) {{
                badge.innerText = `已鎖定: ${{clickedCompetitor}}`;
                badge.classList.remove('hide');
            }} else {{ badge.classList.add('hide'); }}

            if (competitorChartInstance) competitorChartInstance.destroy();
            const ctx = document.getElementById('competitorChart').getContext('2d');
            competitorChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: '提及數量',
                        data: counts,
                        backgroundColor: labels.map(l => l === clickedCompetitor ? PALETTE.black : '#edf2f7'),
                        borderColor: PALETTE.red,
                        borderWidth: 1,
                        borderRadius: 4
                    }}]
                }},
                options: {{
                    responsive: true, maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    onClick: (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            const targetLabel = labels[idx];
                            clickedCompetitor = (clickedCompetitor === targetLabel) ? null : targetLabel;
                            applyCrossFilters();
                        }}
                    }}
                }}
            }});
        }}

        // 🚀 關鍵升級 3：地圖改為「體驗怨氣分佈圖」（動態計算該地點正負面比例並變色）
        function renderMapMarkers(data) {{
            markerLayerGroup.clearLayers();

            const locStats = {{}};
            data.forEach(d => {{
                if(d.coords && d.location && d.location !== '未提及' && d.location !== '香港') {{
                    if(!locStats[d.location]) {{
                        locStats[d.location] = {{ coords: d.coords, total: 0, pos: 0, neg: 0 }};
                    }}
                    locStats[d.location].total++;
                    if(d.emotion_basic === '正面') locStats[d.location].pos++;
                    if(d.emotion_basic === '負面') locStats[d.location].neg++;
                }}
            }});

            Object.entries(locStats).forEach(([name, stat]) => {{
                // 計算正面率和負面率，決定大頭針顏色
                let pinColor = PALETTE.blue; // 預設藍色 (中性居多)
                if (stat.pos > stat.neg) pinColor = PALETTE.green; // 好評居多變綠色
                if (stat.neg > stat.pos) pinColor = PALETTE.red;   // 怨氣重變紅色

                // 如果當前地點被選中，加一個黑色邊框強調
                const borderStyle = (clickedLocation === name) ? "border: 3px solid #000; scale: 1.1;" : "border: 2px solid white;";

                const customIcon = L.divIcon({{
                    html: `<div style="background-color: ${{pinColor}}; ${{borderStyle}}" class="w-7 h-7 rounded-full flex items-center justify-center text-white text-xs font-bold shadow-md">${{stat.total}}</div>`,
                    className: '',
                    iconSize: [28, 28]
                }});

                const marker = L.marker(stat.coords, {{ icon: customIcon }}).addTo(markerLayerGroup);

                const posRate = Math.round((stat.pos / stat.total) * 100);
                const negRate = Math.round((stat.neg / stat.total) * 100);

                marker.bindPopup(`
                    <div class="text-xs">
                        <b>📍 ${{name}}</b><br>
                        總提及數: <b>${{stat.total}}</b> 次<br>
                        🟢 好評率: <b>${{posRate}}%</b> | 🔴 怨氣率: <b>${{negRate}}%</b><br>
                        <span style="color:#888;">(點擊大頭針鎖定此地評論)</span>
                    </div>
                `);

                marker.on('click', () => {{
                    clickedLocation = (clickedLocation === name) ? null : name;
                    applyCrossFilters();
                }});
            }});
        }}

        function renderDataTable(data) {{
            const tbody = document.getElementById('raw-data-table-body');
            tbody.innerHTML = '';

            if(data.length === 0) {{
                tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-gray-400 font-medium">沒有符合當前多維動態沙盒條件的發文數據。</td></tr>`;
                return;
            }}

            data.forEach(row => {{
                let emotionBadge = '';
                if(row.emotion_basic === '正面') emotionBadge = `<span class="bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded border border-emerald-200 text-[11px] font-semibold">🟢 ${{row.emotion_detailed}}</span>`;
                else if(row.emotion_basic === '負面') emotionBadge = `<span class="bg-rose-50 text-rose-700 px-2 py-0.5 rounded border border-rose-200 text-[11px] font-semibold">🔴 ${{row.emotion_detailed}}</span>`;
                else emotionBadge = `<span class="bg-gray-100 text-gray-700 px-2 py-0.5 rounded border border-gray-200 text-[11px] font-semibold">⚪ ${{row.emotion_detailed}}</span>`;

                // 把車款和競品合併顯示，方便一眼對照用戶畫像
                const modelStr = row.model !== 'null' ? row.model : '-';
                const compStr = row.competitor !== 'null' ? row.competitor : '-';

                const tr = document.createElement('tr');
                tr.className = "hover:bg-gray-50/80 transition-colors";

                // 高讚突出顯示（加火苗圖標）
                const likesDisplay = row.likes >= 5 ? `🔥 <span class="text-amber-600 font-bold">${{row.likes}}</span>` : row.likes;

                tr.innerHTML = `
                    <td class="p-3 text-center text-gray-500 font-mono text-xs">${{row.id}}</td>
                    <td class="p-3 text-xs text-gray-600 font-mono">${{row.time}}</td>
                    <td class="p-3 text-xs text-gray-500 truncate" style="max-width:120px;" title="${{row.source}}">${{row.source}}</td>
                    <td class="p-3 text-xs text-center">
                        <div class="font-semibold text-blue-600">🚗 ${{modelStr}}</div>
                        <div class="text-[10px] text-gray-400">🥊 競品: ${{compStr}}</div>
                    </td>
                    <td class="p-3 text-center">${{emotionBadge}}</td>
                    <td class="p-3 text-center font-medium text-gray-600 text-xs">${{row.location !== 'null' && row.location !== '未提及' ? '📍 ' + row.location : '<span class="text-gray-300">-</span>'}}</td>
                    <td class="p-3 text-gray-700 leading-relaxed text-xs max-w-md break-words font-sans">${{row.text}}</td>
                    <td class="p-3 text-center font-mono">${{likesDisplay}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}
    </script>
</body>
</html>
"""

# 將整合後的程式碼寫入本地檔案
output_filename = "ev_dashboard_comprehensive.html"
with open(output_filename, "w", encoding="utf-8") as f:
    f.write(html_content)

print("\n" + "=" * 50)
print(f"🎉 商業決策版 Dashboard 構建成功！")
print(f"1. 打通了 EV Model 與 Competitor 柱狀圖的『反向穿透點擊』功能，點擊直接過濾原始留言。")
print(f"2. 地圖進化為『體驗怨氣分佈圖』，自動計算區域正負面情緒比率並以紅、綠、藍變色呈現。")
print(f"3. 引入行銷專用『🔥 只看高讚真實意見』篩選鍵，秒級提煉社群爆點與高價值痛點。")
print(f"請在瀏覽器中打開: {output_filename} 查看全新交互效果！")
print("=" * 50 + "\n")