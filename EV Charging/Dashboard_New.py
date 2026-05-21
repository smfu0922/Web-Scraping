# python
import pandas as pd
import json
import os

# 1. 讀取清洗後的 EV 數據
file_path = r'C:\Users\User\Downloads\ev_scraping_full.xlsx'
sheet = 'Cleaned EV Data'

if not os.path.exists(file_path):
    print(f"錯誤：找不到檔案 {file_path}，請確保檔案在同一個目錄下！")
    exit()

df = pd.read_excel(file_path, sheet_name=sheet)

# 2. 數據過濾：完全排除廣告與促銷的內容
initial_count = len(df)
df = df[~df['Purpose'].str.contains('廣告|促銷', na=False, case=False)]
df = df[~df['Theme'].str.contains('廣告|促銷', na=False, case=False)]
print(f"數據過濾完成：已排除廣告。從 {initial_count} 筆過濾至 {len(df)} 筆有效評論。")

# 填補空值
df['Emotion'] = df['Emotion'].fillna('中性').str.strip()
df['Emotion_Detail'] = df['Emotion_Detail'].fillna('陳述').str.strip()
df['Theme'] = df['Theme'].fillna('其他').str.strip()
df['Location'] = df['Location'].fillna('null').str.strip()
df['EV_Model'] = df['EV_Model'].fillna('null').str.strip()
df['Competitor'] = df['Competitor'].fillna('null').str.strip()
df['Charging_Type'] = df['Charging_Type'].fillna('null').str.strip()
df['Text'] = df['Text'].fillna('').str.strip()
df['User'] = df['User'].fillna('匿名用戶').str.strip()
df['Likes'] = pd.to_numeric(df['Likes'], errors='coerce').fillna(0).astype(int)


# 🕒 強大防呆：將 Excel 內各種混亂、整數天數或日期欄位，通通強制清洗為完美的 YYYY-MM-DD
def clean_to_date_str(val):
    if pd.isna(val):
        return '未提及'
    if isinstance(val, pd.Timestamp):
        return val.strftime('%Y-%m-%d')

    val_str = str(val).strip()
    if val_str.lower() in ['nan', 'nat', 'none', '']:
        return '未提及'

    # 如果是 Excel 拋出來的 5 位數整數天數 (例如 45284)
    if val_str.isdigit() or (val_str.replace('.', '', 1).isdigit() and float(val_str).is_integer()):
        try:
            dt = pd.to_datetime(int(float(val_str)), unit='D', origin='1899-12-30')
            return dt.strftime('%Y-%m-%d')
        except:
            pass

    # 常規 YYYY-MM-DD 或是帶有時間戳格式的通用解析
    try:
        dt = pd.to_datetime(val_str, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime('%Y-%m-%d')
    except:
        pass

    return val_str[:10] if len(val_str) >= 10 and val_str[:4].isdigit() else val_str


df['Formatted_Time'] = df['Time'].apply(clean_to_date_str)

# 3. 建立香港地區座標對照字典
geo_dict = {
    '荃灣': [22.3686, 114.1114], '荃灣京匯中心': [22.3703, 114.1122], '海之戀': [22.3670, 114.1110],
    '荃灣馬角街18號': [22.3655, 114.1171],
    '九龍灣': [22.3210, 114.2100], '九龍灣大昌行': [22.3242, 114.2091], '九龍灣大昌': [22.3242, 114.2091],
    '將軍澳': [22.3120, 114.2560], '將軍澳尚德邨': [22.3135, 114.2601], '觀塘': [22.3130, 114.2250],
    '鴻圖道6號': [22.3151, 114.2215], '觀塘鴻圖道6號': [22.3151, 114.2215], '偉業街': [22.3180, 114.2190],
    '大角咀': [22.3220, 114.1620], '大角咀 One Bedford Place': [22.3236, 114.1610], '奧海城': [22.3175, 114.1602],
    '佐敦': [22.3050, 114.1710], '佐敦高臨停車場': [22.3045, 114.1705], '佐敦西貢街': [22.3061, 114.1708],
    '尖沙咀': [22.2988, 114.1722], '紅磡': [22.3075, 114.1831], '深水埗': [22.3300, 114.1620],
    '元朗': [22.4450, 114.0220], '元朗新潭路': [22.4720, 114.0530], '錦上路站': [22.4340, 114.0620],
    '錦上路': [22.4340, 114.0620], '大棠': [22.4170, 114.0220],
    '屯門': [22.3950, 113.9740], '火炭': [22.3980, 114.1950], '沙田': [22.3830, 114.1830],
    '馬鞍山': [22.4230, 114.2300], '馬鞍山頌安停車場': [22.4255, 114.2285],
    '大埔': [22.4500, 114.1650], '大埔東昌街': [22.4468, 114.1691], '大埔東昌街體育館停車場': [22.4468, 114.1691],
    '粉嶺': [22.4920, 114.1380], '聯和墟': [22.4990, 114.1420], '香園圍': [22.5530, 114.1560],
    '灣仔': [22.2790, 114.1720], '灣仔尚翹峰灣仔街市': [22.2754, 114.1742], '金鐘': [22.2780, 114.1650],
    '中環': [22.2820, 114.1580],
    '鰂魚涌': [22.2840, 114.2120], 'K11 ATELIER King’s Road Carpark': [22.2851, 114.2110],
    '筲箕灣': [22.2780, 114.2290], '筲箕灣天悅廣場': [22.2788, 114.2278],
    '數碼港': [22.2610, 114.1300], '海洋公園': [22.2467, 114.1757], '香港科技大學': [22.3364, 114.2654],
    '機場': [22.3080, 113.9185],
    '西貢': [22.3830, 114.2700], '龍蝦灣': [22.3160, 114.2900]
}

HK_CENTER = [22.3527, 114.1272]

records = []
for idx, row in df.iterrows():
    loc_name = row['Location']
    coords = geo_dict.get(loc_name, HK_CENTER if loc_name != '未提及' and loc_name != '香港' else None)

    records.append({
        "id": int(row['ID']),
        "source": row['Source'],
        "time": row['Formatted_Time'],  # 完美的 YYYY-MM-DD
        "text": row['Text'],
        "emotion_basic": row['Emotion'],
        "emotion_detailed": row['Emotion_Detail'],
        "theme": row['Theme'],
        "location": loc_name,
        "coords": coords,
        "model": row['EV_Model'],
        "competitor": row['Competitor'],
        "charging_type": row['Charging_Type'],
        "likes": int(row['Likes'])
    })

json_data = json.dumps(records, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="zh-HK">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>香港電動車充電輿情智慧儀表板</title>
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
        .border-beige {{ border-color: var(--color-beige); }}
        .text-black-custom {{ color: var(--color-black); }}

        .active-btn-green {{ background-color: var(--color-green) !important; color: white !important; }}
        .active-btn-red {{ background-color: var(--color-red) !important; color: white !important; }}
        .active-btn-black {{ background-color: var(--color-black) !important; color: white !important; }}

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
                ⚡ 香港電動車充電社群輿情觀測儀表板
            </h1>
            <p class="text-xs text-gray-300 mt-1">深度社群網絡聆聽與數據穿透分析系統 (已排除廣告宣傳干擾)</p>
        </div>
        <div class="text-xs text-right mt-2 md:mt-0 bg-black/20 px-3 py-2 rounded-lg">
            數據總量: <span id="total-count-badge" class="font-bold text-yellow-300">0</span> 筆篩選中
        </div>
    </header>

    <div class="mb-6 grid grid-cols-1 lg:grid-cols-12 gap-4 border border-gray-200 p-4 rounded-xl bg-white shadow-xs">
        <div class="lg:col-span-8 border-r border-gray-100 pr-2">
            <div class="text-xs font-bold text-gray-500 mb-2 flex justify-between items-center">
                <span>📊 歷史全期發文日期分佈趨勢 (純視覺對照)</span>
                <span class="text-[10px] bg-gray-100 px-1.5 py-0.5 rounded text-gray-400 font-normal">不引入交互</span>
            </div>
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

    <section class="mb-6 p-5 bg-white rounded-xl shadow-xs border border-gray-200">
        <h2 class="text-sm font-semibold text-gray-500 uppercase tracking-wider mb-4 flex items-center gap-1.5">
            🎛️ 多維度智能交叉過濾核心
        </h2>
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
            <div class="lg:col-span-5 border-r border-gray-100 pr-2">
                <label class="block text-sm font-medium text-gray-700 mb-2 font-bold">基本傾向傾向 (Emotion Basic)</label>
                <div class="flex gap-2" id="emotion-basic-container">
                    <button onclick="toggleBasicEmotion('全部')" id="eb-all" class="flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-300 transition-all cursor-pointer text-center bg-gray-100 text-gray-700 font-bold active-btn-black">全部</button>
                    <button onclick="toggleBasicEmotion('正面')" id="eb-pos" class="flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-200 transition-all cursor-pointer text-center bg-white text-gray-700 font-bold hover:bg-gray-50">🟢 正面</button>
                    <button onclick="toggleBasicEmotion('中性')" id="eb-neu" class="flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-200 transition-all cursor-pointer text-center bg-white text-gray-700 font-bold hover:bg-gray-50">⚪ 中性</button>
                    <button onclick="toggleBasicEmotion('負面')" id="eb-neg" class="flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-200 transition-all cursor-pointer text-center bg-white text-gray-700 font-bold hover:bg-gray-50">🔴 負面</button>
                </div>
            </div>

            <div class="lg:col-span-7">
                <label class="block text-sm font-medium text-gray-700 mb-2 font-bold">細分語氣情緒 (Emotion Detailed) <span class="text-xs font-normal text-gray-400">(點擊可多選或複選組合)</span></label>
                <div id="emotion-detailed-container" class="flex flex-wrap gap-2"></div>
            </div>
        </div>
    </section>

    <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 mb-6">
        <div class="lg:col-span-5 flex flex-col gap-6">
            <div class="bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex-1 flex flex-col">
                <div class="flex justify-between items-center mb-4">
                    <h3 class="font-bold text-gray-800 flex items-center gap-2">
                        📋 核心提及主題分佈 (Theme)
                        <span id="theme-count-badge" class="ml-2 inline-flex items-center bg-[#6fa8dc] text-white text-xs font-semibold px-2 py-0.5 rounded-full">0</span>
                    </h3>
                    <span class="text-xs text-gray-400">點擊柱狀圖可反向鎖定發文</span>
                </div>
                <div class="relative w-full h-64 flex-1">
                    <canvas id="themeChart"></canvas>
                </div>
            </div>
        </div>

        <div class="lg:col-span-7 bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col h-[360px] lg:h-auto">
            <div class="flex justify-between items-center mb-2">
                <h3 class="font-bold text-gray-800 flex items-center gap-2">
                    📍 香港 EV 充電網民提及熱點分布圖
                </h3>
                <span class="text-xs text-gray-400">點擊地圖大頭針看該區評論</span>
            </div>
            <div id="map" class="w-full flex-1 rounded-lg z-10 border border-gray-200"></div>
        </div>
    </div>

    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
        <div class="bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col">
            <h3 class="font-bold text-gray-800 mb-4 flex items-center gap-2">
                🚗 提及車款型號聲量 (EV Model)
            </h3>
            <div class="relative w-full h-56">
                <canvas id="modelChart"></canvas>
            </div>
        </div>

        <div class="bg-white p-5 rounded-xl shadow-xs border border-gray-200 flex flex-col">
            <h3 class="font-bold text-gray-800 mb-4 flex items-center gap-2">
                🥊 提及競爭品牌與對手 (Competitor)
            </h3>
            <div class="relative w-full h-56">
                <canvas id="competitorChart"></canvas>
            </div>
        </div>
    </div>

    <section class="bg-white rounded-xl shadow-xs border border-gray-200 p-5">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center border-b border-gray-100 pb-4 mb-4 gap-4">
            <div>
                <h3 class="text-lg font-bold text-gray-800 flex items-center gap-2">
                    🔍 穿透式原始發文與評論詳情探索器
                </h3>
                <p class="text-xs text-gray-400 mt-0.5">點擊上方圖表、地圖或按鈕，這裡會精確同步穿透並顯示對應的真實留言</p>
            </div>
            <button onclick="resetAllFilters()" class="text-xs bg-gray-600 hover:bg-gray-700 text-white font-medium px-4 py-2 rounded-lg transition-all cursor-pointer flex items-center gap-1">
                🔄 重置所有高亮鎖定
            </button>
        </div>

        <div id="filter-status-bar" class="mb-4 flex flex-wrap gap-2 text-xs font-medium text-gray-600 bg-[#f7f6f0] p-3 rounded-lg border border-gray-200">
            <div>當前鎖定：</div>
            <div id="status-basic" class="bg-gray-200 px-2 py-0.5 rounded">基本情緒: 全部</div>
            <div id="status-detail" class="bg-gray-200 px-2 py-0.5 rounded">細分情緒: 全部</div>
            <div id="status-theme" class="bg-gray-200 px-2 py-0.5 rounded">主題鎖定: 無</div>
            <div id="status-loc" class="bg-gray-200 px-2 py-0.5 rounded">地點鎖定: 無</div>
            <div id="status-time" class="bg-gray-200 px-2 py-0.5 rounded">時間區間: 全部</div>
        </div>

        <div class="mb-4 p-3 bg-gray-50 rounded-xl border border-gray-200 flex flex-wrap items-center gap-3">
            <span class="text-xs font-bold text-gray-700">📅 歷史時間軸穿透過濾 (YYYY-MM-DD)：</span>
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
                        <th class="p-3 w-40">時間 (Time)</th>
                        <th class="p-3 w-32">發言社群來源</th>
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

        let currentBasicEmotion = '全部';
        let selectedDetailedEmotions = [];
        let clickedTheme = null;
        let clickedLocation = null;

        let filterStartDate = null;
        let filterEndDate = null;

        let themeChartInstance = null;
        let modelChartInstance = null;
        let competitorChartInstance = null;
        let mapInstance = null;
        let markerLayerGroup = null;

        window.addEventListener('DOMContentLoaded', () => {{
            initDetailedEmotionFilters();
            initLeafletMap();
            buildStaticVisuals(); // 在此加載最頂端的歷史大盤
            applyCrossFilters();
        }});

        // 🎨 構建最頂層歷史宏觀看板 (70% 歷史時間趨勢 + 30% 情緒比例，無視覺交互)
        function buildStaticVisuals() {{
            const months = rawData.map(d => {{
                let t = d.time ? d.time.trim() : '';
                return (t.length >= 7) ? t.substring(0, 7) : null;
            }}).filter(m => m && m.match(/^\d{{4}}-\d{{2}}$/)).sort();

            if (months.length === 0) return;

            const minMonth = months[0];
            const maxMonth = months[months.length - 1];

            const monthCounts = {{}};
            months.forEach(m => {{ monthCounts[m] = (monthCounts[m] || 0) + 1; }});

            const uniqueMonths = [...new Set(months)];
            const chartLabels = uniqueMonths;
            const chartData = uniqueMonths.map(m => monthCounts[m]);

            // 1. 繪製 70% 歷史月度分佈
            const trendCtx = document.getElementById('staticTrendChart').getContext('2d');
            new Chart(trendCtx, {{
                type: 'bar',
                data: {{
                    labels: chartLabels,
                    datasets: [{{
                        label: `發文總量 (${{minMonth}} 至 ${{maxMonth}})`,
                        data: chartData,
                        backgroundColor: '#dcd8bc95',
                        borderColor: '#58595b',
                        borderWidth: 1,
                        borderRadius: 3,
                        maxBarThickness: 20
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: true, labels: {{ boxWidth: 10, font: {{ size: 10 }} }} }} }},
                    scales: {{
                        x: {{ grid: {{ display: false }}, ticks: {{ font: {{ size: 9 }} }} }},
                        y: {{ grid: {{ color: '#f3f4f6' }}, ticks: {{ font: {{ size: 9 }}, precision: 0 }} }}
                    }}
                }}
            }});

            // 2. 統計全期基本傾向數據
            let posCount = 0, neuCount = 0, negCount = 0;
            rawData.forEach(d => {{
                if (d.emotion_basic === '正面') posCount++;
                else if (d.emotion_basic === '負面') negCount++;
                else neuCount++;
            }});

            // 3. 繪製 30% 基本情感大盤圓餅圖
            const pieCtx = document.getElementById('staticPieChart').getContext('2d');
            new Chart(pieCtx, {{
                type: 'pie',
                data: {{
                    labels: ['正面', '中性', '負面'],
                    datasets: [{{
                        data: [posCount, neuCount, negCount],
                        backgroundColor: [PALETTE.green, '#e5e7eb', PALETTE.red],
                        borderWidth: 1
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{ position: 'right', labels: {{ boxWidth: 8, font: {{ size: 9 }} }} }}
                    }}
                }}
            }});
        }}

        function initDetailedEmotionFilters() {{
            const details = [...new Set(rawData.map(d => d.emotion_detailed))].filter(Boolean);
            const container = document.getElementById('emotion-detailed-container');
            container.innerHTML = '';

            details.forEach(det => {{
                const btn = document.createElement('button');
                btn.className = "py-1.5 px-3 rounded-full text-xs font-medium border border-gray-300 bg-white hover:bg-gray-50 cursor-pointer transition-all";
                btn.id = `ed-${{det}}`;
                btn.innerText = det;
                btn.onclick = () => toggleDetailedEmotion(det);
                container.appendChild(btn);
            }});
        }}

        function initLeafletMap() {{
            mapInstance = L.map('map').setView([22.3527, 114.1272], 11);
            L.tileLayer('https://{{s}}.basemaps.cartocdn.com/light_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; OpenStreetMap &copy; CARTO'
            }}).addTo(mapInstance);

            markerLayerGroup = L.layerGroup().addTo(mapInstance);
        }}

        function toggleBasicEmotion(type) {{
            currentBasicEmotion = type;

            document.getElementById('eb-all').className = "flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-300 cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50";
            document.getElementById('eb-pos').className = "flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-200 cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50";
            document.getElementById('eb-neu').className = "flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-200 cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50";
            document.getElementById('eb-neg').className = "flex-1 py-2.5 px-4 rounded-lg font-medium text-sm border border-gray-200 cursor-pointer text-center bg-white text-gray-700 hover:bg-gray-50";

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
                btn.className = "py-1.5 px-3 rounded-full text-xs font-medium border border-gray-300 bg-white hover:bg-gray-50 cursor-pointer transition-all";
            }} else {{
                selectedDetailedEmotions.push(det);
                btn.className = "py-1.5 px-3 rounded-full text-xs font-medium text-white border bg-[#6fa8dc] border-[#6fa8dc] cursor-pointer transition-all";
            }}
            applyCrossFilters();
        }}

        function handleDateFilterChange() {{
            filterStartDate = document.getElementById('calendar-start').value || null;
            filterEndDate = document.getElementById('calendar-end').value || null;
            applyCrossFilters();
        }}

        function clearCalendarFilter() {{
            filterStartDate = null;
            filterEndDate = null;
            document.getElementById('calendar-start').value = '';
            document.getElementById('calendar-end').value = '';
            applyCrossFilters();
        }}

        function resetAllFilters() {{
            clickedTheme = null;
            clickedLocation = null;
            currentBasicEmotion = '全部';
            selectedDetailedEmotions = [];
            filterStartDate = null;
            filterEndDate = null;

            document.getElementById('calendar-start').value = '';
            document.getElementById('calendar-end').value = '';

            toggleBasicEmotion('全部');
            initDetailedEmotionFilters();
        }}

        function applyCrossFilters() {{
            let filtered = rawData;

            if (currentBasicEmotion !== '全部') {{
                filtered = filtered.filter(d => d.emotion_basic === currentBasicEmotion);
            }}
            if (selectedDetailedEmotions.length > 0) {{
                filtered = filtered.filter(d => selectedDetailedEmotions.includes(d.emotion_detailed));
            }}
            if (clickedTheme) {{
                filtered = filtered.filter(d => d.theme === clickedTheme);
            }}
            if (clickedLocation) {{
                filtered = filtered.filter(d => d.location === clickedLocation);
            }}

            if (filterStartDate) {{
                filtered = filtered.filter(d => d.time && d.time >= filterStartDate);
            }}
            if (filterEndDate) {{
                filtered = filtered.filter(d => d.time && d.time <= filterEndDate);
            }}

            document.getElementById('total-count-badge').innerText = filtered.length;
            document.getElementById('status-basic').innerText = `基本情緒: ${{currentBasicEmotion}}`;
            document.getElementById('status-detail').innerText = `細分情緒: ${{selectedDetailedEmotions.length ? selectedDetailedEmotions.join(', ') : '全部'}}`;
            document.getElementById('status-theme').innerText = `主題鎖定: ${{clickedTheme ? clickedTheme : '無'}}`;
            document.getElementById('status-loc').innerText = `地點鎖定: ${{clickedLocation ? clickedLocation : '無'}}`;
            document.getElementById('status-time').innerText = `時間區間: ${{filterStartDate || '全期'}} ~ ${{filterEndDate || '全期'}}`;

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
                parts.forEach(p => {{
                    if(p === 'null' || !p) return;
                    counts[p] = (counts[p] || 0) + 1;
                }});
            }});
            return Object.entries(counts).sort((a,b) => b[1] - a[1]);
        }}

        function renderThemeChart(data) {{
            const sortedData = getCounts(data, 'theme');
            const labels = sortedData.map(d => d[0]);
            const counts = sortedData.map(d => Number(d[1] || 0));

            let barColor = PALETTE.blue;
            if (currentBasicEmotion === '正面') barColor = PALETTE.green;
            if (currentBasicEmotion === '負面') barColor = PALETTE.red;

            const themeTotal = counts.reduce((s, n) => s + n, 0);
            const badgeEl = document.getElementById('theme-count-badge');
            if (badgeEl) {{
                badgeEl.innerText = `${{themeTotal}}`;
                badgeEl.style.backgroundColor = barColor;
            }}

            if (labels.length === 0) {{
                if (themeChartInstance) themeChartInstance.destroy();
                themeChartInstance = null;
                return;
            }}

            if (themeChartInstance) themeChartInstance.destroy();

            const ctx = document.getElementById('themeChart').getContext('2d');
            themeChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        label: '提及次數',
                        data: counts,
                        backgroundColor: barColor,
                        borderRadius: 6,
                        maxBarThickness: 25
                    }}]
                }},
                options: {{
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    scales: {{ x: {{ grid: {{ display: false }}, ticks: {{ precision: 0 }} }} }},
                    onClick: (e, elements) => {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            const clickedLabel = labels[idx];
                            clickedTheme = (clickedTheme === clickedLabel) ? null : clickedLabel;
                            applyCrossFilters();
                        }}
                    }}
                }}
            }});
        }}

        function renderModelChart(data) {{
            const sortedData = getCounts(data, 'model').slice(0, 8);
            const labels = sortedData.map(d => d[0]);
            const counts = sortedData.map(d => d[1]);

            if (modelChartInstance) modelChartInstance.destroy();
            const ctx = document.getElementById('modelChart').getContext('2d');
            modelChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        data: counts,
                        backgroundColor: PALETTE.beige,
                        borderColor: PALETTE.black,
                        borderWidth: 1,
                        borderRadius: 4
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    scales: {{ y: {{ grid: {{ drawBorder: false }}, ticks: {{ precision: 0 }} }} }}
                }}
            }});
        }}

        function renderCompetitorChart(data) {{
            const sortedData = getCounts(data, 'competitor').slice(0, 8);
            const labels = sortedData.map(d => d[0]);
            const counts = sortedData.map(d => d[1]);

            if (competitorChartInstance) competitorChartInstance.destroy();
            const ctx = document.getElementById('competitorChart').getContext('2d');
            competitorChartInstance = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: labels,
                    datasets: [{{
                        data: counts,
                        backgroundColor: '#edf2f7',
                        borderColor: PALETTE.red,
                        borderWidth: 1.5,
                        borderRadius: 4
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{ legend: {{ display: false }} }},
                    scales: {{ y: {{ ticks: {{ precision: 0 }} }} }}
                }}
            }});
        }}

        function renderMapMarkers(data) {{
            markerLayerGroup.clearLayers();

            const locCounts = {{}};
            data.forEach(d => {{
                if(d.coords && d.location && d.location !== '未提及' && d.location !== '香港') {{
                    if(!locCounts[d.location]) {{
                        locCounts[d.location] = {{ coords: d.coords, count: 0 }};
                    }}
                    locCounts[d.location].count++;
                }}
            }});

            Object.entries(locCounts).forEach(([name, info]) => {{
                let badgeColor = PALETTE.blue;
                if (currentBasicEmotion === '正面') badgeColor = PALETTE.green;
                if (currentBasicEmotion === '負面') badgeColor = PALETTE.red;

                const customIcon = L.divIcon({{
                    html: `<div style="background-color: ${{badgeColor}}; border: 2px solid white;" class="w-7 h-7 rounded-full flex items-center justify-center text-white text-xs font-bold shadow-md">${{info.count}}</div>`,
                    className: '',
                    iconSize: [28, 28]
                }});

                const marker = L.marker(info.coords, {{ icon: customIcon }}).addTo(markerLayerGroup);
                marker.bindPopup(`<b>📍 ${{name}}</b><br>當前篩選提及數: <b>${{info.count}}</b> 次<br><span style="font-size:11px;color:#888;">(點擊該點可反向穿透鎖定數據列)</span>`);

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
                tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-gray-400 font-medium">沒有符合當前動態交叉篩選條件的發文數據。</td></tr>`;
                return;
            }}

            data.forEach(row => {{
                let emotionBadge = '';
                if(row.emotion_basic === '正面') emotionBadge = `<span class="bg-emerald-50 text-emerald-700 px-2.5 py-1 rounded-full font-semibold text-xs border border-emerald-200">🟢 正面 [${{row.emotion_detailed}}]</span>`;
                else if(row.emotion_basic === '負面') emotionBadge = `<span class="bg-rose-50 text-rose-700 px-2.5 py-1 rounded-full font-semibold text-xs border border-rose-200">🔴 負面 [${{row.emotion_detailed}}]</span>`;
                else emotionBadge = `<span class="bg-gray-100 text-gray-700 px-2.5 py-1 rounded-full font-semibold text-xs border border-gray-200">⚪ 中性 [${{row.emotion_detailed}}]</span>`;

                const tr = document.createElement('tr');
                tr.className = "hover:bg-gray-50/80 transition-colors";

                tr.innerHTML = `
                    <td class="p-3 text-center text-gray-500 font-mono text-xs">${{row.id}}</td>
                    <td class="p-3 text-xs font-semibold text-gray-600 font-mono">${{row.time}}</td>
                    <td class="p-3 text-xs font-medium text-gray-600" title="${{row.source}}" style="min-width:130px; max-width:260px; white-space:normal;">${{row.source}}</td>
                    <td class="p-3 text-center" style="width:180px; min-width:180px;">${{emotionBadge}}</td>
                    <td class="p-3 text-center font-medium text-gray-600 text-xs">${{row.location !== '未提及' ? '📍 ' + row.location : '<span class="text-gray-300">-</span>'}}</td>
                    <td class="p-3 text-gray-700 leading-relaxed text-xs max-w-sm md:max-w-xl break-words">${{row.text}}</td>
                    <td class="p-3 text-center font-bold text-gray-600">${{row.likes}}</td>
                `;
                tbody.appendChild(tr);
            }});
        }}
    </script>
</body>
</html>
"""

# 將整合後的程式碼寫入本地檔案
output_filename = "ev_dashboard_new.html"
with open(output_filename, "w", encoding="utf-8") as f:
    f.write(html_content)

print("\n" + "=" * 50)
print(f"🎉 頂層歷史大盤位置更換與日期清洗完畢！")
print(f"1. 70%-30% 的歷史趨勢柱狀圖與情緒比例圓餅圖，已完美挪移至黑色標題 (Header) 正下方。")
print(f"2. 表格 (Table) 的時間全量轉換為標準且純淨的『YYYY-MM-DD』，徹底消滅混亂的 Excel int 數字格式。")
print(f"3. 核心可動態操控控制區完美相連，交互順暢無比。")
print("=" * 50 + "\n")