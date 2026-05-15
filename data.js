// 眾旺餐飲 成本資料
// 最後更新：2026-05-15
// 來源：restaurant_v19.xlsx
//
// =============================================
// 更新方式：將新版 xlsx 傳給 Claude，取得新的 data.js 替換此檔案
// =============================================

// 食材庫（90筆）
// 欄位：id, name, spec(規格), price(進貨單價), unit_cost(單位成本), unit(單位), category(類別), store(門店)
const MATERIALS = [
  {
    "id": "ING001",
    "name": "天裕芝麻醬",
    "spec": "18kg/桶",
    "price": 1570.0,
    "unit_cost": 0.087222,
    "unit": "g",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING002",
    "name": "竣佑芝麻醬",
    "spec": "18kg/桶",
    "price": 2800.0,
    "unit_cost": 0.155556,
    "unit": "g",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING003",
    "name": "四川花椒辣油",
    "spec": "2.5kg/包",
    "price": 520.0,
    "unit_cost": 0.208,
    "unit": "g",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING004",
    "name": "四川藤椒油",
    "spec": "2.5kg/包",
    "price": 520.0,
    "unit_cost": 0.208,
    "unit": "g",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING005",
    "name": "四季和風胡麻醬",
    "spec": "1kg/罐",
    "price": 160.0,
    "unit_cost": 0.16,
    "unit": "g",
    "category": "調味醬料",
    "store": "麻媽家"
  },
  {
    "id": "ING006",
    "name": "MAEPLOY泰式醬",
    "spec": "4kg/罐",
    "price": 320.0,
    "unit_cost": 0.08,
    "unit": "g",
    "category": "調味醬料",
    "store": "麻媽家"
  },
  {
    "id": "ING007",
    "name": "萬家鄉醬油",
    "spec": "6kg/桶",
    "price": 200.0,
    "unit_cost": 0.033333,
    "unit": "ml",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING008",
    "name": "鎮江香醋",
    "spec": "4.8kg/桶",
    "price": 110.0,
    "unit_cost": 0.022917,
    "unit": "ml",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING009",
    "name": "大蒜",
    "spec": "600g/斤",
    "price": 95.0,
    "unit_cost": 0.158333,
    "unit": "g",
    "category": "辛香料",
    "store": "ALL"
  },
  {
    "id": "ING010",
    "name": "特砂糖",
    "spec": "50kg/袋",
    "price": 1780.0,
    "unit_cost": 0.0356,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING011",
    "name": "料水粉",
    "spec": "2kg/包",
    "price": 520.0,
    "unit_cost": 0.26,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING012",
    "name": "檸檬汁",
    "spec": "800ml/瓶",
    "price": 102.0,
    "unit_cost": 0.1275,
    "unit": "ml",
    "category": "飲品原料",
    "store": "ALL"
  },
  {
    "id": "ING013",
    "name": "丸進味噌",
    "spec": "9kg/箱",
    "price": 400.0,
    "unit_cost": 0.044444,
    "unit": "g",
    "category": "（未分類）",
    "store": "ALL"
  },
  {
    "id": "ING014",
    "name": "味霖",
    "spec": "1.6L/桶",
    "price": 160.0,
    "unit_cost": 0.1,
    "unit": "ml",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING015",
    "name": "鰹魚粉",
    "spec": "1kg/盒",
    "price": 210.0,
    "unit_cost": 0.21,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING016",
    "name": "柴魚花",
    "spec": "600g/袋",
    "price": 260.0,
    "unit_cost": 0.433333,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING017",
    "name": "東北骨湯組",
    "spec": "1180g/組",
    "price": 300.0,
    "unit_cost": 0.254237,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "ALL"
  },
  {
    "id": "ING018",
    "name": "冬蔭功醬",
    "spec": "454g/瓶",
    "price": 82.0,
    "unit_cost": 0.180617,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "ALL"
  },
  {
    "id": "ING019",
    "name": "烤肉醬",
    "spec": "1kg/瓶",
    "price": 145.0,
    "unit_cost": 0.145,
    "unit": "g",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING020",
    "name": "佛特蒙咖哩塊",
    "spec": "10kg/箱",
    "price": 2290.0,
    "unit_cost": 0.229,
    "unit": "g",
    "category": "調味醬料",
    "store": "麻媽家"
  },
  {
    "id": "ING021",
    "name": "洋蔥",
    "spec": "600g/斤",
    "price": 24.0,
    "unit_cost": 0.04,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING022",
    "name": "紅蘿蔔",
    "spec": "600g/斤",
    "price": 32.0,
    "unit_cost": 0.053333,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING023",
    "name": "馬鈴薯",
    "spec": "600g/斤",
    "price": 30.0,
    "unit_cost": 0.05,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING024",
    "name": "鰹魚醬油",
    "spec": "1.8L/瓶",
    "price": 169.0,
    "unit_cost": 0.093889,
    "unit": "ml",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING025",
    "name": "美食家沙拉油",
    "spec": "3kg/桶",
    "price": 160.0,
    "unit_cost": 0.053333,
    "unit": "ml",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING026",
    "name": "小麥澱粉",
    "spec": "25kg/袋",
    "price": 750.0,
    "unit_cost": 0.03,
    "unit": "g",
    "category": "主食澱粉",
    "store": "ALL"
  },
  {
    "id": "ING027",
    "name": "高筋麵粉",
    "spec": "22kg/袋",
    "price": 600.0,
    "unit_cost": 0.027273,
    "unit": "g",
    "category": "主食澱粉",
    "store": "ALL"
  },
  {
    "id": "ING028",
    "name": "涼皮改良劑",
    "spec": "1kg/袋",
    "price": 190.0,
    "unit_cost": 0.19,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING029",
    "name": "振芳保水劑",
    "spec": "1kg/袋",
    "price": 350.0,
    "unit_cost": 0.35,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING030",
    "name": "鹽",
    "spec": "24kg/袋",
    "price": 330.0,
    "unit_cost": 0.01375,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING031",
    "name": "涼粉",
    "spec": "500g/克",
    "price": 140.0,
    "unit_cost": 0.28,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "ALL"
  },
  {
    "id": "ING032",
    "name": "百香原汁",
    "spec": "800ml/瓶",
    "price": 110.0,
    "unit_cost": 0.1375,
    "unit": "ml",
    "category": "飲品原料",
    "store": "ALL"
  },
  {
    "id": "ING033",
    "name": "決明子散裝",
    "spec": "600g/斤",
    "price": 60.0,
    "unit_cost": 0.1,
    "unit": "g",
    "category": "飲品原料",
    "store": "ALL"
  },
  {
    "id": "ING034",
    "name": "麥茶",
    "spec": "600g/斤",
    "price": 30.0,
    "unit_cost": 0.05,
    "unit": "g",
    "category": "飲品原料",
    "store": "ALL"
  },
  {
    "id": "ING035",
    "name": "仙草汁",
    "spec": "3kg/罐",
    "price": 130.0,
    "unit_cost": 0.043333,
    "unit": "ml",
    "category": "飲品原料",
    "store": "ALL"
  },
  {
    "id": "ING036",
    "name": "鴨血",
    "spec": "450g/片",
    "price": 12.0,
    "unit_cost": 0.026667,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "ALL"
  },
  {
    "id": "ING037",
    "name": "麻辣滷包",
    "spec": "500g/袋",
    "price": 148.5,
    "unit_cost": 0.297,
    "unit": "g",
    "category": "調味醬料",
    "store": "ALL"
  },
  {
    "id": "ING038",
    "name": "紅燒牛肉湯包",
    "spec": "2kg/公斤",
    "price": 180.0,
    "unit_cost": 0.09,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "ALL"
  },
  {
    "id": "ING039",
    "name": "骨湯T(銀)",
    "spec": "210g/袋",
    "price": 75.0,
    "unit_cost": 0.357143,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "ALL"
  },
  {
    "id": "ING040",
    "name": "骨湯D(油)",
    "spec": "600g/袋",
    "price": 175.0,
    "unit_cost": 0.291667,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "ALL"
  },
  {
    "id": "ING041",
    "name": "骨湯B(奶)",
    "spec": "370g/袋",
    "price": 155.0,
    "unit_cost": 0.418919,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "ALL"
  },
  {
    "id": "ING042",
    "name": "涼麵",
    "spec": "600g/斤",
    "price": 18.0,
    "unit_cost": 0.03,
    "unit": "g",
    "category": "主食澱粉",
    "store": "ALL"
  },
  {
    "id": "ING043",
    "name": "涼皮",
    "spec": "1公斤",
    "price": 25.0,
    "unit_cost": 0.025,
    "unit": "g",
    "category": "主食澱粉",
    "store": "ALL"
  },
  {
    "id": "ING044",
    "name": "壽司米",
    "spec": "60kg/袋",
    "price": 1300.0,
    "unit_cost": 0.021667,
    "unit": "g",
    "category": "主食澱粉",
    "store": "ALL"
  },
  {
    "id": "ING045",
    "name": "王子麵",
    "spec": "40個/箱",
    "price": 190.0,
    "unit_cost": 4.75,
    "unit": "個",
    "category": "主食澱粉",
    "store": "ALL"
  },
  {
    "id": "ING046",
    "name": "岑品冷凍豬肉片",
    "spec": "1kg/包",
    "price": 130.0,
    "unit_cost": 0.13,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "ALL"
  },
  {
    "id": "ING047",
    "name": "岑品牛五花片",
    "spec": "1kg/包",
    "price": 235.0,
    "unit_cost": 0.235,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "ALL"
  },
  {
    "id": "ING048",
    "name": "杏包菇",
    "spec": "600g/斤",
    "price": 70.0,
    "unit_cost": 0.116667,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING049",
    "name": "金針菇",
    "spec": "1kg/包",
    "price": 62.0,
    "unit_cost": 0.062,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING050",
    "name": "凍豆腐",
    "spec": "150g/包",
    "price": 225.0,
    "unit_cost": 1.5,
    "unit": "g",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING051",
    "name": "玉米筍",
    "spec": "10個/盒",
    "price": 28.0,
    "unit_cost": 2.8,
    "unit": "個",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING052",
    "name": "高麗菜",
    "spec": "10kg/斤",
    "price": 290.0,
    "unit_cost": 0.029,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING053",
    "name": "黃豆芽菜",
    "spec": "600g/斤",
    "price": 23.0,
    "unit_cost": 0.038333,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING054",
    "name": "冷凍青花菜",
    "spec": "1kg/包",
    "price": 60.0,
    "unit_cost": 0.06,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING055",
    "name": "蔥花",
    "spec": "1kg/包",
    "price": 70.0,
    "unit_cost": 0.07,
    "unit": "g",
    "category": "辛香料",
    "store": "ALL"
  },
  {
    "id": "ING056",
    "name": "立民黑滷",
    "spec": "20kg/袋",
    "price": 5880.0,
    "unit_cost": 0.294,
    "unit": "g",
    "category": "調味醬料",
    "store": "麻媽家"
  },
  {
    "id": "ING057",
    "name": "牛丼粉A",
    "spec": "1kg/包",
    "price": 690.0,
    "unit_cost": 0.69,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "麻媽家"
  },
  {
    "id": "ING058",
    "name": "牛丼粉B",
    "spec": "1kg/包",
    "price": 338.0,
    "unit_cost": 0.338,
    "unit": "g",
    "category": "湯底醬汁",
    "store": "麻媽家"
  },
  {
    "id": "ING059",
    "name": "滷筍絲",
    "spec": "600g/斤",
    "price": 75.0,
    "unit_cost": 0.125,
    "unit": "g",
    "category": "糖鹽調味",
    "store": "麻媽家"
  },
  {
    "id": "ING060",
    "name": "雞蛋",
    "spec": "12kg/框",
    "price": 840.0,
    "unit_cost": 0.07,
    "unit": "個",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING061",
    "name": "泡菜",
    "spec": "600g/斤",
    "price": 80.0,
    "unit_cost": 0.133333,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING062",
    "name": "烤豬肉片",
    "spec": "1kg/公斤",
    "price": 245.0,
    "unit_cost": 0.245,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "麻媽家"
  },
  {
    "id": "ING063",
    "name": "嫩雞肉片",
    "spec": "3kg/公斤",
    "price": 435.0,
    "unit_cost": 0.145,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "ALL"
  },
  {
    "id": "ING064",
    "name": "滷蛋",
    "spec": "20個/包",
    "price": 160.0,
    "unit_cost": 8.0,
    "unit": "個",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING065",
    "name": "滷蛋半顆",
    "spec": "20個/包",
    "price": 160.0,
    "unit_cost": 8.0,
    "unit": "個",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING066",
    "name": "岑品冷凍梅花豬肉片",
    "spec": "1kg/包",
    "price": 245.0,
    "unit_cost": 0.245,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "麻媽家"
  },
  {
    "id": "ING067",
    "name": "白芝麻",
    "spec": "600g/斤",
    "price": 11.0,
    "unit_cost": 0.018333,
    "unit": "g",
    "category": "辛香料",
    "store": "ALL"
  },
  {
    "id": "ING068",
    "name": "小黃瓜",
    "spec": "600g/斤",
    "price": 21.0,
    "unit_cost": 0.035,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING069",
    "name": "火腿片",
    "spec": "600g/斤",
    "price": 12.0,
    "unit_cost": 0.02,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "ALL"
  },
  {
    "id": "ING070",
    "name": "玉子燒",
    "spec": "32個/包",
    "price": 160.0,
    "unit_cost": 5.0,
    "unit": "個",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING071",
    "name": "海帶芽",
    "spec": "3.6kg/斤",
    "price": 65.0,
    "unit_cost": 0.018056,
    "unit": "g",
    "category": "辛香料",
    "store": "ALL"
  },
  {
    "id": "ING072",
    "name": "板豆腐",
    "spec": "950g/盒",
    "price": 45.0,
    "unit_cost": 0.047368,
    "unit": "g",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING073",
    "name": "貢丸",
    "spec": "3kg/袋",
    "price": 500.0,
    "unit_cost": 0.166667,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "ALL"
  },
  {
    "id": "ING074",
    "name": "川耳(木耳)",
    "spec": "600g/斤",
    "price": 65.0,
    "unit_cost": 0.108333,
    "unit": "g",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING075",
    "name": "乾絲",
    "spec": "600g/斤",
    "price": 45.0,
    "unit_cost": 0.075,
    "unit": "g",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING076",
    "name": "豆干",
    "spec": "600g/斤",
    "price": 32.0,
    "unit_cost": 0.053333,
    "unit": "g",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING077",
    "name": "嫩豆腐",
    "spec": "2.2kg/盒",
    "price": 60.0,
    "unit_cost": 0.027273,
    "unit": "g",
    "category": "豆蛋乳製品",
    "store": "ALL"
  },
  {
    "id": "ING078",
    "name": "青江菜",
    "spec": "600g/斤",
    "price": 27.0,
    "unit_cost": 0.045,
    "unit": "g",
    "category": "蔬菜",
    "store": "ALL"
  },
  {
    "id": "ING079",
    "name": "醬香牛腱肉",
    "spec": "1kg/公斤",
    "price": 485.0,
    "unit_cost": 0.485,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "涼三伯"
  },
  {
    "id": "ING080",
    "name": "酸菜",
    "spec": "2.7kg/公斤",
    "price": 250.0,
    "unit_cost": 0.092593,
    "unit": "g",
    "category": "蔬菜",
    "store": "涼三伯"
  },
  {
    "id": "ING081",
    "name": "紅燒牛筋",
    "spec": "1kg/公斤",
    "price": 435.0,
    "unit_cost": 0.435,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "涼三伯"
  },
  {
    "id": "ING082",
    "name": "紅燒牛肚",
    "spec": "1kg/公斤",
    "price": 438.0,
    "unit_cost": 0.438,
    "unit": "g",
    "category": "肉品海鮮",
    "store": "涼三伯"
  },
  {
    "id": "ING083",
    "name": "新輝赤肉羹",
    "spec": "3公斤/袋",
    "price": 615.0,
    "unit_cost": 0.205,
    "unit": "g",
    "category": "涼三伯",
    "store": "肉品海鮮"
  },
  {
    "id": "ING084",
    "name": "香菜",
    "spec": "600g/斤",
    "price": 80.0,
    "unit_cost": 0.133333,
    "unit": "g",
    "category": "ALL",
    "store": "蔬菜"
  },
  {
    "id": "ING085",
    "name": "一滴香油",
    "spec": "3L/桶",
    "price": 270.0,
    "unit_cost": 0.09,
    "unit": "ml",
    "category": "ALL",
    "store": "調味醬料"
  },
  {
    "id": "ING086",
    "name": "白胡椒粉",
    "spec": "g(單位成本0.04/g)",
    "price": 40.0,
    "unit_cost": 0.04,
    "unit": "g",
    "category": "ALL",
    "store": "糖鹽調味"
  },
  {
    "id": "ING087",
    "name": "烏醋",
    "spec": "600ml/瓶",
    "price": 225.0,
    "unit_cost": 0.375,
    "unit": "ml",
    "category": "ALL",
    "store": "調味醬料"
  },
  {
    "id": "ING088",
    "name": "毛豆仁",
    "spec": "g(單位成本0.135/g)",
    "price": 135.0,
    "unit_cost": 0.135,
    "unit": "g",
    "category": "涼三伯",
    "store": "蔬菜"
  },
  {
    "id": "ING089",
    "name": "紅甜椒",
    "spec": "g(單位成本0.12/g)",
    "price": 120.0,
    "unit_cost": 0.12,
    "unit": "g",
    "category": "涼三伯",
    "store": "蔬菜"
  },
  {
    "id": "ING090",
    "name": "冷凍蝦仁",
    "spec": "g",
    "price": 459.0,
    "unit_cost": 0.505507,
    "unit": "g",
    "category": "涼三伯",
    "store": "肉品海鮮"
  }
];

// 包材庫（34筆）
// 欄位：id, name, spec(規格), price(進貨單價), qty(每箱數量), unit_cost(單位成本/個), store(門店)
const PACKAGING = [
  {
    "id": "PKG001",
    "name": "902乾拌紙碗",
    "spec": "600個/箱",
    "price": 850,
    "qty": 600,
    "unit_cost": 1.4167,
    "store": "麻媽家"
  },
  {
    "id": "PKG002",
    "name": "165蓋子",
    "spec": "600個/箱",
    "price": 480,
    "qty": 600,
    "unit_cost": 0.8,
    "store": "ALL"
  },
  {
    "id": "PKG003",
    "name": "免洗筷子",
    "spec": "2400個/件",
    "price": 500,
    "qty": 2400,
    "unit_cost": 0.2083,
    "store": "ALL"
  },
  {
    "id": "PKG004",
    "name": "390S湯碗",
    "spec": "1000個/箱",
    "price": 950,
    "qty": 1000,
    "unit_cost": 0.95,
    "store": "ALL"
  },
  {
    "id": "PKG005",
    "name": "520蓋子",
    "spec": "1000個/箱",
    "price": 390,
    "qty": 1000,
    "unit_cost": 0.39,
    "store": "ALL"
  },
  {
    "id": "PKG006",
    "name": "1000湯碗",
    "spec": "600個/箱",
    "price": 940,
    "qty": 600,
    "unit_cost": 1.5667,
    "store": "麻媽家"
  },
  {
    "id": "PKG007",
    "name": "850蓋子",
    "spec": "600個/條",
    "price": 360,
    "qty": 600,
    "unit_cost": 0.6,
    "store": "ALL"
  },
  {
    "id": "PKG008",
    "name": "耐熱袋",
    "spec": "110個/包",
    "price": 27,
    "qty": 110,
    "unit_cost": 0.2455,
    "store": "麻媽家"
  },
  {
    "id": "PKG009",
    "name": "單支包膜湯勺",
    "spec": "2400個/隻",
    "price": 1500,
    "qty": 2400,
    "unit_cost": 0.625,
    "store": "ALL"
  },
  {
    "id": "PKG010",
    "name": "滷肉便當602碗",
    "spec": "600個/箱",
    "price": 950,
    "qty": 600,
    "unit_cost": 1.5833,
    "store": "麻媽家"
  },
  {
    "id": "PKG011",
    "name": "滷肉飯吐司盒",
    "spec": "100個/條",
    "price": 100,
    "qty": 100,
    "unit_cost": 1.0,
    "store": "麻媽家"
  },
  {
    "id": "PKG012",
    "name": "丼飯780碗",
    "spec": "50個/條",
    "price": 90,
    "qty": 50,
    "unit_cost": 1.8,
    "store": "麻媽家"
  },
  {
    "id": "PKG013",
    "name": "165內襯",
    "spec": "1000個/箱",
    "price": 720,
    "qty": 1000,
    "unit_cost": 0.72,
    "store": "麻媽家"
  },
  {
    "id": "PKG014",
    "name": "HK-P40烤肉飯盒+蓋",
    "spec": "1個/組",
    "price": 3.6,
    "qty": 1,
    "unit_cost": 3.6,
    "store": "麻媽家"
  },
  {
    "id": "PKG015",
    "name": "HK-108四格小菜盒",
    "spec": "1個/組",
    "price": 1.9,
    "qty": 1,
    "unit_cost": 1.9,
    "store": "麻媽家"
  },
  {
    "id": "PKG016",
    "name": "橡皮筋",
    "spec": "100個/包",
    "price": 27,
    "qty": 100,
    "unit_cost": 0.27,
    "store": "麻媽家"
  },
  {
    "id": "PKG017",
    "name": "涼麵中602碗",
    "spec": "600個/箱",
    "price": 850,
    "qty": 600,
    "unit_cost": 1.4167,
    "store": "麻媽家"
  },
  {
    "id": "PKG018",
    "name": "902大碗(彩盒用)",
    "spec": "600個/箱",
    "price": 850,
    "qty": 600,
    "unit_cost": 1.4167,
    "store": "ALL"
  },
  {
    "id": "PKG019",
    "name": "彩盒內襯",
    "spec": "600個/箱",
    "price": 720,
    "qty": 600,
    "unit_cost": 1.2,
    "store": "麻媽家"
  },
  {
    "id": "PKG020",
    "name": "520湯碗",
    "spec": "1000個/箱",
    "price": 860,
    "qty": 1000,
    "unit_cost": 0.86,
    "store": "ALL"
  },
  {
    "id": "PKG021",
    "name": "400CC紙杯",
    "spec": "1000個/箱",
    "price": 1050,
    "qty": 1000,
    "unit_cost": 1.05,
    "store": "ALL"
  },
  {
    "id": "PKG022",
    "name": "封口膜",
    "spec": "3000個/捲",
    "price": 365,
    "qty": 3000,
    "unit_cost": 0.1217,
    "store": "ALL"
  },
  {
    "id": "PKG023",
    "name": "斜口吸管(粗)",
    "spec": "250個/包",
    "price": 70,
    "qty": 250,
    "unit_cost": 0.28,
    "store": "ALL"
  },
  {
    "id": "PKG024",
    "name": "斜口吸管(細)",
    "spec": "250個/包",
    "price": 35,
    "qty": 250,
    "unit_cost": 0.14,
    "store": "ALL"
  },
  {
    "id": "PKG025",
    "name": "180ml小菜盒+蓋",
    "spec": "2000個/箱",
    "price": 3000,
    "qty": 2000,
    "unit_cost": 1.5,
    "store": "ALL"
  },
  {
    "id": "PKG026",
    "name": "850紙湯碗",
    "spec": "600個/箱",
    "price": 630,
    "qty": 600,
    "unit_cost": 1.05,
    "store": "涼三伯"
  },
  {
    "id": "PKG027",
    "name": "902碗",
    "spec": "600個/箱",
    "price": 850,
    "qty": 600,
    "unit_cost": 1.4167,
    "store": "ALL"
  },
  {
    "id": "PKG028",
    "name": "520紙湯杯(牛皮)",
    "spec": "1000個/箱",
    "price": 1150,
    "qty": 1000,
    "unit_cost": 1.15,
    "store": "涼三伯"
  },
  {
    "id": "PKG029",
    "name": "K-520湯杯蓋(金億昇)",
    "spec": "600個/箱",
    "price": 240,
    "qty": 600,
    "unit_cost": 0.4,
    "store": "涼三伯"
  },
  {
    "id": "PKG030",
    "name": "622碗(牛皮)",
    "spec": "600個/箱",
    "price": 1440,
    "qty": 600,
    "unit_cost": 2.4,
    "store": "涼三伯"
  },
  {
    "id": "PKG031",
    "name": "HK-165-6三格內襯",
    "spec": "600個/箱",
    "price": 720,
    "qty": 600,
    "unit_cost": 1.2,
    "store": "涼三伯"
  },
  {
    "id": "PKG032",
    "name": "透明小盒+蓋 LT-4",
    "spec": "1個/組",
    "price": 3.5,
    "qty": 1,
    "unit_cost": 3.5,
    "store": "涼三伯"
  },
  {
    "id": "PKG033",
    "name": "902圓形扁碗(牛皮)",
    "spec": "600個/箱",
    "price": 1260,
    "qty": 600,
    "unit_cost": 2.1,
    "store": "涼三伯"
  },
  {
    "id": "PKG034",
    "name": "孟宗竹筷",
    "spec": "2400雙/件",
    "price": 500,
    "qty": 2400,
    "unit_cost": 0.2083,
    "store": "ALL"
  }
];

// 半成品主檔（55筆）
// 欄位：name, total_cost(總材料成本), yield_g(產出量g), unit_cost(單位成本/g), note(備註)
const SEMIS = [
  {
    "name": "蒜水",
    "total_cost": 39.58,
    "yield_g": 1000.0,
    "unit_cost": 0.039583,
    "note": ""
  },
  {
    "name": "天裕芝麻水",
    "total_cost": 43.61,
    "yield_g": 1000.0,
    "unit_cost": 0.043611,
    "note": ""
  },
  {
    "name": "竣佑芝麻水",
    "total_cost": 77.78,
    "yield_g": 1000.0,
    "unit_cost": 0.077778,
    "note": ""
  },
  {
    "name": "濃糖",
    "total_cost": 35.6,
    "yield_g": 1680.0,
    "unit_cost": 0.02119,
    "note": ""
  },
  {
    "name": "料水",
    "total_cost": 78.0,
    "yield_g": 4500.0,
    "unit_cost": 0.017333,
    "note": ""
  },
  {
    "name": "和風胡麻醬汁",
    "total_cost": 160.0,
    "yield_g": 2000.0,
    "unit_cost": 0.08,
    "note": ""
  },
  {
    "name": "泰式酸甜醬汁",
    "total_cost": 107.0,
    "yield_g": 2600.0,
    "unit_cost": 0.041154,
    "note": ""
  },
  {
    "name": "味噌湯底",
    "total_cost": 34.44,
    "yield_g": 4460.0,
    "unit_cost": 0.007721,
    "note": ""
  },
  {
    "name": "東北骨湯湯底",
    "total_cost": 405.0,
    "yield_g": 16180.0,
    "unit_cost": 0.025031,
    "note": ""
  },
  {
    "name": "韓式烤肉醬汁",
    "total_cost": 145.0,
    "yield_g": 3500.0,
    "unit_cost": 0.041429,
    "note": ""
  },
  {
    "name": "咖哩醬",
    "total_cost": 282.31,
    "yield_g": 9250.0,
    "unit_cost": 0.03052,
    "note": ""
  },
  {
    "name": "涼皮麵漿",
    "total_cost": 199.09,
    "yield_g": 14800.0,
    "unit_cost": 0.013452,
    "note": ""
  },
  {
    "name": "冰粉底",
    "total_cost": 28.0,
    "yield_g": 4100.0,
    "unit_cost": 0.006829,
    "note": ""
  },
  {
    "name": "仙草茶",
    "total_cost": 211.88,
    "yield_g": 18000.0,
    "unit_cost": 0.011771,
    "note": ""
  },
  {
    "name": "無糖麥茶",
    "total_cost": 17.5,
    "yield_g": 15300.0,
    "unit_cost": 0.001144,
    "note": ""
  },
  {
    "name": "麻辣鴨血",
    "total_cost": 388.5,
    "yield_g": 19000.0,
    "unit_cost": 0.020447,
    "note": ""
  },
  {
    "name": "糖水",
    "total_cost": 35.6,
    "yield_g": 4500.0,
    "unit_cost": 0.007911,
    "note": "引用濃糖半成品"
  },
  {
    "name": "醬油醋水",
    "total_cost": 273.64,
    "yield_g": 11700.0,
    "unit_cost": 0.023388,
    "note": "引用濃糖半成品"
  },
  {
    "name": "椒香微辣醬汁",
    "total_cost": 7.72,
    "yield_g": 140.0,
    "unit_cost": 0.055143,
    "note": "引用芝麻水/糖水/料水半成品"
  },
  {
    "name": "傳統麻醬醬汁",
    "total_cost": 6.07,
    "yield_g": 140.0,
    "unit_cost": 0.043357,
    "note": "引用芝麻水/醬油醋水/蒜水半成品"
  },
  {
    "name": "素食無蒜醬汁",
    "total_cost": 5.6,
    "yield_g": 130.0,
    "unit_cost": 0.043077,
    "note": "引用芝麻水/醬油醋水半成品"
  },
  {
    "name": "大麻醬(無蒜)",
    "total_cost": 62.29,
    "yield_g": 3000.0,
    "unit_cost": 0.020763,
    "note": "引用芝麻水/糖水/料水半成品"
  },
  {
    "name": "大麻醬(有蒜)",
    "total_cost": 145.79,
    "yield_g": 3000.0,
    "unit_cost": 0.048597,
    "note": "引用芝麻水/醬油醋水/蒜水半成品"
  },
  {
    "name": "椒香湯底",
    "total_cost": 11.87,
    "yield_g": 480.0,
    "unit_cost": 0.024729,
    "note": "引用味噌湯底/大麻醬半成品"
  },
  {
    "name": "泰式冬蔭功湯底",
    "total_cost": 11.82,
    "yield_g": 410.0,
    "unit_cost": 0.028829,
    "note": "引用東北骨湯湯底半成品"
  },
  {
    "name": "涼皮小料水",
    "total_cost": 33.23,
    "yield_g": 780.0,
    "unit_cost": 0.042603,
    "note": ""
  },
  {
    "name": "百香底",
    "total_cost": 120.28,
    "yield_g": 2300.0,
    "unit_cost": 0.052296,
    "note": ""
  },
  {
    "name": "檸檬底",
    "total_cost": 76.58,
    "yield_g": 2020.0,
    "unit_cost": 0.037911,
    "note": ""
  },
  {
    "name": "紅燒牛肉湯",
    "total_cost": 180.0,
    "yield_g": 8000.0,
    "unit_cost": 0.0225,
    "note": ""
  },
  {
    "name": "烤肉醬汁",
    "total_cost": 145.0,
    "yield_g": 3650.0,
    "unit_cost": 0.039726,
    "note": ""
  },
  {
    "name": "醬包(涼三伯標準)",
    "total_cost": 10.0,
    "yield_g": 1.0,
    "unit_cost": 10.0,
    "note": "涼三伯涼麵/涼皮標準醬包，每份成本10元"
  },
  {
    "name": "醬包(涼三伯鮮蔬)",
    "total_cost": 7.3,
    "yield_g": 1.0,
    "unit_cost": 7.3,
    "note": "涼三伯鮮蔬系列醬包，每份成本7.3元"
  },
  {
    "name": "肉羹湯底",
    "total_cost": 249.3,
    "yield_g": 10630.4,
    "unit_cost": 0.023451,
    "note": "自製備料品，詳見BOM4"
  },
  {
    "name": "蔥油醬",
    "total_cost": 68.74,
    "yield_g": 539.1,
    "unit_cost": 0.12751,
    "note": "自製備料品，詳見BOM4"
  },
  {
    "name": "黑胡椒醬",
    "total_cost": 44.93,
    "yield_g": 508.8,
    "unit_cost": 0.088302,
    "note": "自製備料品，詳見BOM4"
  },
  {
    "name": "椒香醬(涼三伯)",
    "total_cost": 43.74,
    "yield_g": 603.25,
    "unit_cost": 0.072502,
    "note": "自製備料品，詳見BOM4"
  },
  {
    "name": "蒜蓉醬(涼三伯)",
    "total_cost": 42.09,
    "yield_g": 540.0,
    "unit_cost": 0.077947,
    "note": "自製備料品，詳見BOM4"
  },
  {
    "name": "燒肉醬",
    "total_cost": 189.5,
    "yield_g": 3300.0,
    "unit_cost": 0.057424,
    "note": "自製備料品，詳見BOM4"
  },
  {
    "name": "秘製滷蛋",
    "total_cost": 6.5,
    "yield_g": 1.0,
    "unit_cost": 6.5,
    "note": "半顆3.25元，一顆6.5元"
  },
  {
    "name": "韓式馬鈴薯",
    "total_cost": 21.57,
    "yield_g": 328.0,
    "unit_cost": 0.065749,
    "note": "BOM自製備料品 | 食材數:5"
  },
  {
    "name": "韓式黃豆芽",
    "total_cost": 15.26,
    "yield_g": 330.0,
    "unit_cost": 0.046229,
    "note": "BOM自製備料品 | 食材數:3"
  },
  {
    "name": "韓式海帶芽",
    "total_cost": 28.16,
    "yield_g": 320.0,
    "unit_cost": 0.087986,
    "note": "BOM自製備料品 | 食材數:2"
  },
  {
    "name": "韓式龍鬚菜",
    "total_cost": 48.16,
    "yield_g": 320.0,
    "unit_cost": 0.150486,
    "note": "BOM自製備料品 | 食材數:2"
  },
  {
    "name": "韓式紅蘿蔔",
    "total_cost": 15.66,
    "yield_g": 320.0,
    "unit_cost": 0.048924,
    "note": "BOM自製備料品 | 食材數:2"
  },
  {
    "name": "韓式拌飯醬（冷調）",
    "total_cost": 252.0,
    "yield_g": 2020.0,
    "unit_cost": 0.124752,
    "note": "BOM自製備料品 | 食材數:8"
  },
  {
    "name": "韓式香油",
    "total_cost": 37.5,
    "yield_g": 450.0,
    "unit_cost": 0.083333,
    "note": "BOM自製備料品 | 食材數:1"
  },
  {
    "name": "咖哩基底醬",
    "total_cost": 347.06,
    "yield_g": 8559.5,
    "unit_cost": 0.040546,
    "note": "BOM自製備料品 | 食材數:6"
  },
  {
    "name": "味噌湯",
    "total_cost": 41.57,
    "yield_g": 7932.6,
    "unit_cost": 0.00524,
    "note": "BOM自製備料品 | 食材數:5"
  },
  {
    "name": "關東煮湯底",
    "total_cost": 22.74,
    "yield_g": 6553.8,
    "unit_cost": 0.00347,
    "note": "BOM自製備料品 | 食材數:7"
  },
  {
    "name": "關東煮醬汁",
    "total_cost": 9.43,
    "yield_g": 194.75,
    "unit_cost": 0.048411,
    "note": "BOM自製備料品 | 食材數:6"
  },
  {
    "name": "黏糊糊麻醬",
    "total_cost": 7.65,
    "yield_g": 100.0,
    "unit_cost": 0.076538,
    "note": "BOM自製備料品 | 食材數:2"
  },
  {
    "name": "三杯醬",
    "total_cost": 3.12,
    "yield_g": 152.0,
    "unit_cost": 0.020493,
    "note": "BOM自製備料品 | 食材數:8"
  },
  {
    "name": "醃製雞丁",
    "total_cost": 185.02,
    "yield_g": 1106.75,
    "unit_cost": 0.167178,
    "note": "BOM自製備料品 | 食材數:6"
  },
  {
    "name": "原牛丼汁",
    "total_cost": 116.75,
    "yield_g": 3344.0,
    "unit_cost": 0.034913,
    "note": "BOM自製備料品 | 食材數:2"
  },
  {
    "name": "傳統麻醬",
    "total_cost": 12.77,
    "yield_g": 3135.0,
    "unit_cost": 0.004073,
    "note": "BOM自製備料品 | 食材數:6"
  }
];

// 產品主檔（83筆）
// 欄位：name, category(類別), store(門店), price(售價), cost(單份成本), cost_pct(成本率%)
const PRODUCTS = [
  {
    "name": "黑金滷肉飯(中)",
    "category": "飯類",
    "store": "麻媽家",
    "price": 40.0,
    "cost": 23.77,
    "cost_pct": 59.4
  },
  {
    "name": "黑金滷肉飯(大)",
    "category": "飯類",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 31.53,
    "cost_pct": 63.1
  },
  {
    "name": "椒香滷肉飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 25.84,
    "cost_pct": 51.7
  },
  {
    "name": "滷肉飯便單(含滷蛋)",
    "category": "飯類",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 59.52,
    "cost_pct": 66.1
  },
  {
    "name": "雪花豬丼飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 29.02,
    "cost_pct": 32.2
  },
  {
    "name": "滑蛋雪花豬丼",
    "category": "飯類",
    "store": "麻媽家",
    "price": 105.0,
    "cost": 33.92,
    "cost_pct": 32.3
  },
  {
    "name": "雪花牛丼飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 100.0,
    "cost": 39.54,
    "cost_pct": 39.5
  },
  {
    "name": "滑蛋雪花牛丼",
    "category": "飯類",
    "store": "麻媽家",
    "price": 115.0,
    "cost": 44.44,
    "cost_pct": 38.6
  },
  {
    "name": "泡菜牛丼飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 115.0,
    "cost": 46.2,
    "cost_pct": 40.2
  },
  {
    "name": "滑蛋泡菜牛丼飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 130.0,
    "cost": 51.1,
    "cost_pct": 39.3
  },
  {
    "name": "香香雞丼飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 27.64,
    "cost_pct": 30.7
  },
  {
    "name": "滑蛋香香雞丼飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 105.0,
    "cost": 32.54,
    "cost_pct": 31.0
  },
  {
    "name": "韓式烤牛飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 125.0,
    "cost": 33.63,
    "cost_pct": 26.9
  },
  {
    "name": "肉多多烤牛飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 155.0,
    "cost": 40.19,
    "cost_pct": 25.9
  },
  {
    "name": "韓式烤豬飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 125.0,
    "cost": 38.06,
    "cost_pct": 30.4
  },
  {
    "name": "肉多多烤豬飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 155.0,
    "cost": 43.79,
    "cost_pct": 28.3
  },
  {
    "name": "家常咖哩飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 70.0,
    "cost": 26.55,
    "cost_pct": 37.9
  },
  {
    "name": "烤牛咖哩飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 125.0,
    "cost": 36.92,
    "cost_pct": 29.5
  },
  {
    "name": "椒香烤牛咖哩飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 135.0,
    "cost": 39.0,
    "cost_pct": 28.9
  },
  {
    "name": "嫩雞咖哩飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 30.9,
    "cost_pct": 34.3
  },
  {
    "name": "椒香家常咖哩飯",
    "category": "飯類",
    "store": "麻媽家",
    "price": 80.0,
    "cost": 28.63,
    "cost_pct": 35.8
  },
  {
    "name": "涼麵(中)-椒香",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 15.59,
    "cost_pct": 31.2
  },
  {
    "name": "涼麵(中)-傳統",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 14.12,
    "cost_pct": 28.2
  },
  {
    "name": "涼麵(中)-胡麻",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 18.7,
    "cost_pct": 37.4
  },
  {
    "name": "涼麵(中)-泰式",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 13.84,
    "cost_pct": 27.7
  },
  {
    "name": "涼麵(中)-素食",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 14.08,
    "cost_pct": 28.2
  },
  {
    "name": "涼麵(大)-椒香",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 60.0,
    "cost": 17.99,
    "cost_pct": 30.0
  },
  {
    "name": "涼麵(大)-傳統",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 60.0,
    "cost": 16.52,
    "cost_pct": 27.5
  },
  {
    "name": "涼麵(大)-胡麻",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 60.0,
    "cost": 21.1,
    "cost_pct": 35.2
  },
  {
    "name": "涼麵(大)-泰式",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 60.0,
    "cost": 16.24,
    "cost_pct": 27.1
  },
  {
    "name": "涼麵(大)-素食",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 60.0,
    "cost": 16.48,
    "cost_pct": 27.5
  },
  {
    "name": "嫩雞涼麵-椒香",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 80.0,
    "cost": 27.19,
    "cost_pct": 34.0
  },
  {
    "name": "嫩雞涼麵-傳統",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 80.0,
    "cost": 25.72,
    "cost_pct": 32.1
  },
  {
    "name": "嫩雞涼麵-胡麻",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 80.0,
    "cost": 30.3,
    "cost_pct": 37.9
  },
  {
    "name": "嫩雞涼麵-泰式",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 80.0,
    "cost": 25.44,
    "cost_pct": 31.8
  },
  {
    "name": "嫩雞涼麵-素食",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 80.0,
    "cost": 25.68,
    "cost_pct": 32.1
  },
  {
    "name": "川陝手工涼皮-椒香",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 70.0,
    "cost": 15.69,
    "cost_pct": 22.4
  },
  {
    "name": "川陝手工涼皮-傳統",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 70.0,
    "cost": 14.22,
    "cost_pct": 20.3
  },
  {
    "name": "川陝手工涼皮-胡麻",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 70.0,
    "cost": 18.8,
    "cost_pct": 26.9
  },
  {
    "name": "川陝手工涼皮-泰式",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 70.0,
    "cost": 13.94,
    "cost_pct": 19.9
  },
  {
    "name": "川陝手工涼皮-素食",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 70.0,
    "cost": 14.18,
    "cost_pct": 20.3
  },
  {
    "name": "嫩雞涼皮-椒香",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 27.29,
    "cost_pct": 30.3
  },
  {
    "name": "嫩雞涼皮-傳統",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 25.82,
    "cost_pct": 28.7
  },
  {
    "name": "嫩雞涼皮-胡麻",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 30.4,
    "cost_pct": 33.8
  },
  {
    "name": "嫩雞涼皮-泰式",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 25.54,
    "cost_pct": 28.4
  },
  {
    "name": "嫩雞涼皮-素食",
    "category": "涼麵涼皮",
    "store": "麻媽家",
    "price": 90.0,
    "cost": 25.78,
    "cost_pct": 28.6
  },
  {
    "name": "味噌湯",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 30.0,
    "cost": 5.74,
    "cost_pct": 19.1
  },
  {
    "name": "味噌蛋花湯",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 40.0,
    "cost": 10.26,
    "cost_pct": 25.6
  },
  {
    "name": "味噌貢丸湯",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 40.0,
    "cost": 13.69,
    "cost_pct": 34.2
  },
  {
    "name": "味噌三合一湯",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 50.0,
    "cost": 18.59,
    "cost_pct": 37.2
  },
  {
    "name": "檸檬冰粉",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 30.0,
    "cost": 5.96,
    "cost_pct": 19.9
  },
  {
    "name": "百香果冰粉",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 30.0,
    "cost": 6.9,
    "cost_pct": 23.0
  },
  {
    "name": "無糖麥茶",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 20.0,
    "cost": 1.72,
    "cost_pct": 8.6
  },
  {
    "name": "仙草茶",
    "category": "湯品冰粉",
    "store": "麻媽家",
    "price": 20.0,
    "cost": 5.55,
    "cost_pct": 27.7
  },
  {
    "name": "韓式泡菜",
    "category": "小菜",
    "store": "麻媽家",
    "price": 35.0,
    "cost": 14.83,
    "cost_pct": 42.4
  },
  {
    "name": "涼拌黃瓜",
    "category": "小菜",
    "store": "麻媽家",
    "price": 35.0,
    "cost": 5.35,
    "cost_pct": 15.3
  },
  {
    "name": "涼拌木耳",
    "category": "小菜",
    "store": "麻媽家",
    "price": 35.0,
    "cost": 13.42,
    "cost_pct": 38.3
  },
  {
    "name": "招牌乾絲",
    "category": "小菜",
    "store": "麻媽家",
    "price": 35.0,
    "cost": 7.5,
    "cost_pct": 21.4
  },
  {
    "name": "香滷豆干",
    "category": "小菜",
    "store": "麻媽家",
    "price": 35.0,
    "cost": 13.25,
    "cost_pct": 37.8
  },
  {
    "name": "麻辣鴨血",
    "category": "小菜",
    "store": "麻媽家",
    "price": 35.0,
    "cost": 7.38,
    "cost_pct": 21.1
  },
  {
    "name": "椒香燙青菜",
    "category": "小菜",
    "store": "麻媽家",
    "price": 40.0,
    "cost": 9.43,
    "cost_pct": 23.6
  },
  {
    "name": "藤椒燙青菜",
    "category": "小菜",
    "store": "麻媽家",
    "price": 40.0,
    "cost": 9.43,
    "cost_pct": 23.6
  },
  {
    "name": "滷汁燙青菜",
    "category": "小菜",
    "store": "麻媽家",
    "price": 40.0,
    "cost": 9.86,
    "cost_pct": 24.6
  },
  {
    "name": "肉羹湯（小）",
    "category": "新品",
    "store": "涼三伯",
    "price": 39.0,
    "cost": 12.79,
    "cost_pct": 32.8
  },
  {
    "name": "肉羹湯",
    "category": "新品",
    "store": "涼三伯",
    "price": 70.0,
    "cost": 16.89,
    "cost_pct": 24.1
  },
  {
    "name": "基礎涼麵(中)",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 55.0,
    "cost": 13.93,
    "cost_pct": 25.3
  },
  {
    "name": "基礎涼麵(大)",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 65.0,
    "cost": 13.63,
    "cost_pct": 21.0
  },
  {
    "name": "基礎涼麵(特大)",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 75.0,
    "cost": 13.63,
    "cost_pct": 18.2
  },
  {
    "name": "霸王嫩雞涼麵",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 110.0,
    "cost": 28.31,
    "cost_pct": 25.7
  },
  {
    "name": "嫩雞彩盒涼麵",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 110.0,
    "cost": 26.89,
    "cost_pct": 24.4
  },
  {
    "name": "鮮蔬彩盒涼麵",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 23.37,
    "cost_pct": 23.4
  },
  {
    "name": "高階海陸涼麵",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 140.0,
    "cost": 53.86,
    "cost_pct": 38.5
  },
  {
    "name": "川陝涼皮",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 75.0,
    "cost": 13.63,
    "cost_pct": 18.2
  },
  {
    "name": "鮮蔬彩盒涼皮",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 23.37,
    "cost_pct": 23.4
  },
  {
    "name": "嫩雞彩盒涼皮",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 110.0,
    "cost": 26.89,
    "cost_pct": 24.4
  },
  {
    "name": "霸王嫩雞涼皮",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 110.0,
    "cost": 28.31,
    "cost_pct": 25.7
  },
  {
    "name": "高階海陸涼皮",
    "category": "涼麵涼皮",
    "store": "涼三伯",
    "price": 150.0,
    "cost": 53.86,
    "cost_pct": 35.9
  },
  {
    "name": "蔥油嫩雞飯",
    "category": "飯類",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 29.19,
    "cost_pct": 29.2
  },
  {
    "name": "黑胡椒雞丁飯",
    "category": "飯類",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 29.93,
    "cost_pct": 29.9
  },
  {
    "name": "椒香雞丁飯",
    "category": "飯類",
    "store": "涼三伯",
    "price": 95.0,
    "cost": 29.46,
    "cost_pct": 31.0
  },
  {
    "name": "蒜泥白肉飯",
    "category": "飯類",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 28.47,
    "cost_pct": 28.5
  },
  {
    "name": "炙燒豬肉飯",
    "category": "飯類",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 39.01,
    "cost_pct": 39.0
  },
  {
    "name": "炙燒牛肉飯",
    "category": "飯類",
    "store": "涼三伯",
    "price": 100.0,
    "cost": 38.01,
    "cost_pct": 38.0
  }
];
