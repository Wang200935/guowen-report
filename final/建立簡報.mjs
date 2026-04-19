const fs = await import("node:fs/promises");
const path = await import("node:path");
const { Presentation, PresentationFile } = await import("@oai/artifact-tool");

const W = 1280;
const H = 720;

const 輸出目錄 = path.resolve("/Users/wang/Documents/CTF/guowen-report/final");
const 參考圖目錄 = path.resolve("/Users/wang/Documents/CTF/guowen-report/tmp/slides/tainan-tech-refresh/reference");
const 暫存目錄 = path.resolve("/Users/wang/Documents/CTF/guowen-report/tmp/slides/tainan-tech-refresh");
const 預覽目錄 = path.join(暫存目錄, "preview");
const 驗證目錄 = path.join(暫存目錄, "verification");
const 檢查紀錄 = path.join(暫存目錄, "inspect.ndjson");

const 色票 = {
  深夜藍: "#071521",
  深海藍: "#0B2031",
  霧藍: "#112A40",
  青綠: "#38D6C8",
  淺青: "#A9FFF8",
  金色: "#F4C96B",
  珊瑚: "#FF7C63",
  紙白: "#F7FBFF",
  霧白: "#F7FBFFD8",
  淺灰: "#D8E5F2",
  中灰: "#9DB1C5",
  墨黑: "#091017",
  透明: "#00000000",
};

const 標題字 = "Aptos";
const 內文字 = "Aptos";
const 等寬字 = "Aptos Mono";

const 頁面資料 = [
  {
    類型: "封面",
    標籤: "國文分組報告",
    標題: "臺南城市變遷歷史",
    副標題: "從港口據點到大臺南",
    階段: [
      ["01", "開港建城", "先因海港而興"],
      ["02", "府城定型", "行政與商業核心成形"],
      ["03", "近代轉型", "空間秩序重新設計"],
      ["04", "當代再定位", "歷史與現代並存"],
    ],
  },
  {
    類型: "總覽",
    標籤: "主線總覽",
    標題: "四個時代，看懂臺南怎麼變",
    副標題: "從海港據點到今天的大臺南，城市角色一共經過四次明顯轉折。",
    時間線: [
      ["17 世紀", "開港建城", "先從海港與政權據點開始"],
      ["清領時期", "府城成形", "再成為全臺首府與府城"],
      ["近代化時期", "城市轉型", "接著拆城修路走向近代都市"],
      ["戰後至今", "重新定位", "最後發展成今天的大臺南"],
    ],
  },
  {
    類型: "階段",
    標籤: "第一階段",
    年代: "1624—1683",
    標題: "臺南的起點，不是古都，而是海港與政權中心",
    判斷: "城市最早的重要性，來自港口、貿易與統治需求，而不是先有文化古都形象。",
    要點: [
      "荷蘭建立熱蘭遮城與普羅民遮城後，臺南被拉進對外貿易網絡。",
      "港口、堡壘與行政據點相互扣合，形成最早的政商核心。",
      "明鄭接手後延續統治中心功能，讓城市角色沒有中斷。",
    ],
    標籤欄: [
      ["城市角色", "海港據點＋政權核心"],
      ["空間變化", "由港灣與堡壘帶動城市雛形"],
      ["留下影響", "奠定臺灣最早城市核心的起點"],
    ],
  },
  {
    類型: "階段",
    標籤: "第二階段",
    年代: "1684—19 世紀中葉",
    標題: "到了清代，臺南真正有了府城的樣子",
    判斷: "臺南不只是熱鬧港口，而是變成全臺政治、商業與文化中心。",
    要點: [
      "1684 年設臺灣府後，臺南成為行政核心，城市位階明確拉高。",
      "城牆、街區、官署與廟宇逐步穩定，生活圈與治理秩序一起成形。",
      "今天大家熟悉的府城與古都印象，就是這個階段打下的底子。",
    ],
    標籤欄: [
      ["城市角色", "全臺首府與商業文化中心"],
      ["空間變化", "由點狀據點走向穩定府城格局"],
      ["留下影響", "建立府城與古都的集體印象"],
    ],
  },
  {
    類型: "階段",
    標籤: "第三階段",
    年代: "19 世紀後期—1945",
    標題: "失去全臺首位後，臺南反而開始變現代",
    判斷: "政治中心地位改變後，臺南沒有停住，而是透過近代化改造出新的城市樣貌。",
    要點: [
      "晚清之後，港口功能與政治重心變動，城市開始面對重新調整。",
      "日治時期拆城、修路、引進新式公共建築與城市設施，空間秩序被重新設計。",
      "臺南從封閉府城走向更開放的近代都市，生活方式也跟著改變。",
    ],
    標籤欄: [
      ["城市角色", "南部重要城市與近代治理節點"],
      ["空間變化", "從城牆邊界轉向道路與公共設施"],
      ["留下影響", "完成從傳統府城到近代都市的轉型"],
    ],
  },
  {
    類型: "階段",
    標籤: "第四階段",
    年代: "1945—現在",
    標題: "戰後的臺南，不再只是舊府城，而是今天的大臺南",
    判斷: "城市不再是全臺政治中心，但透過文化保存、都市擴張與行政整合重新建立定位。",
    要點: [
      "戰後臺南的城市功能持續擴張，不再只圍繞舊城區運作。",
      "2010 年縣市合併後，城市尺度放大，形成更完整的大臺南都會空間。",
      "老城保存、新區發展與產業升級並行，讓臺南同時保有歷史與當代面貌。",
    ],
    標籤欄: [
      ["城市角色", "文化古都＋區域都會"],
      ["空間變化", "從舊城核心擴張成更大的生活圈"],
      ["留下影響", "形成今天既老又新的臺南"],
    ],
  },
  {
    類型: "統整",
    標籤: "統整頁",
    標題: "如果只記三件事，臺南到底變了什麼？",
    副標題: "看完四個階段後，可以把臺南的變化濃縮成角色、空間、定位三條線。",
    卡片: [
      ["城市角色變了", "海港據點 → 府城 → 近代都市 → 大臺南", "每一次轉折，都是城市功能重新分配的結果。"],
      ["城市空間變了", "城內核心 → 有邊界的府城 → 開放型都市 → 更大的都會生活圈", "臺南不是只有老城，而是逐步擴大到今天的都市尺度。"],
      ["城市定位變了", "政治中心 → 文化古都 → 文化與現代並存", "大家理解臺南的方式，也跟著城市角色一起改變。"],
    ],
  },
  {
    類型: "結論",
    標籤: "結論",
    標題: "臺南最特別的不是很老，\n而是一直在變",
    句子: [
      "從 海 港 據 點 到 府 城，",
      "從 傳 統 城 市 到 近 代 都 市，",
      "再 到 今 天 的 大 臺 南，",
      "臺 南 每 一 次 失 去 舊 角 色 後，",
      "都 會 長 出 新 的 城 市 樣 子。",
    ],
  },
  {
    類型: "分工",
    標籤: "分工頁",
    標題: "組員分工",
    副標題: "請依實際組員填入姓名與負責內容，版型已保留正式報告用的視覺系統。",
    表頭: ["組員", "負責內容", "簡報／上台"],
    列: [
      ["請填入", "資料整理", "請填入"],
      ["請填入", "文案撰寫", "請填入"],
      ["請填入", "簡報製作", "請填入"],
      ["請填入", "上台報告", "請填入"],
    ],
    註記: "分工頁只要清楚，不要花；真正的重點應留在前八頁。",
  },
];

const 來源資料 = {
  primary: "臺南市立博物館、臺南市政府文化局、臺南市政府民政局、Wikimedia Commons 歷史地圖與影像資料。",
};

const 檢查資料 = [];

function 正規化文字(文字) {
  if (Array.isArray(文字)) return 文字.join("\n");
  return String(文字 ?? "");
}

function 行數(文字) {
  const 值 = 正規化文字(文字);
  return 值 ? 值.split(/\n/).length : 0;
}

function 需求高度(文字, 字級, 行高 = 1.22) {
  return Math.max(12, 行數(文字) * 字級 * 行高);
}

function 檢查高度(文字, 高度, 字級, 角色) {
  if (!正規化文字(文字).trim()) return;
  const 需要 = 需求高度(文字, 字級);
  if (高度 + Math.max(4, 字級 * 0.08) < 需要) {
    throw new Error(`${角色} 文字框過短：高度 ${高度}，至少需要 ${需要}。`);
  }
}

async function 讀圖片(imagePath) {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function 加圖片(slide, 頁碼, 檔名, left, top, width, height, alt, fit = "cover") {
  const imagePath = path.join(參考圖目錄, 檔名);
  const image = slide.images.add({
    blob: await 讀圖片(imagePath),
    fit,
    alt,
  });
  image.position = { left, top, width, height };
  檢查資料.push({ kind: "image", slide: 頁碼, role: alt, path: imagePath, bbox: [left, top, width, height] });
  return image;
}

function 加形狀(slide, 頁碼, geometry, left, top, width, height, fill, lineFill = 色票.透明, lineWidth = 0, role = geometry) {
  const shape = slide.shapes.add({
    geometry,
    position: { left, top, width, height },
    fill,
    line: { style: "solid", fill: lineFill, width: lineWidth },
  });
  檢查資料.push({ kind: "shape", slide: 頁碼, role, bbox: [left, top, width, height] });
  return shape;
}

function 加文字(slide, 頁碼, 文字, left, top, width, height, {
  字級 = 20,
  顏色 = 色票.紙白,
  粗體 = false,
  字體 = 內文字,
  對齊 = "left",
  垂直 = "top",
  填色 = 色票.透明,
  邊框 = 色票.透明,
  邊框寬 = 0,
  自動縮小 = null,
  角色 = "text",
  檢查 = true,
} = {}) {
  if (檢查) 檢查高度(文字, height, 字級, 角色);
  const box = 加形狀(slide, 頁碼, "rect", left, top, width, height, 填色, 邊框, 邊框寬, `${角色}-box`);
  box.text = 文字;
  box.text.fontSize = 字級;
  box.text.color = 顏色;
  box.text.bold = 粗體;
  box.text.typeface = 字體;
  box.text.alignment = 對齊;
  box.text.verticalAlignment = 垂直;
  box.text.insets = { left: 0, right: 0, top: 0, bottom: 0 };
  if (自動縮小) box.text.autoFit = 自動縮小;
  檢查資料.push({
    kind: "textbox",
    slide: 頁碼,
    role: 角色,
    text: 正規化文字(文字),
    textChars: 正規化文字(文字).length,
    textLines: 行數(文字),
    bbox: [left, top, width, height],
  });
  return box;
}

function 加頁碼框(slide, 頁碼, 標籤) {
  加文字(slide, 頁碼, 標籤, 84, 38, 260, 22, {
    字級: 12,
    顏色: 色票.淺青,
    粗體: true,
    字體: 等寬字,
    角色: "頁首標籤",
    檢查: false,
  });
  加文字(slide, 頁碼, `第 ${String(頁碼).padStart(2, "0")} 頁`, 1096, 38, 100, 22, {
    字級: 12,
    顏色: 色票.淺灰,
    粗體: true,
    字體: 等寬字,
    對齊: "right",
    角色: "頁碼",
    檢查: false,
  });
  加形狀(slide, 頁碼, "rect", 84, 64, 1112, 2, 色票.青綠, 色票.透明, 0, "頁首線");
}

function 加圓角卡(slide, 頁碼, left, top, width, height, fill, line = "#1E3B56", role = "卡片") {
  return 加形狀(slide, 頁碼, "roundRect", left, top, width, height, fill, line, 1.2, role);
}

function 加條列(slide, 頁碼, items, left, top, width) {
  items.forEach((item, index) => {
    const y = top + index * 110;
    加形狀(slide, 頁碼, "ellipse", left, y + 12, 22, 22, 色票.青綠, 色票.透明, 0, "條列點");
    加文字(slide, 頁碼, item, left + 42, y + 2, width - 42, 68, {
      字級: 20,
      顏色: 色票.紙白,
      字體: 內文字,
      角色: `條列-${index + 1}`,
    });
  });
}

function 加標籤欄(slide, 頁碼, 欄位, left, top, width) {
  欄位.forEach(([標題, 內容], index) => {
    const y = top + index * 96;
    加圓角卡(slide, 頁碼, left, y, width, 78, "#0F2A3DE6", "#27516E", `標籤卡-${index + 1}`);
    加文字(slide, 頁碼, 標題, left + 22, y + 14, 150, 22, {
      字級: 12,
      顏色: 色票.淺青,
      粗體: true,
      字體: 等寬字,
      角色: `標籤標題-${index + 1}`,
      檢查: false,
    });
    加文字(slide, 頁碼, 內容, left + 22, y + 38, width - 44, 28, {
      字級: 18,
      顏色: 色票.紙白,
      粗體: true,
      字體: 內文字,
      角色: `標籤內容-${index + 1}`,
      檢查: false,
    });
  });
}

function 加講者備註(slide, 文字) {
  slide.speakerNotes.setText(`${文字}\n\n[Sources]\n- ${來源資料.primary}`);
}

async function 做封面(presentation, data) {
  const 頁碼 = 1;
  const slide = presentation.slides.add();
  await 加圖片(slide, 頁碼, "slide-01.png", 0, 0, W, H, "封面背景");
  加形狀(slide, 頁碼, "rect", 0, 0, W, H, "#04101B66", 色票.透明, 0, "封面整體遮罩");
  加形狀(slide, 頁碼, "rect", 0, 0, 760, H, "#061726F0", 色票.透明, 0, "封面左欄");
  加頁碼框(slide, 頁碼, data.標籤);
  加文字(slide, 頁碼, data.標題, 84, 136, 640, 126, {
    字級: 56,
    顏色: 色票.紙白,
    粗體: true,
    字體: 標題字,
    角色: "封面標題",
  });
  加文字(slide, 頁碼, data.副標題, 90, 292, 520, 50, {
    字級: 28,
    顏色: "#D7E4F2",
    粗體: true,
    字體: 內文字,
    角色: "封面副標題",
  });
  加文字(slide, 頁碼, "把臺南的變化拆成四個清楚段落，報告時一眼就能抓到主題。", 90, 386, 560, 64, {
    字級: 20,
    顏色: 色票.淺灰,
    字體: 內文字,
    角色: "封面說明",
  });

  加圓角卡(slide, 頁碼, 792, 120, 390, 446, "#081A29D8", "#25526E", "封面階段面板");
  加文字(slide, 頁碼, "目錄", 828, 154, 180, 28, {
    字級: 18,
    顏色: 色票.淺青,
    粗體: true,
    字體: 等寬字,
    角色: "封面面板標題",
    檢查: false,
  });
  data.階段.forEach(([序號, 標題, 說明], index) => {
    const y = 212 + index * 84;
    加圓角卡(slide, 頁碼, 826, y, 326, 64, "#102A40F0", "#21455D", `封面階段卡-${index + 1}`);
    加文字(slide, 頁碼, 序號, 850, y + 12, 40, 24, {
      字級: 20,
      顏色: 色票.青綠,
      粗體: true,
      字體: 標題字,
      角色: `封面序號-${index + 1}`,
      檢查: false,
    });
    加文字(slide, 頁碼, 標題, 910, y + 10, 190, 24, {
      字級: 19,
      顏色: 色票.紙白,
      粗體: true,
      字體: 內文字,
      角色: `封面階段標題-${index + 1}`,
      檢查: false,
    });
    加文字(slide, 頁碼, 說明, 910, y + 34, 208, 20, {
      字級: 13,
      顏色: 色票.淺灰,
      字體: 內文字,
      角色: `封面階段說明-${index + 1}`,
      檢查: false,
    });
  });
  加文字(slide, 頁碼, "班級／組別／姓名", 90, 938, 220, 22, {
    字級: 14,
    顏色: "#D7E4F2",
    字體: 內文字,
    角色: "封面欄位",
    檢查: false,
  });
  加講者備註(slide, "封面要先講清楚：這份報告不是景點導覽，而是用四個轉折看臺南如何改變角色與定位。");
}

async function 做總覽(presentation, data) {
  const 頁碼 = 2;
  const slide = presentation.slides.add();
  await 加圖片(slide, 頁碼, "slide-02.png", 0, 0, W, H, "總覽背景");
  加形狀(slide, 頁碼, "rect", 0, 0, W, H, "#F4FAFFEA", 色票.透明, 0, "總覽遮罩");
  加頁碼框(slide, 頁碼, data.標籤);
  加文字(slide, 頁碼, data.標題, 84, 118, 920, 66, {
    字級: 42,
    顏色: 色票.墨黑,
    粗體: true,
    字體: 標題字,
    角色: "總覽標題",
  });
  加文字(slide, 頁碼, data.副標題, 84, 206, 780, 56, {
    字級: 24,
    顏色: 色票.深海藍,
    字體: 內文字,
    角色: "總覽副標題",
  });
  data.時間線.forEach(([時間, 標題, 說明], index) => {
    const x = 72 + index * 292;
    加圓角卡(slide, 頁碼, x, 352, 268, 248, "#FFFFFFEE", "#C7DDED", `時間線卡-${index + 1}`);
    加文字(slide, 頁碼, 時間, x + 24, 378, 200, 24, {
      字級: 14,
      顏色: 色票.青綠,
      粗體: true,
      字體: 等寬字,
      角色: `時間線時間-${index + 1}`,
      檢查: false,
    });
    加文字(slide, 頁碼, 標題, x + 24, 422, 220, 58, {
      字級: 28,
      顏色: 色票.墨黑,
      粗體: true,
      字體: 標題字,
      角色: `時間線標題-${index + 1}`,
    });
    加文字(slide, 頁碼, 說明, x + 24, 500, 220, 72, {
      字級: 17,
      顏色: 色票.深海藍,
      字體: 內文字,
      角色: `時間線說明-${index + 1}`,
    });
    if (index < 3) {
      加形狀(slide, 頁碼, "rect", x + 268, 470, 24, 4, 色票.青綠, 色票.透明, 0, `時間線連接-${index + 1}`);
      加文字(slide, 頁碼, "→", x + 270, 448, 20, 20, {
        字級: 18,
        顏色: 色票.青綠,
        粗體: true,
        字體: 標題字,
        角色: `時間線箭頭-${index + 1}`,
        檢查: false,
      });
    }
  });
  加講者備註(slide, "總覽頁先讓老師知道：後面不是照朝代背，而是看臺南如何在四段時期改變自己的角色。");
}

async function 做階段頁(presentation, 頁碼, data, 參考圖) {
  const slide = presentation.slides.add();
  await 加圖片(slide, 頁碼, 參考圖, 0, 0, W, H, data.標題);
  加形狀(slide, 頁碼, "rect", 0, 0, W, H, "#05121CD8", 色票.透明, 0, "階段遮罩");
  加形狀(slide, 頁碼, "rect", 0, 0, 760, H, "#061828EB", 色票.透明, 0, "左側閱讀區");
  加形狀(slide, 頁碼, "rect", 760, 0, W - 760, H, "#081625B8", 色票.透明, 0, "右側統一濾罩");
  加頁碼框(slide, 頁碼, data.標籤);
  加圓角卡(slide, 頁碼, 84, 112, 220, 46, "#10334B", "#1E5573", "年代膠囊");
  加文字(slide, 頁碼, data.年代, 110, 124, 170, 22, {
    字級: 16,
    顏色: 色票.淺青,
    粗體: true,
    字體: 等寬字,
    角色: "年代",
    檢查: false,
  });
  加文字(slide, 頁碼, data.標題, 84, 186, 620, 116, {
    字級: 42,
    顏色: 色票.紙白,
    粗體: true,
    字體: 標題字,
    角色: "階段標題",
  });
  加文字(slide, 頁碼, data.判斷, 84, 332, 620, 74, {
    字級: 25,
    顏色: 色票.淺灰,
    粗體: true,
    字體: 內文字,
    角色: "階段判斷",
  });
  加條列(slide, 頁碼, data.要點, 84, 448, 730);
  加標籤欄(slide, 頁碼, data.標籤欄, 940, 398, 240);
  加講者備註(slide, `${data.標題}：先講判斷，再挑三個重點帶過，不要把細節年份講太滿。`);
}

async function 做統整頁(presentation, data) {
  const 頁碼 = 7;
  const slide = presentation.slides.add();
  await 加圖片(slide, 頁碼, "slide-07.png", 0, 0, W, H, data.標題);
  加形狀(slide, 頁碼, "rect", 0, 0, W, H, "#07131FD9", 色票.透明, 0, "統整遮罩");
  加頁碼框(slide, 頁碼, data.標籤);
  加文字(slide, 頁碼, data.標題, 84, 118, 920, 66, {
    字級: 40,
    顏色: 色票.紙白,
    粗體: true,
    字體: 標題字,
    角色: "統整標題",
  });
  加文字(slide, 頁碼, data.副標題, 84, 204, 840, 54, {
    字級: 22,
    顏色: 色票.淺灰,
    字體: 內文字,
    角色: "統整副標題",
  });
  data.卡片.forEach(([標題, 變化, 說明], index) => {
    const x = 74 + index * 384;
    加圓角卡(slide, 頁碼, x, 326, 356, 286, "#0D2236EE", "#204A65", `統整卡-${index + 1}`);
    加文字(slide, 頁碼, `0${index + 1}`, x + 24, 350, 34, 24, {
      字級: 20,
      顏色: index === 1 ? 色票.金色 : index === 2 ? 色票.珊瑚 : 色票.青綠,
      粗體: true,
      字體: 標題字,
      角色: `統整序號-${index + 1}`,
      檢查: false,
    });
    加文字(slide, 頁碼, 標題, x + 70, 348, 242, 26, {
      字級: 20,
      顏色: 色票.紙白,
      粗體: true,
      字體: 內文字,
      角色: `統整標題-${index + 1}`,
      檢查: false,
    });
    加文字(slide, 頁碼, 變化, x + 24, 410, 306, 96, {
      字級: 28,
      顏色: 色票.紙白,
      粗體: true,
      字體: 標題字,
      角色: `統整變化-${index + 1}`,
    });
    加文字(slide, 頁碼, 說明, x + 24, 530, 306, 64, {
      字級: 17,
      顏色: 色票.淺灰,
      字體: 內文字,
      角色: `統整說明-${index + 1}`,
    });
  });
  加講者備註(slide, "統整頁不要重講全部內容，只要把角色、空間、定位三條線收斂給老師。");
}

async function 做結論頁(presentation, data) {
  const 頁碼 = 8;
  const slide = presentation.slides.add();
  await 加圖片(slide, 頁碼, "slide-08.png", 0, 0, W, H, data.標題);
  加形狀(slide, 頁碼, "rect", 0, 0, W, H, "#05121CDD", 色票.透明, 0, "結論遮罩");
  加形狀(slide, 頁碼, "rect", 0, 0, 940, H, "#061A2CF0", 色票.透明, 0, "結論左欄");
  加頁碼框(slide, 頁碼, data.標籤);
  加文字(slide, 頁碼, data.標題, 84, 138, 780, 126, {
    字級: 50,
    顏色: 色票.紙白,
    粗體: true,
    字體: 標題字,
    角色: "結論標題",
  });
  加文字(slide, 頁碼, data.句子, 96, 338, 720, 280, {
    字級: 30,
    顏色: 色票.紙白,
    粗體: true,
    字體: 標題字,
    角色: "結論主句",
  });
  加講者備註(slide, "結論頁要把題目收回來：臺南最值得講的不是它很老，而是它一次次改變自己的方式。");
}

async function 做分工頁(presentation, data) {
  const 頁碼 = 9;
  const slide = presentation.slides.add();
  await 加圖片(slide, 頁碼, "slide-09.png", 0, 0, W, H, data.標題);
  加形狀(slide, 頁碼, "rect", 0, 0, W, H, "#F7FBFFF2", 色票.透明, 0, "分工遮罩");
  加頁碼框(slide, 頁碼, data.標籤);
  加文字(slide, 頁碼, data.標題, 84, 108, 420, 48, {
    字級: 34,
    顏色: 色票.墨黑,
    粗體: true,
    字體: 標題字,
    角色: "分工標題",
  });
  加文字(slide, 頁碼, data.副標題, 84, 172, 760, 40, {
    字級: 18,
    顏色: 色票.深海藍,
    字體: 內文字,
    角色: "分工副標題",
  });

  加圓角卡(slide, 頁碼, 84, 250, 820, 318, "#FFFFFFE8", "#D6E7F4", "分工表格");
  const 欄寬 = [150, 390, 220];
  let x = 112;
  data.表頭.forEach((標頭, index) => {
    加文字(slide, 頁碼, 標頭, x, 276, 欄寬[index] - 12, 20, {
      字級: 14,
      顏色: 色票.深夜藍,
      粗體: true,
      字體: 等寬字,
      角色: `表頭-${index + 1}`,
      檢查: false,
    });
    x += 欄寬[index];
  });
  for (let i = 0; i < 5; i += 1) {
    加形狀(slide, 頁碼, "rect", 112, 312 + i * 50, 764, 1.5, "#D3E4F0", 色票.透明, 0, `表格橫線-${i + 1}`);
  }
  data.列.forEach((row, rowIndex) => {
    let cellX = 112;
    row.forEach((cell, cellIndex) => {
      加文字(slide, 頁碼, cell, cellX, 328 + rowIndex * 50, 欄寬[cellIndex] - 12, 22, {
        字級: 16,
        顏色: 色票.墨黑,
        字體: 內文字,
        角色: `表格內容-${rowIndex + 1}-${cellIndex + 1}`,
        檢查: false,
      });
      cellX += 欄寬[cellIndex];
    });
  });
  加圓角卡(slide, 頁碼, 944, 250, 252, 318, "#0A2135", "#285370", "分工側卡");
  加文字(slide, 頁碼, "使用提醒", 972, 278, 120, 18, {
    字級: 12,
    顏色: 色票.青綠,
    粗體: true,
    字體: 等寬字,
    角色: "分工側卡標題",
    檢查: false,
  });
  加文字(slide, 頁碼, [
    "1. 先把「請填入」改成實際組員姓名。",
    "2. 簡報／上台欄位可填負責頁面。",
    "3. 這頁保持乾淨即可，不需要再加太多圖像。"
  ], 972, 314, 196, 118, {
    字級: 14,
    顏色: 色票.紙白,
    字體: 內文字,
    角色: "分工側卡內容",
  });
  加文字(slide, 頁碼, data.註記, 972, 468, 196, 52, {
    字級: 14,
    顏色: 色票.淺灰,
    字體: 內文字,
    角色: "分工註記",
  });
  加講者備註(slide, "分工頁照實際狀況填入即可，這頁的功能是交代責任，不要再延伸新的內容。");
}

async function 建立簡報() {
  await fs.mkdir(輸出目錄, { recursive: true });
  await fs.mkdir(暫存目錄, { recursive: true });
  await fs.mkdir(預覽目錄, { recursive: true });
  await fs.mkdir(驗證目錄, { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  await 做封面(presentation, 頁面資料[0]);
  await 做總覽(presentation, 頁面資料[1]);
  await 做階段頁(presentation, 3, 頁面資料[2], "slide-03.png");
  await 做階段頁(presentation, 4, 頁面資料[3], "slide-04.png");
  await 做階段頁(presentation, 5, 頁面資料[4], "slide-05.png");
  await 做階段頁(presentation, 6, 頁面資料[5], "slide-06.png");
  await 做統整頁(presentation, 頁面資料[6]);
  await 做結論頁(presentation, 頁面資料[7]);
  await 做分工頁(presentation, 頁面資料[8]);
  return presentation;
}

async function 寫入檢查紀錄(presentation) {
  const lines = [
    JSON.stringify({ kind: "deck", slideCount: presentation.slides.count, slideSize: { width: W, height: H } }),
    ...檢查資料.map((entry) => JSON.stringify(entry)),
  ].join("\n") + "\n";
  await fs.writeFile(檢查紀錄, lines, "utf8");
}

async function 輸出預覽與簡報(presentation) {
  await 寫入檢查紀錄(presentation);
  for (let i = 0; i < presentation.slides.items.length; i += 1) {
    const slide = presentation.slides.items[i];
    const preview = await presentation.export({ slide, format: "png", scale: 1 });
    const bytes = new Uint8Array(await preview.arrayBuffer());
    await fs.writeFile(path.join(預覽目錄, `slide-${String(i + 1).padStart(2, "0")}.png`), bytes);
  }
  const pptx = await PresentationFile.exportPptx(presentation);
  const outputPath = path.join(輸出目錄, "output.pptx");
  await pptx.save(outputPath);
  return outputPath;
}

const presentation = await 建立簡報();
const outputPath = await 輸出預覽與簡報(presentation);
console.log(outputPath);
