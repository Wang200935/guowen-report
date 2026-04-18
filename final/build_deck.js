const pptxgen = require('pptxgenjs');
const path = require('path');

const pptx = new pptxgen();
pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM';
pptx.author = 'OpenAI Codex';
pptx.company = 'guowen-report';
pptx.subject = '臺南城市變遷歷史';
pptx.title = '臺南城市變遷歷史簡報';
pptx.lang = 'zh-TW';
pptx.theme = { headFontFace: 'Aptos Display', bodyFontFace: 'Aptos', lang: 'zh-TW' };

const C = {
  bg: '0B1020',
  panel: '111A2E',
  panel2: '16213B',
  line: '2B5FFF',
  lineSoft: '315BBD',
  cyan: '5BC0EB',
  white: 'F8FAFC',
  text: 'DCE7F5',
  muted: '98A9C7',
  border: '24314F',
};

const assets = {
  p1: path.join(__dirname, '..', 'assets', 'phase1.jpg'),
  p2: path.join(__dirname, '..', 'assets', 'phase2.jpg'),
  p3: path.join(__dirname, '..', 'assets', 'phase3.jpg'),
  p4: path.join(__dirname, '..', 'assets', 'phase4.jpg'),
};

function baseSlide(slide, kicker, title) {
  slide.background = { color: C.bg };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.16, line: { color: C.line }, fill: { color: C.line } });
  slide.addText(kicker, { x: 0.75, y: 0.42, w: 3.2, h: 0.2, fontSize: 11, bold: true, color: C.cyan, margin: 0, charSpace: 1.5 });
  slide.addText(title, { x: 0.75, y: 0.72, w: 6.35, h: 0.7, fontFace: 'Aptos Display', fontSize: 25, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.72, y: 1.58, w: 5.95, h: 4.95,
    rectRadius: 0.05,
    line: { color: C.border, pt: 1 },
    fill: { color: C.panel }
  });
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.0, y: 1.58, w: 5.58, h: 4.95,
    rectRadius: 0.05,
    line: { color: C.border, pt: 1 },
    fill: { color: C.panel2 }
  });
}

function addCore(slide, text) {
  slide.addText('CORE MESSAGE', { x: 0.98, y: 1.9, w: 1.5, h: 0.18, fontSize: 10, bold: true, color: C.cyan, margin: 0, charSpace: 1.2 });
  slide.addText(text, { x: 0.98, y: 2.12, w: 5.2, h: 0.52, fontSize: 18, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  slide.addShape(pptx.ShapeType.line, { x: 0.98, y: 2.78, w: 4.95, h: 0, line: { color: C.lineSoft, pt: 1.2 } });
}

function addBullets(slide, lines) {
  const runs = [];
  lines.forEach((line, idx) => {
    runs.push({ text: line, options: { bullet: { indent: 14 }, breakLine: idx < lines.length - 1, hanging: 4, color: C.text, fontSize: 14 } });
  });
  slide.addText(runs, { x: 1.02, y: 3.05, w: 4.95, h: 2.05, margin: 0.02, paraSpaceAfterPt: 7, valign: 'top' });
}

function addTakeaway(slide, text) {
  slide.addText('TAKEAWAY', { x: 0.98, y: 5.45, w: 1.1, h: 0.18, fontSize: 10, bold: true, color: C.cyan, margin: 0, charSpace: 1.2 });
  slide.addText(text, { x: 0.98, y: 5.66, w: 5.15, h: 0.4, fontSize: 15, bold: true, color: C.white, margin: 0, fit: 'shrink' });
}

function addImagePanel(slide, imagePath, caption) {
  slide.addImage({ path: imagePath, x: 7.24, y: 1.82, w: 5.12, h: 3.85, sizing: { type: 'cover', x: 7.24, y: 1.82, w: 5.12, h: 3.85 } });
  slide.addShape(pptx.ShapeType.rect, { x: 7.24, y: 1.82, w: 5.12, h: 3.85, line: { color: C.panel2, transparency: 100 }, fill: { color: C.bg, transparency: 18 } });
  slide.addText('VISUAL SIGNAL', { x: 7.34, y: 5.93, w: 1.25, h: 0.18, fontSize: 10, bold: true, color: C.cyan, margin: 0, charSpace: 1.2 });
  slide.addText(caption, { x: 7.34, y: 6.16, w: 4.95, h: 0.28, fontSize: 10, color: C.muted, italic: true, margin: 0, fit: 'shrink' });
}

function addPhaseSlide({ kicker, title, core, bullets, takeaway, image, caption }) {
  const s = pptx.addSlide();
  baseSlide(s, kicker, title);
  addCore(s, core);
  addBullets(s, bullets);
  addTakeaway(s, takeaway);
  addImagePanel(s, image, caption);
}

// Cover
{
  const s = pptx.addSlide();
  s.background = { color: C.bg };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.18, line: { color: C.line }, fill: { color: C.line } });
  s.addText('CITY TRANSFORMATION REPORT', { x: 0.82, y: 0.58, w: 3.8, h: 0.2, fontSize: 11, bold: true, color: C.cyan, margin: 0, charSpace: 1.6 });
  s.addText('從港口據點到大臺南', { x: 0.82, y: 1.0, w: 5.7, h: 0.62, fontFace: 'Aptos Display', fontSize: 28, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  s.addText('臺南的城市變遷歷史', { x: 0.85, y: 1.72, w: 4.5, h: 0.25, fontSize: 18, bold: true, color: C.text, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.82, y: 2.3, w: 5.25, h: 1.55, rectRadius: 0.05, line: { color: C.border, pt: 1 }, fill: { color: C.panel } });
  s.addText('這份簡報不是介紹景點，而是用四個轉折點，說清楚臺南如何在不同時代改變自己的角色、空間與定位。', { x: 1.05, y: 2.62, w: 4.7, h: 0.9, fontSize: 18, color: C.text, margin: 0, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.82, y: 6.05, w: 3.15, h: 0.45, rectRadius: 0.04, line: { color: C.border, pt: 1 }, fill: { color: C.panel2 } });
  s.addText('國文分組報告｜班級／組別／姓名', { x: 1.05, y: 6.19, w: 2.7, h: 0.16, fontSize: 11, color: C.white, margin: 0, align: 'center' });
  s.addImage({ path: assets.p4, x: 7.15, y: 0.78, w: 5.5, h: 5.9, sizing: { type: 'cover', x: 7.15, y: 0.78, w: 5.5, h: 5.9 } });
  s.addShape(pptx.ShapeType.rect, { x: 7.15, y: 0.78, w: 5.5, h: 5.9, line: { color: C.bg, transparency: 100 }, fill: { color: C.bg, transparency: 36 } });
}

// Overview
{
  const s = pptx.addSlide();
  baseSlide(s, 'NARRATIVE OVERVIEW', '為什麼用臺南，看得出城市怎麼變？');
  addCore(s, '因為臺南不是只累積歷史，而是在不同時代裡，一次次改變自己的城市角色。');
  const stages = [
    ['01', '17世紀開港建城期', '先因海港、貿易與統治而成城'],
    ['02', '清領府城成形期', '城市格局穩定，成為全臺首府'],
    ['03', '19世紀後期開港與日治近代化期', '從傳統府城轉向近代都市'],
    ['04', '戰後重組與當代再定位期', '從舊首府走向今天的大臺南'],
  ];
  stages.forEach((it, i) => {
    const x = 1.0 + (i % 2) * 2.5;
    const y = 3.05 + Math.floor(i / 2) * 1.22;
    s.addShape(pptx.ShapeType.roundRect, { x, y, w: 2.22, h: 0.95, rectRadius: 0.04, line: { color: C.border, pt: 1 }, fill: { color: C.panel2 } });
    s.addText(it[0], { x: x + 0.18, y: y + 0.14, w: 0.3, h: 0.18, fontSize: 11, bold: true, color: C.cyan, margin: 0 });
    s.addText(it[1], { x: x + 0.55, y: y + 0.12, w: 1.45, h: 0.22, fontSize: 12, bold: true, color: C.white, margin: 0, fit: 'shrink' });
    s.addText(it[2], { x: x + 0.18, y: y + 0.42, w: 1.8, h: 0.24, fontSize: 10, color: C.muted, margin: 0, fit: 'shrink' });
  });
  addTakeaway(s, '這份報告要回答的，不是臺南有多老，而是臺南怎麼一直在變。');
  addImagePanel(s, assets.p2, '歷史地圖的作用，是先讓觀眾建立「四段變遷」的空間感。');
}

addPhaseSlide({
  kicker: 'PHASE 01',
  title: '臺南的起點，不是古都，而是海港與政權中心',
  core: '臺南最早的重要性，來自港口、貿易與統治需求，而不是先有文化古都形象。',
  bullets: [
    '1624 年荷蘭在大員建立據點，臺南開始進入對外貿易網絡',
    '熱蘭遮城與普羅民遮城周邊形成早期政商核心',
    '明鄭接手後，延續臺南作為統治中心的角色',
    '這時的臺南，已經具備港口城市與政權核心的雙重性格'
  ],
  takeaway: '臺南是一座先因海而興、再因政權而穩的城市。',
  image: assets.p1,
  caption: '海圖與堡壘位置，比觀光照更能證明「臺南先因海而成城」。'
});

addPhaseSlide({
  kicker: 'PHASE 02',
  title: '到了清代，臺南真正有了「府城」的樣子',
  core: '清領時期的臺南，不只是熱鬧港口，而是成為全臺政治、商業與文化中心。',
  bullets: [
    '1684 年設臺灣府後，臺南成為全臺行政核心',
    '城牆、街區、官署、廟宇與生活圈逐步穩定',
    '城市空間從點狀據點，變成有清楚秩序的府城格局',
    '今日大家對臺南「府城」「古都」的印象，就是這個階段留下的結果'
  ],
  takeaway: '這一段的關鍵不是清朝統治，而是臺南完成了首府城市的定型。',
  image: assets.p2,
  caption: '府城圖的價值，在於讓老師一眼看到「城市格局已經穩定下來」。'
});

addPhaseSlide({
  kicker: 'PHASE 03',
  title: '失去全臺首位後，臺南反而開始變現代',
  core: '臺南雖然不再是全臺唯一中心，但正因為這樣，它在近代化中換了一種新的城市樣貌繼續存在。',
  bullets: [
    '19 世紀後期後，港口與政治地位開始變化',
    '日治時期拆城、修路、引進公共建築與新式都市設施',
    '城市外貌從封閉府城，變成更開放的近代都市',
    '這一段最大的重點，是城市功能與空間秩序一起被重新設計'
  ],
  takeaway: '臺南不是在這一段衰落，而是在這一段完成近代轉型。',
  image: assets.p3,
  caption: '真正的重點不是日治建築很美，而是臺南的城市樣子被重做了一次。'
});

addPhaseSlide({
  kicker: 'PHASE 04',
  title: '戰後的臺南，不再只是舊府城，而是今天的大臺南',
  core: '戰後的臺南雖然不再是全臺政治中心，但它透過文化保存、都市擴張與行政整合，重新建立新的城市定位。',
  bullets: [
    '戰後臺南的城市功能持續擴張',
    '2010 年縣市合併，城市尺度被重新放大',
    '老城保存與新區發展同時進行',
    '今天的臺南不只是古都，而是一座文化與現代生活並存的城市'
  ],
  takeaway: '今天的臺南，最大的特徵不是停在過去，而是把歷史和現代疊在一起。',
  image: assets.p4,
  caption: '這頁不是城市宣傳，而是告訴老師：臺南今天的定位已經和過去不同。'
});

// synthesis
{
  const s = pptx.addSlide();
  baseSlide(s, 'SYNTHESIS', '如果只記三件事，臺南到底變了什麼？');
  addCore(s, '看完四個階段之後，臺南的變化其實可以濃縮成角色、空間、定位三條線。');
  const items = [
    ['城市角色變了', '從海港據點 → 府城 → 近代都市 → 大臺南'],
    ['城市空間變了', '從城內核心 → 有邊界的府城 → 開放型都市 → 更大的都會生活圈'],
    ['城市定位變了', '從政治中心 → 文化古都 → 文化與現代並存的城市'],
  ];
  items.forEach((it, i) => {
    s.addShape(pptx.ShapeType.roundRect, { x: 1.0, y: 3.02 + i * 0.88, w: 5.1, h: 0.68, rectRadius: 0.04, line: { color: C.border, pt: 1 }, fill: { color: C.panel2 } });
    s.addText(it[0], { x: 1.18, y: 3.16 + i * 0.88, w: 1.4, h: 0.18, fontSize: 13, bold: true, color: C.white, margin: 0 });
    s.addText(it[1], { x: 2.72, y: 3.16 + i * 0.88, w: 3.1, h: 0.18, fontSize: 11, color: C.text, margin: 0, fit: 'shrink' });
  });
  addTakeaway(s, '臺南最特別的地方，不只是它保留歷史，而是它一直透過變化活到今天。');
  addImagePanel(s, assets.p4, '統整頁也維持同一種右圖結構，整套簡報才不會忽黑忽白。');
}

// conclusion
{
  const s = pptx.addSlide();
  baseSlide(s, 'FINAL TAKEAWAY', '臺南最特別的不是很老，而是一直在變');
  addCore(s, '臺南的歷史不是停在過去，而是在不同時代裡不斷重新找到自己的位置。');
  s.addText('從海港據點到府城，從傳統城市到近代都市，再到今天的大臺南，臺南每次失去舊角色後，都會長出新的城市樣子。', { x: 1.02, y: 3.05, w: 5.08, h: 0.78, fontSize: 17, color: C.text, margin: 0, fit: 'shrink' });
  s.addText('這就是我們認為臺南最符合「城市變遷歷史」這個題目的原因。', { x: 1.02, y: 4.45, w: 5.1, h: 0.38, fontSize: 16, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  addTakeaway(s, '重點不是臺南很老，而是它一直都在變。');
  addImagePanel(s, assets.p3, '結尾頁也維持同一系統：左邊結論、右邊視覺，整套才像一家公司做的。');
}

// roles
{
  const s = pptx.addSlide();
  baseSlide(s, 'EXECUTION', '組員分工');
  addCore(s, '分工頁也必須沿用同一套視覺系統，不能突然變成另外一種作業風。');
  s.addShape(pptx.ShapeType.roundRect, { x: 1.0, y: 3.0, w: 5.1, h: 2.35, rectRadius: 0.04, line: { color: C.border, pt: 1 }, fill: { color: C.panel2 } });
  s.addText('姓名', { x: 1.18, y: 3.22, w: 0.7, h: 0.18, fontSize: 13, bold: true, color: C.white, margin: 0 });
  s.addText('負責內容', { x: 2.55, y: 3.22, w: 1.2, h: 0.18, fontSize: 13, bold: true, color: C.white, margin: 0 });
  s.addText('簡報 / 上台部分', { x: 4.55, y: 3.22, w: 1.35, h: 0.18, fontSize: 13, bold: true, color: C.white, margin: 0 });
  [3.55, 4.05, 4.55, 5.05].forEach(y => s.addShape(pptx.ShapeType.line,{x:1.12,y,w:4.84,h:0,line:{color:C.border,pt:1}}));
  s.addText('請填入組員姓名與分工', { x: 1.18, y: 5.72, w: 3.2, h: 0.18, fontSize: 11, color: C.muted, margin: 0 });
  addImagePanel(s, assets.p1, '分工頁右側仍保留主視覺卡片，整套版面才完全一致。');
  addTakeaway(s, '分工頁只要清楚，不要花；正式內容的重點應留在前八頁。');
}

pptx.writeFile({ fileName: path.join(__dirname, '臺南城市變遷歷史簡報.pptx') });
