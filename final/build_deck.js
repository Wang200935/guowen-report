const pptxgen = require('pptxgenjs');
const path = require('path');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI Codex';
pptx.company = 'guowen-report';
pptx.subject = '臺南城市變遷歷史';
pptx.title = '臺南城市變遷歷史簡報';
pptx.lang = 'zh-TW';
pptx.theme = {
  headFontFace: 'Aptos Display',
  bodyFontFace: 'Aptos',
  lang: 'zh-TW'
};
pptx.defineLayout({ name: 'CUSTOM', width: 13.333, height: 7.5 });
pptx.layout = 'CUSTOM';

const C = {
  navy: '0B1020',
  blue: '2563EB',
  blue2: '1D4ED8',
  cyan: '38BDF8',
  ink: '0F172A',
  slate: '334155',
  muted: '64748B',
  paper: 'F8FAFC',
  white: 'FFFFFF',
  line: 'D6DEE8',
  lightBlue: 'E8F0FF',
  soft: 'EEF4FF',
  red: 'B91C1C'
};

const assets = {
  p1: path.join(__dirname, '..', 'assets', 'phase1.jpg'),
  p2: path.join(__dirname, '..', 'assets', 'phase2.jpg'),
  p3: path.join(__dirname, '..', 'assets', 'phase3.jpg'),
  p4: path.join(__dirname, '..', 'assets', 'phase4.jpg'),
};

function addTopBar(slide) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.18, line: { color: C.blue }, fill: { color: C.blue } });
}

function addTitle(slide, title, kicker) {
  if (kicker) {
    slide.addText(kicker, { x: 0.7, y: 0.45, w: 5.5, h: 0.25, fontFace: 'Aptos', fontSize: 11, bold: true, color: C.blue, margin: 0 });
  }
  slide.addText(title, { x: 0.7, y: 0.72, w: 7.3, h: 0.65, fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.ink, margin: 0 });
}

function addCoreMessage(slide, text) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.7, y: 1.55, w: 5.8, h: 0.78,
    rectRadius: 0.08,
    line: { color: C.blue, transparency: 100 },
    fill: { color: C.soft }
  });
  slide.addText([
    { text: 'CORE MESSAGE  ', options: { bold: true, color: C.blue, fontSize: 11 } },
    { text, options: { color: C.ink, fontSize: 16, bold: true } }
  ], { x: 0.92, y: 1.78, w: 5.3, h: 0.28, margin: 0, fit: 'shrink' });
}

function addBullets(slide, lines) {
  const runs = [];
  lines.forEach((line, i) => {
    runs.push({ text: line, options: { bullet: { indent: 14 }, breakLine: i < lines.length - 1, hanging: 3, color: C.slate, fontSize: 15 } });
  });
  slide.addText(runs, { x: 0.92, y: 2.65, w: 5.2, h: 2.45, margin: 0.02, paraSpaceAfterPt: 8, valign: 'top' });
}

function addConclusion(slide, text) {
  slide.addShape(pptx.ShapeType.line, { x: 0.92, y: 5.35, w: 4.9, h: 0, line: { color: C.line, pt: 1.2 } });
  slide.addText('結論', { x: 0.92, y: 5.48, w: 0.8, h: 0.2, fontSize: 11, bold: true, color: C.blue, margin: 0 });
  slide.addText(text, { x: 0.92, y: 5.7, w: 5.3, h: 0.5, fontSize: 15, color: C.ink, bold: true, margin: 0 });
}

function addRightImage(slide, imgPath, caption) {
  slide.addShape(pptx.ShapeType.roundRect, { x: 7.3, y: 1.55, w: 5.28, h: 4.7, rectRadius: 0.05, line: { color: C.line, pt: 1 }, fill: { color: C.white }, shadow: { type: 'outer', color: 'AAB4C0', blur: 1, angle: 45, distance: 1, opacity: 0.15 } });
  slide.addImage({ path: imgPath, x: 7.45, y: 1.7, w: 4.98, h: 3.72, sizing: { type: 'cover', x: 7.45, y: 1.7, w: 4.98, h: 3.72 } });
  slide.addText(caption, { x: 7.55, y: 5.62, w: 4.7, h: 0.36, fontSize: 10, color: C.muted, italic: true, margin: 0, fit: 'shrink' });
}

function makeContentSlide({ kicker, title, core, bullets, conclusion, image, caption }) {
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addTopBar(slide);
  addTitle(slide, title, kicker);
  addCoreMessage(slide, core);
  addBullets(slide, bullets);
  addConclusion(slide, conclusion);
  addRightImage(slide, image, caption);
  return slide;
}

// Slide 1
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.22, line: { color: C.cyan }, fill: { color: C.cyan } });
  s.addText('FROM PORT TO METRO', { x: 0.82, y: 0.78, w: 4.2, h: 0.25, fontSize: 13, bold: true, color: C.cyan, charSpace: 1, margin: 0 });
  s.addText('從港口據點到大臺南', { x: 0.8, y: 1.15, w: 6.6, h: 0.7, fontFace: 'Aptos Display', fontSize: 28, bold: true, color: C.white, margin: 0 });
  s.addText('臺南的城市變遷歷史', { x: 0.82, y: 1.92, w: 5.5, h: 0.35, fontSize: 20, color: 'DCE7FF', bold: true, margin: 0 });
  s.addText('不是介紹景點，而是用四個轉折點，解釋臺南這座城市如何在不同時代改變自己的角色、空間與定位。', { x: 0.82, y: 2.55, w: 5.85, h: 0.9, fontSize: 18, color: 'D6E3F8', margin: 0, breakLine: false, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect, { x: 0.82, y: 5.85, w: 3.0, h: 0.55, rectRadius: 0.04, line: { color: C.cyan, transparency: 100 }, fill: { color: '13203B' } });
  s.addText('國文分組報告｜班級／組別／姓名', { x: 1.0, y: 6.02, w: 2.7, h: 0.18, fontSize: 11, color: C.white, margin: 0, align: 'center' });
  s.addImage({ path: assets.p4, x: 7.15, y: 0.72, w: 5.55, h: 5.85, sizing: { type: 'cover', x: 7.15, y: 0.72, w: 5.55, h: 5.85 } });
  s.addShape(pptx.ShapeType.rect, { x: 7.15, y: 0.72, w: 5.55, h: 5.85, line: { color: C.navy, transparency: 100 }, fill: { color: C.navy, transparency: 42 } });
  s.addText('Tainan skyline as today’s endpoint', { x: 7.35, y: 6.15, w: 3.8, h: 0.16, fontSize: 10, color: 'D6E3F8', italic: true, margin: 0 });
}

// Slide 2
{
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  addTopBar(s);
  addTitle(s, '為什麼用臺南，看得出城市怎麼變？', 'NARRATIVE OVERVIEW');
  addCoreMessage(s, '臺南不是只累積歷史，而是在不同時代裡，一次次改變自己的城市角色。');
  const cards = [
    ['01', '17世紀開港建城期', '先因海港、貿易與統治而成城'],
    ['02', '清領府城成形期', '城市格局穩定，成為全臺首府'],
    ['03', '19世紀後期開港與日治近代化期', '從傳統府城轉向近代都市'],
    ['04', '戰後重組與當代再定位期', '從舊首府走向今天的大臺南'],
  ];
  let y=2.7;
  cards.forEach(([num,title,desc], idx)=>{
    s.addShape(pptx.ShapeType.roundRect,{x:0.85+(idx%2)*3.15,y:y+Math.floor(idx/2)*1.35,w:2.85,h:1.05,rectRadius:0.04,line:{color:C.line,pt:1},fill:{color:C.white}});
    s.addText(num,{x:1.05+(idx%2)*3.15,y:y+0.16+Math.floor(idx/2)*1.35,w:0.45,h:0.2,fontSize:12,bold:true,color:C.blue,margin:0});
    s.addText(title,{x:1.5+(idx%2)*3.15,y:y+0.14+Math.floor(idx/2)*1.35,w:1.95,h:0.28,fontSize:13,bold:true,color:C.ink,margin:0,fit:'shrink'});
    s.addText(desc,{x:1.05+(idx%2)*3.15,y:y+0.48+Math.floor(idx/2)*1.35,w:2.35,h:0.35,fontSize:11,color:C.muted,margin:0,fit:'shrink'});
  });
  s.addText('這份報告要回答的，不是臺南有多老，而是臺南怎麼一直在變。', { x: 0.92, y: 5.85, w: 6.05, h: 0.32, fontSize: 15, bold: true, color: C.ink, margin: 0 });
  s.addImage({ path: assets.p2, x: 7.55, y: 2.0, w: 4.9, h: 3.6, sizing: { type: 'cover', x: 7.55, y: 2.0, w: 4.9, h: 3.6 } });
  s.addText('用歷史地圖先建立「四段變遷」的空間感', { x: 7.65, y: 5.75, w: 4.4, h: 0.22, fontSize: 10, italic: true, color: C.muted, margin: 0 });
}

makeContentSlide({
  kicker: 'PHASE 01',
  title: '臺南的起點，不是古都，而是海港與政權中心',
  core: '臺南最早的重要性，來自港口、貿易與統治需求，而不是先有文化古都形象。',
  bullets: [
    '1624 年荷蘭在大員建立據點，臺南開始進入對外貿易網絡',
    '熱蘭遮城與普羅民遮城周邊形成早期政商核心',
    '明鄭接手後，延續臺南作為統治中心的角色',
    '這時的臺南，已經具備港口城市與政權核心的雙重性格'
  ],
  conclusion: '臺南是一座先因海而興、再因政權而穩的城市。',
  image: assets.p1,
  caption: '海圖與堡壘位置，比觀光照更能證明「臺南先因海而成城」。'
});

makeContentSlide({
  kicker: 'PHASE 02',
  title: '到了清代，臺南真正有了「府城」的樣子',
  core: '清領時期的臺南，不只是熱鬧港口，而是成為全臺政治、商業與文化中心。',
  bullets: [
    '1684 年設臺灣府後，臺南成為全臺行政核心',
    '城牆、街區、官署、廟宇與生活圈逐步穩定',
    '城市空間從點狀據點，變成有清楚秩序的府城格局',
    '今日大家對臺南「府城」「古都」的印象，就是這個階段留下的結果'
  ],
  conclusion: '這一段的關鍵不是清朝統治，而是臺南完成了首府城市的定型。',
  image: assets.p2,
  caption: '府城圖的價值，在於讓老師一眼看到「城市格局已經穩定下來」。'
});

makeContentSlide({
  kicker: 'PHASE 03',
  title: '失去全臺首位後，臺南反而開始變現代',
  core: '臺南雖然不再是全臺唯一中心，但正因為這樣，它在近代化中換了一種新的城市樣貌繼續存在。',
  bullets: [
    '19 世紀後期後，港口與政治地位開始變化',
    '日治時期拆城、修路、引進公共建築與新式都市設施',
    '城市外貌從封閉府城，變成更開放的近代都市',
    '這一段最大的重點，是城市功能與空間秩序一起被重新設計'
  ],
  conclusion: '臺南不是在這一段衰落，而是在這一段完成近代轉型。',
  image: assets.p3,
  caption: '前後對照是最有 NVIDIA 感的做法：畫面直接證明「城市樣子變了」。'
});

makeContentSlide({
  kicker: 'PHASE 04',
  title: '戰後的臺南，不再只是舊府城，而是今天的大臺南',
  core: '戰後的臺南雖然不再是全臺政治中心，但它透過文化保存、都市擴張與行政整合，重新建立新的城市定位。',
  bullets: [
    '戰後臺南的城市功能持續擴張',
    '2010 年縣市合併，城市尺度被重新放大',
    '老城保存與新區發展同時進行',
    '今天的臺南不只是古都，而是一座文化與現代生活並存的城市'
  ],
  conclusion: '今天的臺南，最大的特徵不是停在過去，而是把歷史和現代疊在一起。',
  image: assets.p4,
  caption: '收尾頁的圖要講「今天的臺南」，不是講「臺南很漂亮」。'
});

// Slide 7 synthesis
{
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  addTopBar(s);
  addTitle(s, '如果只記三件事，臺南到底變了什麼？', 'SYNTHESIS');
  const cards = [
    ['城市角色變了','從海港據點 → 府城 → 近代都市 → 大臺南'],
    ['城市空間變了','從城內核心 → 有邊界的府城 → 開放型都市 → 更大的都會生活圈'],
    ['城市定位變了','從政治中心 → 文化古都 → 文化與現代並存的城市']
  ];
  cards.forEach((c, i) => {
    s.addShape(pptx.ShapeType.roundRect,{x:0.92,y:1.95+i*1.35,w:5.8,h:1.02,rectRadius:0.04,line:{color:C.line,pt:1},fill:{color:C.white}});
    s.addText(c[0],{x:1.15,y:2.12+i*1.35,w:2.2,h:0.2,fontSize:16,bold:true,color:C.ink,margin:0});
    s.addText(c[1],{x:1.15,y:2.42+i*1.35,w:5.0,h:0.22,fontSize:12,color:C.muted,margin:0,fit:'shrink'});
  });
  s.addImage({ path: assets.p4, x: 7.6, y: 2.05, w: 4.4, h: 3.5, sizing: { type: 'cover', x: 7.6, y: 2.05, w: 4.4, h: 3.5 } });
  s.addText('臺南最特別的地方，不只是它保留歷史，而是它一直透過變化活到今天。', { x: 0.95, y: 6.22, w: 11.1, h: 0.28, fontSize: 15, bold: true, color: C.ink, margin: 0 });
}

// Slide 8 conclusion
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addShape(pptx.ShapeType.rect,{x:0,y:0,w:13.333,h:0.2,line:{color:C.cyan},fill:{color:C.cyan}});
  s.addText('FINAL TAKEAWAY', { x: 0.82, y: 0.62, w: 2.8, h: 0.2, fontSize: 12, bold: true, color: C.cyan, margin: 0, charSpace: 1.5 });
  s.addText(`臺南最特別的不是很老，
而是一直在變`, { x: 0.82, y: 1.1, w: 5.9, h: 1.2, fontFace: 'Aptos Display', fontSize: 24, bold: true, color: C.white, margin: 0 });
  s.addText(`從海港據點到府城，從傳統城市到近代都市，再到今天的大臺南，
臺南的歷史不是停在過去，而是在不同時代裡不斷重新找到自己的位置。`, { x: 0.85, y: 2.65, w: 5.8, h: 1.0, fontSize: 17, color: 'D6E3F8', margin: 0, fit: 'shrink' });
  s.addShape(pptx.ShapeType.roundRect,{x:0.85,y:4.3,w:5.7,h:0.82,rectRadius:0.05,line:{color:C.cyan,transparency:100},fill:{color:'13203B'}});
  s.addText('這就是我們認為臺南最符合「城市變遷歷史」這個題目的原因。', { x: 1.0, y: 4.58, w: 5.35, h: 0.22, fontSize: 15, bold: true, color: C.white, margin: 0, fit: 'shrink' });
  s.addImage({ path: assets.p3, x: 7.3, y: 0.85, w: 5.15, h: 5.5, sizing: { type: 'cover', x: 7.3, y: 0.85, w: 5.15, h: 5.5 } });
  s.addShape(pptx.ShapeType.rect,{x:7.3,y:0.85,w:5.15,h:5.5,line:{color:C.navy,transparency:100},fill:{color:C.navy,transparency:48}});
}

// Slide 9 roles
{
  const s = pptx.addSlide();
  s.background = { color: C.paper };
  addTopBar(s);
  addTitle(s, '組員分工', 'EXECUTION');
  s.addShape(pptx.ShapeType.roundRect,{x:0.92,y:1.75,w:5.8,h:3.8,rectRadius:0.04,line:{color:C.line,pt:1},fill:{color:C.white}});
  s.addText('姓名', { x: 1.2, y: 2.02, w: 0.8, h: 0.2, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  s.addText('負責內容', { x: 2.8, y: 2.02, w: 1.4, h: 0.2, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  s.addText('簡報 / 上台部分', { x: 5.0, y: 2.02, w: 1.5, h: 0.2, fontSize: 14, bold: true, color: C.ink, margin: 0 });
  [2.35, 2.95, 3.55, 4.15, 4.75].forEach(y => s.addShape(pptx.ShapeType.line,{x:1.1,y,w:5.35,h:0,line:{color:C.line,pt:1}}));
  s.addText('請填入組員姓名與分工', { x: 1.2, y: 5.95, w: 3.2, h: 0.22, fontSize: 12, color: C.muted, margin: 0 });
  s.addShape(pptx.ShapeType.roundRect,{x:7.55,y:2.1,w:4.45,h:2.45,rectRadius:0.05,line:{color:C.blue,transparency:100},fill:{color:C.soft}});
  s.addText('最後提醒', { x: 7.85, y: 2.35, w: 1.3, h: 0.2, fontSize: 13, bold: true, color: C.blue, margin: 0 });
  s.addText([
    { text: '• 分工頁只要清楚，不要花\n', options: { breakLine: true, color: C.slate, fontSize: 13 } },
    { text: '• 報告時口頭只帶一句「最後這是我們的分工」\n', options: { breakLine: true, color: C.slate, fontSize: 13 } },
    { text: '• 正式內容的重點應留在前八頁', options: { color: C.slate, fontSize: 13 } },
  ], { x: 7.85, y: 2.72, w: 3.8, h: 1.2, margin: 0.02 });
  s.addText('謝謝大家', { x: 9.15, y: 5.85, w: 1.7, h: 0.28, fontSize: 18, bold: true, color: C.ink, margin: 0, align: 'center' });
}

pptx.writeFile({ fileName: path.join(__dirname, '臺南城市變遷歷史簡報.pptx') });
