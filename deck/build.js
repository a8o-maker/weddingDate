// 세로형 인포그래픽 청첩장 (PPTX) 생성 스크립트
// 실행: node deck/build.js  → deck/wedding-infographic.pptx
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL_DIR + "/scripts/apply_theme.js");

const OUT = path.join(__dirname, "wedding-infographic.pptx");
const W = 7.5, H = 13.333, M = 0.6, CW = W - M * 2;

const THEME = {
  name: "GyuSoo SeHwa Sky to Sea",
  headFontFace: "Malgun Gothic",
  bodyFontFace: "Malgun Gothic",
  colors: {
    dk1: "030712", lt1: "FFFFFF", dk2: "0F172A", lt2: "CBD5E1",
    accent1: "2DD4BF", accent2: "A5B4FC", accent3: "0D9488",
    accent4: "99F6E4", accent5: "FDE68A", accent6: "FCA5A5",
    hlink: "5EEAD4", folHlink: "A5B4FC",
  },
};

// ── helpers ─────────────────────────────────────────────
async function icon(Comp, hex, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + hex, size: px })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// 배경: 하늘(위) → 바다(아래)로 슬라이드마다 조금씩 깊어지는 그라데이션
async function background(top, bottom, { glow, bubbles, extra = "" } = {}) {
  const s = 150; // px per inch
  const w = W * s, h = H * s;
  let bub = "";
  if (bubbles) {
    const pts = [[0.9, 11.8, 0.10], [1.5, 12.6, 0.06], [6.4, 12.2, 0.12], [6.8, 10.9, 0.05], [5.9, 12.9, 0.07], [0.6, 9.8, 0.05]];
    bub = pts.map(([x, y, r]) =>
      `<circle cx="${x * s}" cy="${y * s}" r="${r * s}" fill="none" stroke="#ffffff" stroke-opacity="0.25" stroke-width="3"/>`).join("");
  }
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
    <defs>
      <linearGradient id="g" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0" stop-color="#${top}"/><stop offset="1" stop-color="#${bottom}"/>
      </linearGradient>
      <radialGradient id="r" cx="0.5" cy="0" r="0.7">
        <stop offset="0" stop-color="#${glow || top}" stop-opacity="${glow ? 0.35 : 0}"/>
        <stop offset="1" stop-color="#${top}" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="100%" height="100%" fill="url(#g)"/>
    <rect width="100%" height="100%" fill="url(#r)"/>
    ${bub}${extra}
  </svg>`;
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return { data: "image/png;base64," + buf.toString("base64") };
}

(async () => {
  const pres = new pptxgen();
  pres.defineLayout({ name: "PORTRAIT_9x16", width: W, height: H });
  pres.layout = "PORTRAIT_9x16";
  pres.title = "손규수 & 문세화 결혼합니다";
  pres.author = "GyuSoo & SeHwa";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  // ── layouts ──
  pres.defineSlideMaster({
    title: "COVER",
    background: { color: THEME.colors.dk1 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: M, y: 1.7, w: CW, h: 0.9, fontSize: 38, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: THEME.colors.dk2 },
    objects: [
      { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.8, w: CW, h: 0.35, fontSize: 12, color: C.accent1, align: "center", valign: "middle", charSpacing: 6, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: M, y: 1.15, w: CW, h: 0.8, fontSize: 30, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 }, text: "" } },
    ],
  });

  // ── common styles (fresh object every call) ──
  const glass = (extra = {}) => ({
    fill: { color: "FFFFFF", transparency: 92 },
    line: { color: "FFFFFF", transparency: 82, width: 0.75 },
    rectRadius: 0.22, ...extra,
  });
  const txt = (o) => ({ isTextBox: true, margin: 0, ...o });
  const heading = (s, eyebrow, title) => {
    s.addText(eyebrow, { placeholder: "eyebrow" });
    s.addText(title, { placeholder: "title" });
  };

  const I = {
    plane: await icon(fa.FaPlane, "C7D2FE"),
    heart: await icon(fa.FaHeart, THEME.colors.accent6),
    users: await icon(fa.FaUserFriends, THEME.colors.accent4),
    inf: await icon(fa.FaInfinity, THEME.colors.accent4),
    clock: await icon(fa.FaClock, THEME.colors.accent4),
    video: await icon(fa.FaVideo, THEME.colors.accent4),
    home: await icon(fa.FaHome, THEME.colors.accent4),
    cogs: await icon(fa.FaCogs, THEME.colors.accent4),
    planeTeal: await icon(fa.FaPlane, THEME.colors.dk1),
    globe: await icon(fa.FaGlobeAsia, THEME.colors.dk1),
    water: await icon(fa.FaWater, THEME.colors.dk1),
    ring: await icon(fa.FaRing, THEME.colors.dk1),
    temp: await icon(fa.FaThermometerHalf, THEME.colors.accent6),
    pin: await icon(fa.FaMapMarkerAlt, THEME.colors.accent1),
    wallet: await icon(fa.FaWallet, THEME.colors.accent4),
  };

  // ════════ 1. COVER ════════
  pres.addSection({ title: "표지" });
  {
    // 샌프란시스코 → 부산 항로 (배경에 점선 아치)
    const s = 150;
    const arc = `<path d="M ${1.1 * s} ${6.4 * s} Q ${3.75 * s} ${3.8 * s} ${6.4 * s} ${6.4 * s}" fill="none" stroke="#A5B4FC" stroke-opacity="0.8" stroke-width="4" stroke-dasharray="14 14"/>`;
    const slide = pres.addSlide({ masterName: "COVER", sectionTitle: "표지" });
    slide.background = await background("030712", "1E293B", { glow: "6366F1", extra: arc });

    slide.addText("SAN FRANCISCO  ✈  BUSAN", txt({ objectName: "Route eyebrow", x: M, y: 1.15, w: CW, h: 0.35, fontSize: 12, color: C.accent2, align: "center", charSpacing: 6 }));
    slide.addText("손규수 & 문세화", { placeholder: "title" });
    slide.addText("태평양을 건너 이어진 특별한 인연", txt({ objectName: "Tagline", x: M, y: 2.7, w: CW, h: 0.4, fontSize: 15, color: C.background2, align: "center" }));

    // 항로 끝점과 비행기
    slide.addShape(pres.shapes.OVAL, { objectName: "SF dot", x: 1.1 - 0.11, y: 6.4 - 0.11, w: 0.22, h: 0.22, fill: { color: C.accent2 }, line: { type: "none" } });
    slide.addShape(pres.shapes.OVAL, { objectName: "Busan dot", x: 6.4 - 0.11, y: 6.4 - 0.11, w: 0.22, h: 0.22, fill: { color: C.accent1 }, line: { type: "none" } });
    slide.addImage({ objectName: "Plane", data: I.plane, x: 3.75 - 0.3, y: 5.1 - 0.3, w: 0.6, h: 0.6 });
    slide.addText("SAN FRANCISCO", txt({ objectName: "SF label", x: 0.3, y: 6.65, w: 1.6, h: 0.3, fontSize: 10, color: C.accent2, align: "center", charSpacing: 2 }));
    slide.addText("BUSAN", txt({ objectName: "Busan label", x: 5.6, y: 6.65, w: 1.6, h: 0.3, fontSize: 10, color: C.accent1, align: "center", charSpacing: 2 }));

    // 거리 인포그래픽
    slide.addText([
      { text: "9,062", options: { fontSize: 60, bold: true, color: C.background1 } },
      { text: " km", options: { fontSize: 24, color: C.background2 } },
    ], txt({ objectName: "Distance stat", x: M, y: 7.3, w: CW, h: 1.0, align: "center", valign: "middle" }));
    slide.addText("샌프란시스코와 부산 사이, 두 사람이 건너온 거리", txt({ objectName: "Distance caption", x: M, y: 8.35, w: CW, h: 0.35, fontSize: 12, color: C.background2, align: "center" }));
    slide.addText([
      { text: "그리고 오늘부터 두 사람 사이의 거리  " },
      { text: "0 km", options: { bold: true, color: C.accent1 } },
    ], txt({ objectName: "Zero km", x: M, y: 8.75, w: CW, h: 0.35, fontSize: 14, color: C.background1, align: "center" }));

    // 일시·장소 카드
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: "Date card", x: M + 0.3, y: 10.0, w: CW - 0.6, h: 1.9 }));
    slide.addText("WE'RE GETTING MARRIED", txt({ objectName: "Card eyebrow", x: M + 0.3, y: 10.25, w: CW - 0.6, h: 0.3, fontSize: 10, color: C.accent1, align: "center", charSpacing: 5 }));
    slide.addText("2026. 10. 25. SUN  12:30", txt({ objectName: "Date", x: M + 0.3, y: 10.65, w: CW - 0.6, h: 0.5, fontSize: 22, bold: true, color: C.background1, align: "center" }));
    slide.addText("양산 더 지펠리체 5층 팰리스홀", txt({ objectName: "Venue", x: M + 0.3, y: 11.25, w: CW - 0.6, h: 0.4, fontSize: 14, color: C.background2, align: "center" }));
    slide.addNotes("표지. 샌프란시스코와 부산 사이의 직선 거리(약 9,062km)를 인포그래픽으로 표현했습니다.");
  }

  // ════════ 2. INVITATION ════════
  pres.addSection({ title: "초대" });
  {
    const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: "초대" });
    slide.background = await background("1E293B", "0D233A");
    heading(slide, "INVITATION", "소중한 분들을 모십니다");

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: "Letter card", x: M, y: 2.35, w: CW, h: 5.6 }));
    slide.addText([
      { text: "오차 없는 설계를 사랑하는 사람과", options: { breakLine: true } },
      { text: "꾸밈없는 기록을 사랑하는 사람이", options: { breakLine: true } },
      { text: "인생이라는 가장 긴 프로젝트를", options: { breakLine: true } },
      { text: "함께 시작합니다.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "설계도에 없던 행복과", options: { breakLine: true } },
      { text: "대본에 없던 웃음으로 채워 가겠습니다.", options: { breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "귀한 걸음 하시어 두 사람의 첫 장면을", options: { breakLine: true } },
      { text: "따뜻한 축복으로 지켜봐 주시면", options: { breakLine: true } },
      { text: "더없는 기쁨으로 간직하겠습니다." },
    ], txt({ objectName: "Invitation letter", x: M + 0.3, y: 2.6, w: CW - 0.6, h: 5.1, fontSize: 15, color: C.background1, align: "center", valign: "middle", lineSpacingMultiple: 1.45 }));

    // 가족 관계도
    const colW = 2.7, lx = M + 0.15, rx = W - M - 0.15 - colW;
    const fam = [
      { x: lx, parents: "손영길 · 김정숙", rel: "의 아들", name: "규수", color: C.accent1 },
      { x: rx, parents: "문재호 · 박순희", rel: "의 딸", name: "세화", color: C.accent2 },
    ];
    fam.forEach((f, i) => {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: `Parents ${i + 1}`, x: f.x, y: 8.55, w: colW, h: 0.7, rectRadius: 0.35 }));
      slide.addText(f.parents, txt({ objectName: `Parents name ${i + 1}`, x: f.x, y: 8.55, w: colW, h: 0.7, fontSize: 15, color: C.background1, align: "center", valign: "middle" }));
      slide.addShape(pres.shapes.LINE, { objectName: `Family line ${i + 1}`, x: f.x + colW / 2, y: 9.25, w: 0, h: 0.75, line: { color: C.background2, width: 1, dashType: "dash" } });
      slide.addText(f.rel, txt({ objectName: `Relation ${i + 1}`, x: f.x + colW / 2 + 0.12, y: 9.45, w: 1.0, h: 0.3, fontSize: 11, color: C.background2 }));
      const d = 1.5;
      slide.addShape(pres.shapes.OVAL, { objectName: `Couple circle ${i + 1}`, x: f.x + (colW - d) / 2, y: 10.0, w: d, h: d, fill: { color: "FFFFFF", transparency: 90 }, line: { color: f.color, width: 2 } });
      slide.addText(f.name, txt({ objectName: `Couple name ${i + 1}`, x: f.x + (colW - d) / 2, y: 10.0, w: d, h: d, fontSize: 24, bold: true, color: f.color, align: "center", valign: "middle" }));
    });
    slide.addShape(pres.shapes.LINE, { objectName: "Couple link", x: lx + (colW + 1.5) / 2, y: 10.75, w: rx - lx - 1.5, h: 0, line: { color: C.accent6, width: 1.5, dashType: "dash" } });
    slide.addShape(pres.shapes.OVAL, { objectName: "Heart backdrop", x: W / 2 - 0.35, y: 10.4, w: 0.7, h: 0.7, fill: { color: THEME.colors.dk2 }, line: { color: C.accent6, width: 1.5 } });
    slide.addImage({ objectName: "Heart", data: I.heart, x: W / 2 - 0.18, y: 10.57, w: 0.36, h: 0.36 });
    slide.addText("두 가족이 한 가족이 되는 날", txt({ objectName: "Family caption", x: M, y: 11.9, w: CW, h: 0.35, fontSize: 13, color: C.background2, align: "center" }));
    slide.addNotes("초대 글과 양가 혼주 관계도.");
  }

  // ════════ 3. SPEC SHEET ════════
  pres.addSection({ title: "사양서" });
  {
    const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: "사양서" });
    slide.background = await background("0D233A", "0A2E32");
    heading(slide, "SPECIFICATION", "본 결혼의 사양서");
    slide.addText("DWG. No. 20261025  ·  평생 공동 운항 프로젝트", txt({ objectName: "Drawing number", x: M, y: 2.05, w: CW, h: 0.35, fontSize: 12, color: C.background2, align: "center", charSpacing: 1 }));

    const stats = [
      { ic: I.users, num: "2", unit: "명", label: "탑승 인원", note: "정원 초과 시 별도 협의" },
      { ic: I.inf, num: "∞", unit: "", label: "보증 기간", note: "A/S는 서로 무상 제공" },
      { ic: I.clock, num: "24/7", unit: "", label: "운항 시간", note: "연중무휴, 서로 대기 중" },
      { ic: I.cogs, num: "±0", unit: "mm", label: "설계 오차", note: "단, 화해 속도는 빠를수록 좋음" },
      { ic: I.video, num: "0", unit: "컷", label: "편집 분량", note: "무편집 다큐, NG도 소중히 보관" },
      { ic: I.home, num: "1", unit: "곳", label: "공동 주소", note: "이제 같은 주소를 씁니다" },
    ];
    const gap = 0.4, cw = (CW - gap) / 2, ch = 2.55, y0 = 2.75;
    stats.forEach((st, i) => {
      const x = M + (i % 2) * (cw + gap), y = y0 + Math.floor(i / 2) * (ch + 0.35);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: `Spec card ${i + 1}`, x, y, w: cw, h: ch }));
      slide.addShape(pres.shapes.OVAL, { objectName: `Spec icon bg ${i + 1}`, x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6, fill: { color: "2DD4BF", transparency: 80 }, line: { type: "none" } });
      slide.addImage({ objectName: `Spec icon ${i + 1}`, data: st.ic, x: x + 0.45, y: y + 0.45, w: 0.3, h: 0.3 });
      slide.addText(st.label, txt({ objectName: `Spec label ${i + 1}`, x: x + 1.05, y: y + 0.3, w: cw - 1.3, h: 0.6, fontSize: 14, bold: true, color: C.background1, valign: "middle" }));
      slide.addText([
        { text: st.num, options: { fontSize: 40, bold: true, color: C.accent1 } },
        { text: st.unit ? " " + st.unit : "", options: { fontSize: 16, color: C.accent4 } },
      ], txt({ objectName: `Spec value ${i + 1}`, x: x + 0.3, y: y + 1.0, w: cw - 0.6, h: 0.8, valign: "middle" }));
      slide.addText(st.note, txt({ objectName: `Spec note ${i + 1}`, x: x + 0.3, y: y + 1.85, w: cw - 0.5, h: 0.45, fontSize: 11, color: C.background2, valign: "top" }));
    });

    // 승인 도장
    const sy = 11.55;
    slide.addShape(pres.shapes.OVAL, { objectName: "Approval stamp", x: W - M - 1.35, y: sy - 0.15, w: 1.25, h: 1.25, fill: { color: "FCA5A5", transparency: 92 }, line: { color: C.accent6, width: 2 }, rotate: -12 });
    slide.addText([
      { text: "APPROVED", options: { fontSize: 9, breakLine: true } },
      { text: "승인", options: { fontSize: 16, bold: true, breakLine: true } },
      { text: "2026.10.25", options: { fontSize: 9 } },
    ], txt({ objectName: "Approval stamp text", x: W - M - 1.35, y: sy - 0.15, w: 1.25, h: 1.25, color: C.accent6, align: "center", valign: "middle", rotate: -12 }));
    slide.addText([
      { text: "검수일  2026. 10. 25. 일요일 12:30", options: { breakLine: true, bold: true, color: C.background1 } },
      { text: "검수에는 하객 여러분의 박수가 꼭 필요합니다." },
    ], txt({ objectName: "Inspection note", x: M, y: sy, w: CW - 1.6, h: 0.95, fontSize: 13, color: C.background2, valign: "middle", lineSpacingMultiple: 1.3 }));
    slide.addNotes("결혼을 공학 사양서처럼 재치 있게 풀어낸 인포그래픽.");
  }

  // ════════ 4. JOURNEY (Timeline: sky → sea) ════════
  pres.addSection({ title: "이야기" });
  {
    const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: "이야기" });
    slide.background = await background("1E1B4B", "022C22", { bubbles: true });
    heading(slide, "OUR JOURNEY", "하늘에서 바다까지");

    const lineX = 1.85;
    const nodes = [
      { ic: I.planeTeal, color: C.accent2, alt: ["고도", "10,000 m"], scene: "SCENE 01 · THE SKY", title: "샌프란시스코의 하늘 아래", desc: "드넓은 하늘 아래에서 두 사람의 이야기가 시작되었습니다." },
      { ic: I.globe, color: C.accent2, alt: ["태평양", "9,062 km"], scene: "SCENE 02 · THE PACIFIC", title: "태평양을 건너", desc: "오차 없는 설계와 꾸밈없는 기록이 만나 한 편의 이야기가 됩니다." },
      { ic: I.water, color: C.accent1, alt: ["수심", "−5 m"], scene: "SCENE 03 · THE SEA", title: "부산의 맑은 바다로", desc: "스노클링처럼 청량하게, 손을 꼭 잡고 함께 뛰어듭니다." },
      { ic: I.ring, color: C.accent5, alt: ["D-DAY", "10.25"], scene: "SCENE 04 · TO BE CONTINUED", title: "그리고, 결혼식", desc: "다음 장면은 여러분과 함께 찍겠습니다." },
    ];
    const y0 = 2.6, step = 2.55, d = 0.8;
    slide.addShape(pres.shapes.LINE, { objectName: "Timeline", x: lineX, y: y0 + d / 2, w: 0, h: step * 3, line: { color: C.background2, width: 1.25, dashType: "dash" } });
    nodes.forEach((n, i) => {
      const y = y0 + i * step;
      slide.addText([
        { text: n.alt[0], options: { fontSize: 10, color: C.background2, breakLine: true } },
        { text: n.alt[1], options: { fontSize: 12, bold: true, color: n.color } },
      ], txt({ objectName: `Altitude ${i + 1}`, x: 0.3, y: y, w: 1.0, h: d, align: "right", valign: "middle" }));
      slide.addShape(pres.shapes.OVAL, { objectName: `Node ${i + 1}`, x: lineX - d / 2, y, w: d, h: d, fill: { color: n.color }, line: { color: "FFFFFF", width: 1.5, transparency: 40 } });
      slide.addImage({ objectName: `Node icon ${i + 1}`, data: n.ic, x: lineX - 0.2, y: y + 0.2, w: 0.4, h: 0.4 });
      const cx = lineX + 0.65, cwid = W - M - cx;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: `Scene card ${i + 1}`, x: cx, y: y - 0.25, w: cwid, h: 2.1 }));
      slide.addText(n.scene, txt({ objectName: `Scene label ${i + 1}`, x: cx + 0.3, y: y - 0.02, w: cwid - 0.6, h: 0.3, fontSize: 10, color: n.color, charSpacing: 2 }));
      slide.addText(n.title, txt({ objectName: `Scene title ${i + 1}`, x: cx + 0.3, y: y + 0.32, w: cwid - 0.6, h: 0.45, fontSize: 18, bold: true, color: C.background1 }));
      slide.addText(n.desc, txt({ objectName: `Scene desc ${i + 1}`, x: cx + 0.3, y: y + 0.85, w: cwid - 0.6, h: 0.85, fontSize: 13, color: C.background2, valign: "top", lineSpacingMultiple: 1.2 }));
    });
    slide.addNotes("하늘(샌프란시스코)에서 바다(부산)로 내려가는 세로 타임라인. 고도/수심 표기는 재미를 위한 연출입니다.");
  }

  // ════════ 5. DATE ════════
  pres.addSection({ title: "일시" });
  {
    const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: "일시" });
    slide.background = await background("022C22", "043D31", { bubbles: true });
    heading(slide, "THE DATE", "예식 일시");

    slide.addText([
      { text: "10.25", options: { fontSize: 72, bold: true, color: C.background1 } },
    ], txt({ objectName: "Big date", x: M, y: 2.2, w: CW, h: 1.2, align: "center", valign: "middle" }));
    slide.addText("2026년 10월 25일 일요일 낮 12시 30분", txt({ objectName: "Date full", x: M, y: 3.45, w: CW, h: 0.4, fontSize: 16, color: C.accent4, align: "center" }));

    // 달력
    const cy = 4.35, chh = 5.45;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: "Calendar card", x: M, y: cy, w: CW, h: chh }));
    slide.addText("2026  OCTOBER", txt({ objectName: "Calendar month", x: M, y: cy + 0.3, w: CW, h: 0.4, fontSize: 16, bold: true, color: C.background1, align: "center", charSpacing: 4 }));
    const days = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
    const gx = M + 0.3, colw = (CW - 0.6) / 7, rowh = 0.72, ry = cy + 0.95;
    days.forEach((d, c) => slide.addText(d, txt({ objectName: `Weekday ${d}`, x: gx + c * colw, y: ry, w: colw, h: 0.35, fontSize: 10, color: c === 0 ? C.accent6 : C.background2, align: "center", valign: "middle" })));
    const first = new Date(2026, 9, 1).getDay();
    for (let day = 1; day <= 31; day++) {
      const idx = first + day - 1, c = idx % 7, r = Math.floor(idx / 7);
      const x = gx + c * colw, y = ry + 0.5 + r * rowh;
      if (day === 25) {
        const dd = 0.6;
        slide.addShape(pres.shapes.OVAL, { objectName: "Wedding day", x: x + (colw - dd) / 2, y: y + (rowh - 0.1 - dd) / 2, w: dd, h: dd, fill: { color: C.accent1 }, line: { type: "none" } });
      }
      slide.addText(String(day), txt({ objectName: `Day ${day}`, x, y, w: colw, h: rowh - 0.1, fontSize: 14, bold: day === 25, align: "center", valign: "middle",
        color: day === 25 ? C.text1 : c === 0 ? C.accent6 : C.background1 }));
    }

    // 수온 게이지
    const gy = 10.25;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: "Gauge card", x: M, y: gy, w: CW, h: 2.2 }));
    slide.addImage({ objectName: "Thermometer", data: I.temp, x: M + 0.35, y: gy + 0.35, w: 0.4, h: 0.4 });
    slide.addText("당일 수온 예보", txt({ objectName: "Gauge title", x: M + 0.9, y: gy + 0.35, w: 3, h: 0.4, fontSize: 15, bold: true, color: C.background1, valign: "middle" }));
    slide.addText("36.5℃", txt({ objectName: "Gauge value", x: W - M - 2.3, y: gy + 0.3, w: 2.0, h: 0.5, fontSize: 22, bold: true, color: C.accent6, align: "right", valign: "middle" }));
    const tx = M + 0.35, tw = CW - 0.7;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Gauge track", x: tx, y: gy + 1.05, w: tw, h: 0.22, rectRadius: 0.11, fill: { color: "FFFFFF", transparency: 85 }, line: { type: "none" } });
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Gauge fill", x: tx, y: gy + 1.05, w: tw * 0.88, h: 0.22, rectRadius: 0.11, fill: { color: C.accent6 }, line: { type: "none" } });
    slide.addText("여러분의 축하 덕분에 따뜻할 예정입니다", txt({ objectName: "Gauge caption", x: tx, y: gy + 1.45, w: tw, h: 0.4, fontSize: 13, color: C.background2 }));
    slide.addNotes("예식 일시 달력과 '당일 수온' 게이지(재미 요소).");
  }

  // ════════ 6. LOCATION ════════
  pres.addSection({ title: "장소" });
  {
    const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: "장소" });
    slide.background = await background("043D31", "01140E", { bubbles: true });
    heading(slide, "LOCATION", "오시는 길");

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: "Venue card", x: M, y: 2.35, w: CW, h: 1.9 }));
    slide.addImage({ objectName: "Pin", data: I.pin, x: M + 0.35, y: 2.65, w: 0.45, h: 0.45 });
    slide.addText("양산 더 지펠리체", txt({ objectName: "Venue name", x: M + 1.0, y: 2.6, w: CW - 1.3, h: 0.55, fontSize: 24, bold: true, color: C.background1, valign: "middle" }));
    slide.addText("5층 팰리스홀", txt({ objectName: "Venue hall", x: M + 1.0, y: 3.15, w: CW - 1.3, h: 0.35, fontSize: 15, color: C.accent1 }));
    slide.addText("경남 양산시 동면 남양산로 (지펠리체 웨딩)", txt({ objectName: "Venue address", x: M + 1.0, y: 3.55, w: CW - 1.3, h: 0.35, fontSize: 12, color: C.background2 }));

    // 층별 안내 인포그래픽
    const bx = 1.2, bw = 2.9, fh = 0.8, fy0 = 4.85;
    slide.addShape(pres.shapes.ISOSCELES_TRIANGLE, { objectName: "Roof", x: bx - 0.15, y: fy0 - 0.65, w: bw + 0.3, h: 0.6, fill: { color: "FFFFFF", transparency: 88 }, line: { color: "FFFFFF", transparency: 70, width: 0.75 } });
    for (let f = 5; f >= 1; f--) {
      const y = fy0 + (5 - f) * (fh + 0.08);
      const hit = f === 5;
      slide.addShape(pres.shapes.RECTANGLE, { objectName: `Floor ${f}`, x: bx, y, w: bw, h: fh,
        fill: hit ? { color: C.accent1 } : { color: "FFFFFF", transparency: 90 },
        line: hit ? { type: "none" } : { color: "FFFFFF", transparency: 75, width: 0.75 } });
      slide.addText(hit ? "5F  팰리스홀" : `${f}F`, txt({ objectName: `Floor label ${f}`, x: bx, y, w: bw, h: fh, fontSize: hit ? 16 : 12, bold: hit, color: hit ? C.text1 : C.background2, align: "center", valign: "middle" }));
    }
    const ay = fy0 + fh / 2;
    slide.addShape(pres.shapes.LINE, { objectName: "Here arrow", x: bx + bw + 0.15, y: ay, w: 0.7, h: 0, line: { color: C.accent5, width: 2, beginArrowType: "triangle" } });
    slide.addText([
      { text: "여기입니다!", options: { bold: true, color: C.accent5, fontSize: 16, breakLine: true } },
      { text: "5층으로 오세요", options: { color: C.background2, fontSize: 12 } },
    ], txt({ objectName: "Here label", x: bx + bw + 0.95, y: ay - 0.45, w: 2.2, h: 0.9, valign: "middle" }));
    slide.addShape(pres.shapes.LINE, { objectName: "Ground", x: bx - 0.4, y: fy0 + 5 * (fh + 0.08) - 0.04, w: bw + 0.8, h: 0, line: { color: C.background2, width: 1 } });

    // 내비게이션 링크
    const ny = 10.0, nw = (CW - 0.6) / 3;
    const nav = [
      { label: "네이버 지도", url: "https://map.naver.com/p/search/%EC%96%91%EC%82%B0%20%EB%8D%94%20%EC%A7%80%ED%8E%A0%EB%A6%AC%EC%B2%B4", fill: "0D9488", color: C.background1 },
      { label: "카카오맵", url: "https://map.kakao.com/link/search/%EB%8D%94%EC%A7%80%ED%8E%A0%EB%A6%AC%EC%B2%B4", fill: "FDE68A", color: C.text1 },
      { label: "티맵", url: "https://tmap.co.kr", fill: "6366F1", color: C.background1 },
    ];
    slide.addText("내비게이션에서 '더 지펠리체'를 검색하세요", txt({ objectName: "Nav caption", x: M, y: ny - 0.6, w: CW, h: 0.35, fontSize: 13, color: C.background2, align: "center" }));
    nav.forEach((n, i) => {
      const x = M + i * (nw + 0.3);
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: `Nav ${n.label}`, x, y: ny, w: nw, h: 0.7, rectRadius: 0.35, fill: { color: n.fill }, line: { type: "none" }, hyperlink: { url: n.url, tooltip: n.label } });
      slide.addText(n.label, txt({ objectName: `Nav text ${n.label}`, x, y: ny, w: nw, h: 0.7, fontSize: 14, bold: true, color: n.color, align: "center", valign: "middle" }));
    });
    slide.addNotes("장소 안내와 층별 인포그래픽. 상세 주소(번지)는 확인 후 보완이 필요합니다.");
  }

  // ════════ 7. ACCOUNTS & THANKS ════════
  pres.addSection({ title: "마음" });
  {
    const slide = pres.addSlide({ masterName: "CONTENT", sectionTitle: "마음" });
    slide.background = await background("01140E", "030712", { glow: "0D9488" });
    heading(slide, "WITH GRATITUDE", "마음 전하실 곳");
    slide.addText("참석이 어려우신 분들을 위해 조심스레 적어 둡니다.\n보내주신 마음, 오래도록 잊지 않겠습니다.", txt({ objectName: "Account intro", x: M, y: 2.1, w: CW, h: 0.8, fontSize: 13, color: C.background2, align: "center", valign: "middle", lineSpacingMultiple: 1.3 }));

    const accts = [
      { side: "신랑 측", bank: "국민은행", num: "123456-04-123456", holder: "손규수", color: C.accent1 },
      { side: "신부 측", bank: "경남은행", num: "789012-52-654321", holder: "문세화", color: C.accent2 },
    ];
    accts.forEach((a, i) => {
      const y = 3.3 + i * 2.15;
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, glass({ objectName: `Account card ${i + 1}`, x: M, y, w: CW, h: 1.85 }));
      slide.addShape(pres.shapes.OVAL, { objectName: `Account icon bg ${i + 1}`, x: M + 0.35, y: y + 0.35, w: 0.6, h: 0.6, fill: { color: "2DD4BF", transparency: 80 }, line: { type: "none" } });
      slide.addImage({ objectName: `Account icon ${i + 1}`, data: I.wallet, x: M + 0.5, y: y + 0.5, w: 0.3, h: 0.3 });
      slide.addText(a.side, txt({ objectName: `Account side ${i + 1}`, x: M + 1.15, y: y + 0.35, w: 2.5, h: 0.6, fontSize: 15, bold: true, color: a.color, valign: "middle" }));
      slide.addText(`${a.bank}  ${a.num}`, txt({ objectName: `Account number ${i + 1}`, x: M + 0.35, y: y + 1.05, w: CW - 0.7, h: 0.4, fontSize: 16, color: C.background1 }));
      slide.addText(`예금주 ${a.holder}`, txt({ objectName: `Account holder ${i + 1}`, x: M + 0.35, y: y + 1.42, w: CW - 0.7, h: 0.3, fontSize: 12, color: C.background2 }));
    });

    slide.addShape(pres.shapes.OVAL, { objectName: "Closing heart bg", x: W / 2 - 0.45, y: 8.15, w: 0.9, h: 0.9, fill: { color: "FCA5A5", transparency: 85 }, line: { type: "none" } });
    slide.addImage({ objectName: "Closing heart", data: I.heart, x: W / 2 - 0.22, y: 8.38, w: 0.44, h: 0.44 });
    slide.addText("함께해 주시는 모든 분들께\n진심으로 감사드립니다", txt({ objectName: "Thanks", x: M, y: 9.3, w: CW, h: 1.0, fontSize: 20, bold: true, color: C.background1, align: "center", valign: "middle", lineSpacingMultiple: 1.25 }));
    slide.addText("규수 그리고 세화 올림", txt({ objectName: "Signature", x: M, y: 10.45, w: CW, h: 0.4, fontSize: 15, color: C.accent4, align: "center" }));
    slide.addText("© 2026 GyuSoo & SeHwa Wedding", txt({ objectName: "Copyright", x: M, y: 12.3, w: CW, h: 0.3, fontSize: 10, color: C.background2, align: "center", charSpacing: 3 }));
    slide.addNotes("마음 전하실 곳과 감사 인사. 계좌번호는 실제 번호로 확인/교체 필요.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
