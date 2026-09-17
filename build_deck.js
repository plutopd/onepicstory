const pptxgen = require("pptxgenjs");
const path = require("path");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.author = "pluto_hy";
pres.title = "One Pick Story — 데모데이 발표자료";

/* ------------------------------------------------------------------
   TYPOGRAPHY

   The font family installed on this Mac is exactly "Pretendard" — one
   family with nine styles. "Pretendard Black" / "SemiBold" etc. are style
   *Full Names*, not family names, so PowerPoint cannot resolve them and
   silently substitutes a wider fallback face. That is what broke the last
   build (headlines wrapped, letter-spacing looked wrong). So: one family
   name, weight expressed only through `bold`.
------------------------------------------------------------------- */
const F = "Pretendard";

/* ------------------------------------------------------------------
   PALETTE — warm light beige, after the two supplied references
   (beige.social, SIENNA portfolio). Every value below is measured
   against the surface it sits on; ratios in comments.
------------------------------------------------------------------- */
const BG = "F2EDE6";      // page — warm light beige
const CREAM = "FBF9F6";   // cards and panels
const INK = "33281F";     // headings + body            12.32:1 on BG
const INK2 = "5C4E43";    // secondary body              6.87:1 on BG
const MUTED = "726356";   // captions and sources        4.96:1 on BG
const ACCENT = "B04517";  // terracotta — emphasis TEXT  4.86:1 on BG
const ACCENT_FILL = "C7561F"; // same hue for shapes/fills (not small text)
const SOFT = "EFDFD0";    // accent tint panel
const TERRA = "B5846A";   // reference-8 block colour
const TERRA_D = "A25A32"; // full-bleed ground, cream text  4.94:1
const SAGE = "5A6950";    // secondary accent             5.05:1 on BG
const SAGE_LT = "E4E8E0";
const HAIR = "DCD3C8";    // hairline rules

const W = 13.333;
const H = 7.5;
const M = 0.65;           // page margin, fixed
const COL_R = 7.35;       // right column x, fixed across slides
const COL_RW = 5.33;      // right column width

/* ---------- shared pieces ---------- */

function page(s, n) {
  s.background = { color: BG };
  if (n) {
    s.addText(String(n).padStart(2, "0"), {
      x: W - 1.3, y: 0.3, w: 0.85, h: 0.3,
      fontFace: F, fontSize: 11, color: MUTED,
      align: "right", isTextBox: true, margin: 0,
    });
    s.addText("One Pick Story", {
      x: M, y: H - 0.52, w: 3, h: 0.3,
      fontFace: F, fontSize: 10.5, bold: true, color: MUTED,
      isTextBox: true, margin: 0,
    });
  }
}

// Small letterspaced label — the reference decks' section marker.
function label(s, text, x, y, color) {
  s.addText(text, {
    x, y, w: 5, h: 0.28,
    fontFace: F, fontSize: 10.5, bold: true, color: color || ACCENT,
    charSpacing: 2, isTextBox: true, margin: 0,
  });
}

// Title carries the conclusion; the clause that matters is set in the
// accent colour so the point lands without extra decoration.
function title(s, runs, sub) {
  s.addText(
    runs.map((r) => ({
      text: r[0],
      options: { color: r[1] || INK, bold: true, breakLine: r[2] === true },
    })),
    {
      x: M, y: 0.62, w: W - M * 2, h: 1.0,
      fontFace: F, fontSize: 29, isTextBox: true, margin: 0,
      lineSpacingMultiple: 1.18,
    }
  );
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.72, w: W - M * 2, h: 0.32,
      fontFace: F, fontSize: 13, color: INK2,
      isTextBox: true, margin: 0,
    });
  }
}

function card(s, x, y, w, h, fill, line) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: fill },
    line: line ? { color: line, width: 1 } : { type: "none" },
  });
}

// Content separator between items inside a card — never under a title and
// never along a card edge.
function rule(s, x, y, w, color) {
  s.addShape(pres.ShapeType.rect, {
    x, y, w, h: 0.012,
    fill: { color: color || HAIR }, line: { type: "none" },
  });
}

function dot(s, x, y, d, color) {
  s.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color }, line: { type: "none" } });
}

// Numbered marker + heading + description, on one shared baseline grid.
function numRow(s, n, x, y, wHead, wBody, head, body) {
  dot(s, x, y + 0.02, 0.28, SOFT);
  s.addText(String(n), {
    x, y: y + 0.02, w: 0.28, h: 0.28,
    fontFace: F, fontSize: 10.5, bold: true, color: ACCENT,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText(head, {
    x: x + 0.44, y, w: wHead, h: 0.3,
    fontFace: F, fontSize: 13, bold: true, color: INK,
    isTextBox: true, margin: 0,
  });
  s.addText(body, {
    x: x + 0.44, y: y + 0.32, w: wBody, h: 0.32,
    fontFace: F, fontSize: 12, color: INK2,
    isTextBox: true, margin: 0,
  });
}

// Figure + caption. Sizes are deliberately restrained — the last build read
// as a poster because the numerals were set at 84–92pt.
function stat(s, x, y, w, value, unit, caption, color) {
  const runs = [{ text: value, options: { fontSize: 44, bold: true, color: color || ACCENT } }];
  // Korean counters read as a separate word ("수백 장"), unlike "%".
  if (unit) {
    const sep = /^[0-9%]/.test(unit) ? "" : " ";
    runs.push({ text: sep + unit, options: { fontSize: 19, bold: true, color: color || ACCENT } });
  }
  s.addText(runs, {
    x, y, w, h: 0.84,
    fontFace: F, isTextBox: true, margin: 0,
  });
  s.addText(caption, {
    x, y: y + 0.80, w, h: 0.3,
    fontFace: F, fontSize: 12, color: INK2,
    isTextBox: true, margin: 0,
  });
}

// The course material's third rule: a slide carrying numbers states where
// they came from.
function source(s, text) {
  s.addText(text, {
    x: M, y: H - 0.85, w: W - M * 2 - 1.0, h: 0.28,
    fontFace: F, fontSize: 10, color: MUTED,
    isTextBox: true, margin: 0,
  });
}

function bullets(s, items, opts) {
  s.addText(
    items.map((t, i) => ({
      text: t,
      options: { bullet: true, breakLine: i !== items.length - 1 },
    })),
    {
      fontFace: F, fontSize: 12.5, color: INK2,
      paraSpaceAfter: 12, lineSpacingMultiple: 1.28,
      isTextBox: true, margin: 0, ...opts,
    }
  );
}

/* ================= 1. 표지 ================= */
{
  const s = pres.addSlide();
  s.background = { color: TERRA_D };

  // One meaningful object rather than abstract blocks: the thing the service
  // actually produces, sitting on a calm terracotta ground (reference 8).
  s.addShape(pres.ShapeType.rect, { x: 8.55, y: 0, w: 4.78, h: H, fill: { color: TERRA }, line: { type: "none" } });
  s.addImage({
    path: path.join(__dirname, "deck_assets", "ui-phoneframe.png"),
    x: 9.5, y: 1.05, w: 2.88, h: 5.4,
  });

  s.addText("2026 소프트웨어 인재키움사업 · 바이브코딩 창업교육과정", {
    x: 1.0, y: 0.72, w: 7.2, h: 0.3,
    fontFace: F, fontSize: 11, bold: true, color: CREAM, charSpacing: 1.5,
    transparency: 22, isTextBox: true, margin: 0,
  });

  s.addText("One Pick Story", {
    x: 1.0, y: 2.05, w: 9, h: 0.55,
    fontFace: F, fontSize: 16, bold: true, color: CREAM, charSpacing: 5,
    isTextBox: true, margin: 0,
  });

  s.addText(
    [
      { text: "사진첩에 잠든 여행을,", options: { breakLine: true } },
      { text: "영상 한 편으로." },
    ],
    {
      x: 1.0, y: 2.75, w: 7.2, h: 1.9,
      fontFace: F, fontSize: 46, bold: true, color: CREAM,
      lineSpacingMultiple: 1.16, isTextBox: true, margin: 0,
    }
  );

  rule(s, 1.0, 4.95, 1.6, CREAM);

  s.addText("수백 장을 찍지만 편집은 못 하는 부모님이, 사진만 올리면 끝나는 서비스", {
    x: 1.0, y: 5.25, w: 7.2, h: 0.75,
    fontFace: F, fontSize: 14.5, color: CREAM, transparency: 12,
    lineSpacingMultiple: 1.35,
    isTextBox: true, margin: 0,
  });

  s.addText("발표자  pluto_hy", {
    x: 1.0, y: H - 0.85, w: 5, h: 0.3,
    fontFace: F, fontSize: 11.5, color: CREAM, transparency: 28,
    isTextBox: true, margin: 0,
  });

  s.addNotes(
    "[20초]\n" +
    "안녕하세요, One Pick Story입니다.\n" +
    "여행에서 사진은 수백 장씩 찍지만 편집은 못 하는 분들이, 사진만 올리면 AI가 알아서 골라 영상으로 만들어 주는 서비스입니다."
  );
}

/* ================= 2. 문제 ================= */
{
  const s = pres.addSlide();
  page(s, 2);
  label(s, "PROBLEM", M, 0.32);
  title(
    s,
    [["고르는 데 지쳐, 결국 ", INK, true], ["아무것도 만들지 못한다", ACCENT]],
    "여행을 다녀온 저희 부모님의 사진첩 이야기"
  );

  const rows = [
    ["대상", "여행마다 수백 장을 찍는 60대 부모님"],
    ["지금 방식", "사진첩을 처음부터 넘기며 눈으로 하나씩 고름"],
    ["치르는 비용", "고르다 지쳐 중단. 결국 공유하는 건 몇 장뿐"],
    ["여태 안 된 이유", "편집 앱은 배우는 데 시간이 들고, 대신 만들어 줄 사람이 없음"],
  ];
  let y = 2.35;
  rows.forEach((r, i) => {
    numRow(s, i + 1, M, y, 2.4, 6.1, r[0], r[1]);
    y += 0.95;
  });

  // Right column: one figure card, then the quote — the bottom-right corner
  // was dead space in the previous build.
  card(s, COL_R, 2.35, COL_RW, 2.85, CREAM);
  stat(s, COL_R + 0.45, 2.62, 4.4, "수백", "장", "여행 한 번에 찍는 사진", ACCENT);
  rule(s, COL_R + 0.45, 3.78, 4.43);
  stat(s, COL_R + 0.45, 4.0, 4.4, "0", "편", "그 사진으로 만든 영상", INK);

  card(s, COL_R, 5.42, COL_RW, 1.12, SOFT);
  s.addText(
    [
      { text: "지인은 같은 사진으로 브이로그를 만들어 공유하는데,", options: { breakLine: true } },
      { text: "부모님은 그걸 못 합니다.", options: { bold: true, color: ACCENT } },
    ],
    {
      x: COL_R + 0.38, y: 5.42, w: COL_RW - 0.76, h: 1.12,
      fontFace: F, fontSize: 12.5, color: INK, lineSpacingMultiple: 1.35,
      valign: "middle", isTextBox: true, margin: 0,
    }
  );

  source(s, "출처: 부모님·지인 관찰 및 대화 · 2026.8");

  s.addNotes(
    "[50초]\n" +
    "저희 부모님은 여행을 다니실 때마다 사진을 수백 장씩 찍으십니다.\n" +
    "그런데 그 사진으로 만든 영상은 지금까지 한 편도 없습니다.\n\n" +
    "이유는 두 가지였습니다. 첫째, 수백 장 중에 어떤 사진이 잘 나왔는지를 일일이 넘겨 보며 고르는 일이 너무 힘들어서 중간에 포기하십니다.\n" +
    "둘째, 지인분들은 같은 여행 사진으로 브이로그를 만들어 공유하시는데, 부모님은 편집을 배운 적이 없어서 그걸 못 하십니다.\n\n" +
    "여기서 필요를 발견했습니다."
  );
}

/* ================= 3. 해결 ================= */
{
  const s = pres.addSlide();
  page(s, 3);
  label(s, "SOLUTION", M, 0.32);
  title(
    s,
    [["사진만 올리면, ", INK], ["나머지는 AI가", ACCENT]],
    "앞에서 짚은 세 가지를 그대로 대응시켰습니다"
  );

  const rows = [
    ["사진 고르기", "수백 장을 일일이 넘겨 확인", "AI 사진 선별", "눈감음·중복·흐린 사진 자동 제외"],
    ["편집", "편집 앱을 새로 배워야 함", "AI 감독", "순서·길이·전환·색감 자동 결정"],
    ["결과물", "사진첩에 그대로 방치", "완성 영상 파일", "바로 공유 가능한 상태로 저장"],
  ];

  let y = 2.35;
  rows.forEach((r, i) => {
    card(s, M, y, 6.35, 1.2, CREAM);
    s.addText(r[0], {
      x: M + 0.32, y: y + 0.18, w: 1.85, h: 0.3,
      fontFace: F, fontSize: 12.5, bold: true, color: INK,
      isTextBox: true, margin: 0,
    });
    s.addText(r[1], {
      x: M + 0.32, y: y + 0.55, w: 2.0, h: 0.45,
      fontFace: F, fontSize: 11, color: MUTED,
      isTextBox: true, margin: 0,
    });
    s.addShape(pres.ShapeType.rect, {
      x: M + 2.55, y: y + 0.28, w: 0.012, h: 0.64,
      fill: { color: HAIR }, line: { type: "none" },
    });
    s.addText(r[2], {
      x: M + 2.85, y: y + 0.18, w: 3.3, h: 0.3,
      fontFace: F, fontSize: 13.5, bold: true, color: ACCENT,
      isTextBox: true, margin: 0,
    });
    s.addText(r[3], {
      x: M + 2.85, y: y + 0.55, w: 3.3, h: 0.45,
      fontFace: F, fontSize: 11, color: INK2,
      isTextBox: true, margin: 0,
    });
    y += 1.38;
  });

  // One visual, per the course material: the real curation screen.
  card(s, COL_R, 2.35, COL_RW, 2.72, CREAM);
  s.addImage({
    path: path.join(__dirname, "deck_assets", "ui-curation.png"),
    x: COL_R + 0.3, y: 2.62, w: 4.73, h: 2.09,
  });
  s.addText("실제 화면 — 눈감음·중복 사진이 회색으로 빠집니다", {
    x: COL_R, y: 5.18, w: COL_RW, h: 0.3,
    fontFace: F, fontSize: 10.5, color: MUTED, align: "center",
    isTextBox: true, margin: 0,
  });

  card(s, COL_R, 5.62, COL_RW, 0.92, SOFT);
  s.addText(
    [
      { text: "AI는 판단만 하고, 실제 합성은 ", options: {} },
      { text: "ffmpeg", options: { bold: true, color: ACCENT } },
      { text: "가 담당합니다.", options: {} },
    ],
    {
      x: COL_R + 0.38, y: 5.62, w: COL_RW - 0.76, h: 0.92,
      fontFace: F, fontSize: 12.5, color: INK,
      valign: "middle", isTextBox: true, margin: 0,
    }
  );

  s.addNotes(
    "[45초]\n" +
    "그래서 앞의 세 가지를 그대로 대응시켰습니다.\n\n" +
    "사진 고르기는 AI 사진 선별이 대신합니다. 눈을 감았거나 초점이 흔들렸거나 비슷한 사진이 여러 장이면 자동으로 걸러 냅니다.\n" +
    "편집은 AI 감독이 대신합니다. 원하는 분위기를 한 줄로 적거나, 준비된 문구를 고르기만 하면 순서와 길이, 전환 효과와 색감을 알아서 정합니다.\n" +
    "그 결과로 바로 공유할 수 있는 영상 파일이 나옵니다.\n\n" +
    "AI는 판단만 하고, 실제 영상 합성은 검증된 오픈소스 렌더링 엔진이 담당합니다."
  );
}

/* ================= 4. 시연 전환 ================= */
{
  const s = pres.addSlide();
  s.background = { color: TERRA_D };
  dot(s, 9.55, 2.15, 3.2, TERRA);

  s.addText("DEMO", {
    x: 1.0, y: 2.85, w: 4, h: 0.3,
    fontFace: F, fontSize: 11, bold: true, color: CREAM, charSpacing: 3,
    transparency: 25, isTextBox: true, margin: 0,
  });
  s.addText("시연", {
    x: 1.0, y: 3.25, w: 6, h: 1.0,
    fontFace: F, fontSize: 52, bold: true, color: CREAM,
    isTextBox: true, margin: 0,
  });
  rule(s, 1.0, 4.45, 1.6, CREAM);
  s.addText("사진 업로드부터 완성된 영상까지, 실제 화면으로 보여 드리겠습니다.", {
    x: 1.0, y: 4.72, w: 8.0, h: 0.35,
    fontFace: F, fontSize: 14.5, color: CREAM, transparency: 12,
    isTextBox: true, margin: 0,
  });

  s.addNotes("[5초]\n실제 화면으로 바로 보여 드리겠습니다.");
}

/* ================= 5. 시연 영상 자리 ================= */
{
  const s = pres.addSlide();
  page(s, 5);
  label(s, "DEMO", M, 0.32);
  title(s, [["사진 업로드부터 ", INK], ["완성 영상까지", ACCENT]], null);

  const scenes = [
    ["시작 화면", "로그인된 상태에서 시작"],
    ["사진 업로드 · AI 선별", "제외된 사진이 실시간으로 회색 처리"],
    ["AI 편집 · 완성 영상", "순서·전환·색감이 정해지고 내보내기"],
  ];
  let sy = 2.2;
  scenes.forEach((sc, i) => {
    numRow(s, i + 1, M, sy, 3.2, 4.4, sc[0], sc[1]);
    sy += 1.0;
  });

  card(s, M, 5.35, 4.6, 1.1, SOFT);
  s.addText(
    [
      { text: "실시간 시연이 막히면 ", options: {} },
      { text: "15초 안에", options: { bold: true, color: ACCENT } },
      { text: " 이 녹화본으로 전환", options: {} },
    ],
    {
      x: M + 0.35, y: 5.35, w: 3.9, h: 1.1,
      fontFace: F, fontSize: 12, color: INK,
      valign: "middle", isTextBox: true, margin: 0,
    }
  );

  // 9:16 placeholder, shaped like the product's own preview frame.
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.75, y: 2.15, w: 3.4, h: 4.55, rectRadius: 0.16,
    fill: { color: CREAM }, line: { color: TERRA, width: 1.5, dashType: "dash" },
  });
  dot(s, 7.15, 4.05, 0.6, SOFT);
  s.addText("▶", {
    x: 7.15, y: 4.05, w: 0.6, h: 0.6,
    fontFace: F, fontSize: 19, color: ACCENT,
    align: "center", valign: "middle", isTextBox: true, margin: 0,
  });
  s.addText("이 도형을 지우고\n시연 영상을 넣어 주세요", {
    x: 6.0, y: 4.85, w: 2.9, h: 0.7,
    fontFace: F, fontSize: 11.5, color: MUTED, align: "center",
    lineSpacingMultiple: 1.3, isTextBox: true, margin: 0,
  });

  card(s, 9.65, 2.15, 3.03, 4.55, CREAM);
  s.addText("장면별 참고 화면", {
    x: 9.95, y: 2.42, w: 2.45, h: 0.28,
    fontFace: F, fontSize: 11, bold: true, color: MUTED,
    isTextBox: true, margin: 0,
  });
  s.addImage({
    path: path.join(__dirname, "deck_assets", "ui-curation.png"),
    x: 9.95, y: 2.85, w: 2.43, h: 1.07,
  });
  s.addImage({
    path: path.join(__dirname, "deck_assets", "ui-timeline.png"),
    x: 9.95, y: 4.12, w: 2.43, h: 0.52,
  });
  s.addText(
    [
      { text: "AI 선별", options: { bold: true, color: ACCENT, breakLine: true } },
      { text: "흐림·중복·눈감음 제외", options: { color: INK2, breakLine: true } },
      { text: " ", options: { breakLine: true } },
      { text: "타임라인", options: { bold: true, color: ACCENT, breakLine: true } },
      { text: "클립별 전환 효과 지정", options: { color: INK2 } },
    ],
    {
      x: 9.95, y: 4.85, w: 2.45, h: 1.5,
      fontFace: F, fontSize: 10.5, lineSpacingMultiple: 1.25,
      isTextBox: true, margin: 0,
    }
  );

  s.addNotes(
    "[80초]  배점이 가장 큰 구간입니다. 영상은 소리 없이도 이해되도록 자막을 넣어 주세요.\n\n" +
    "장면 1 (10초) 시작 화면 — 로그인된 상태에서 시작합니다.\n" +
    "장면 2 (30초) 여행 사진을 한 번에 올리고, AI 사진 선별을 실행합니다. 눈감은 사진과 중복 사진이 회색으로 빠지는 것이 실시간으로 보입니다.\n" +
    "장면 3 (30초) 원하는 분위기를 고르고 AI 감독을 실행하면 순서와 전환, 색감이 정해집니다. 마지막에 영상을 내보내 완성본을 보여 줍니다.\n" +
    "장면 4 (10초) 완성된 영상 재생.\n\n" +
    "※ 실시간 시연이 막히면 15초 안에 이 녹화본으로 전환할 것."
  );
}

/* ================= 6. 증거 ================= */
{
  const s = pres.addSlide();
  page(s, 6);
  label(s, "EVIDENCE", M, 0.32);
  title(
    s,
    [["15명이 직접 만들어 봤고, ", INK], ["전원 완성했습니다", ACCENT]],
    "부모님과 지인 대상 사용 테스트"
  );

  // Three columns on one baseline; content is distributed across the full
  // card height rather than clustered in the middle.
  const top = 2.35, h = 3.75;

  card(s, M, top, 3.35, h, CREAM);
  stat(s, M + 0.4, top + 0.42, 2.6, "15", "명", "서비스를 전달한 사용자", ACCENT);
  rule(s, M + 0.4, top + 1.78, 2.55);
  stat(s, M + 0.4, top + 2.1, 2.6, "100", "%", "영상 완성까지 도달", ACCENT);

  card(s, 4.35, top, 4.1, h, SOFT);
  s.addText("검증한 것", {
    x: 4.72, y: top + 0.35, w: 3.4, h: 0.3,
    fontFace: F, fontSize: 13, bold: true, color: ACCENT,
    isTextBox: true, margin: 0,
  });
  const checked = [
    "편집 경험 없이도 완성까지 도달",
    "중간에 막혀 포기한 사용자 없음",
    "사진 선별이 가장 유용하다는 반응",
  ];
  checked.forEach((t, i) => {
    const y = top + 0.95 + i * 0.95;
    dot(s, 4.72, y + 0.09, 0.1, ACCENT_FILL);
    s.addText(t, {
      x: 4.95, y, w: 3.2, h: 0.55,
      fontFace: F, fontSize: 12, color: INK,
      isTextBox: true, margin: 0,
    });
  });

  card(s, 8.6, top, 4.08, h, CREAM);
  s.addText("다음에 확인할 것", {
    x: 8.97, y: top + 0.35, w: 3.4, h: 0.3,
    fontFace: F, fontSize: 13, bold: true, color: SAGE,
    isTextBox: true, margin: 0,
  });
  const next = [
    ["지인 외 사용자", "현재 표본은 모두 지인입니다"],
    ["재사용 여부", "두 번째 영상을 만드는지"],
    ["결제 의사", "무료 2회 이후 전환율"],
  ];
  next.forEach((n, i) => {
    const y = top + 0.95 + i * 0.95;
    dot(s, 8.97, y + 0.02, 0.26, SAGE_LT);
    s.addText(String(i + 1), {
      x: 8.97, y: y + 0.02, w: 0.26, h: 0.26,
      fontFace: F, fontSize: 10, bold: true, color: SAGE,
      align: "center", valign: "middle", isTextBox: true, margin: 0,
    });
    s.addText(n[0], {
      x: 9.38, y, w: 2.9, h: 0.28,
      fontFace: F, fontSize: 12.5, bold: true, color: INK,
      isTextBox: true, margin: 0,
    });
    s.addText(n[1], {
      x: 9.38, y: y + 0.3, w: 2.9, h: 0.28,
      fontFace: F, fontSize: 11, color: MUTED,
      isTextBox: true, margin: 0,
    });
  });

  source(s, "출처: 부모님·지인 15명 직접 전달 후 사용 관찰 · 2026.9");

  s.addNotes(
    "[35초]\n" +
    "부모님을 포함해 지인 15분께 직접 전달했고, 열다섯 분 모두 영상 완성까지 도달했습니다. 편집을 해 본 적 없는 분들도 중간에 막혀 포기한 경우가 없었습니다.\n\n" +
    "다만 현재 표본은 전부 지인입니다. 그래서 다음 단계는 지인이 아닌 사용자를 확보해서, 두 번째 영상을 만드는지, 무료 횟수를 다 쓴 뒤 결제로 이어지는지를 확인하는 것입니다.\n\n" +
    "※ '지인 아닌 사용자는?' 질문이 나오면 이 슬라이드를 가리키며 다음 단계로 답할 것."
  );
}

/* ================= 7. 차별점 ================= */
{
  const s = pres.addSlide();
  page(s, 7);
  label(s, "MOAT", M, 0.32);
  title(
    s,
    [["복제되는 것과 ", INK], ["쌓이는 것", ACCENT]],
    "같은 도구를 쓰면 화면은 하루면 따라옵니다"
  );

  const top = 2.35, h = 2.95;

  card(s, M, top, 5.85, h, CREAM);
  s.addText("하루면 복제되는 것", {
    x: M + 0.42, y: top + 0.38, w: 4.9, h: 0.3,
    fontFace: F, fontSize: 14, bold: true, color: MUTED,
    isTextBox: true, margin: 0,
  });
  ["화면과 기능 구성", "사용 중인 AI 모델", "렌더링 엔진과 전환 효과"].forEach((t, i) => {
    const y = top + 1.0 + i * 0.6;
    dot(s, M + 0.42, y + 0.11, 0.1, HAIR);
    s.addText(t, {
      x: M + 0.66, y, w: 4.6, h: 0.32,
      fontFace: F, fontSize: 12.5, color: INK2,
      isTextBox: true, margin: 0,
    });
  });

  card(s, 6.95, top, 5.73, h, SOFT);
  s.addText("6개월이 걸리는 것", {
    x: 7.37, y: top + 0.38, w: 4.9, h: 0.3,
    fontFace: F, fontSize: 14, bold: true, color: ACCENT,
    isTextBox: true, margin: 0,
  });
  ["15명의 실제 사용 기록", "부모 세대가 '잘 나왔다'고 느끼는 기준", "선별 결과가 쌓여 만드는 취향 데이터"].forEach((t, i) => {
    const y = top + 1.0 + i * 0.6;
    dot(s, 7.37, y + 0.11, 0.1, ACCENT_FILL);
    s.addText(t, {
      x: 7.61, y, w: 4.6, h: 0.32,
      fontFace: F, fontSize: 12.5, color: INK,
      isTextBox: true, margin: 0,
    });
  });

  card(s, M, 5.62, 12.03, 0.92, TERRA_D);
  s.addText(
    [
      { text: "후발 주자가 돈으로 살 수 없는 것은 ", options: { color: CREAM } },
      { text: "이 기간", options: { color: CREAM, bold: true } },
      { text: "입니다.", options: { color: CREAM } },
    ],
    {
      x: M + 0.42, y: 5.62, w: 11.2, h: 0.92,
      fontFace: F, fontSize: 16, bold: true,
      valign: "middle", isTextBox: true, margin: 0,
    }
  );

  s.addNotes(
    "[30초]\n" +
    "같은 도구를 쓰면 화면과 기능은 하루면 따라올 수 있습니다. 그건 경쟁력이 아니라 참가 자격이라고 생각합니다.\n\n" +
    "대신 지금 쌓이고 있는 건 실제 사용 기록과, 부모 세대가 어떤 사진을 '잘 나왔다'고 느끼는지에 대한 기준입니다.\n" +
    "이건 후발 주자가 돈으로 살 수 없고, 먼저 시작한 기간만큼만 쌓입니다."
  );
}

/* ================= 8. 사업화 ================= */
{
  const s = pres.addSlide();
  page(s, 8);
  label(s, "BUSINESS", M, 0.32);
  title(
    s,
    [["미리보기는 무제한, ", INK], ["내보낼 때만 과금", ACCENT]],
    "원가가 사용량에 비례하기 때문에 선택한 구조"
  );

  card(s, M, 2.3, 12.03, 3.3, CREAM);
  s.addImage({
    path: path.join(__dirname, "deck_assets", "ui-pricing.png"),
    x: 1.05, y: 2.5, w: 11.23, h: 2.9,
  });

  card(s, M, 5.78, 5.85, 0.88, SOFT);
  s.addText(
    [
      { text: "사진 장수", options: { bold: true, color: ACCENT } },
      { text: "와 ", options: {} },
      { text: "내보내기 횟수", options: { bold: true, color: ACCENT } },
      { text: " — 원가가 늘어나는 두 축", options: {} },
    ],
    {
      x: M + 0.38, y: 5.78, w: 5.1, h: 0.88,
      fontFace: F, fontSize: 12, color: INK,
      valign: "middle", isTextBox: true, margin: 0,
    }
  );
  card(s, 6.95, 5.78, 5.73, 0.88, CREAM);
  s.addText("결제 주체는 사용자 본인. 영상 한 편을 만들어 본 뒤에 결제 시점이 옵니다.", {
    x: 7.33, y: 5.78, w: 5.0, h: 0.88,
    fontFace: F, fontSize: 12, color: INK2,
    valign: "middle", isTextBox: true, margin: 0,
  });

  s.addNotes(
    "[25초]\n" +
    "요금은 미리보기와 편집은 무제한으로 두고, 완성한 영상을 내보낼 때만 횟수를 쓰는 구조입니다.\n" +
    "사진이 많아질수록 AI 호출이 늘고, 영상이 길수록 렌더링 시간이 늘기 때문에, 원가가 늘어나는 두 축인 사진 장수와 내보내기 횟수로 등급을 나눴습니다.\n" +
    "결제하는 사람은 사용자 본인이고, 무료 2회로 완성본을 받아 본 다음에 결제 시점이 옵니다.\n\n" +
    "※ '운영 비용 감당?' 질문 대비 — 렌더링은 서버 CPU 비용, AI 선별은 호출당 비용. 둘 다 내보내기 횟수에 비례."
  );
}

/* ================= 9. 계획 · 마무리 ================= */
{
  const s = pres.addSlide();
  s.background = { color: TERRA_D };
  dot(s, 11.2, 0.55, 1.6, TERRA);

  s.addText("NEXT", {
    x: 1.0, y: 0.8, w: 4, h: 0.3,
    fontFace: F, fontSize: 11, bold: true, color: CREAM, charSpacing: 3,
    transparency: 25, isTextBox: true, margin: 0,
  });
  s.addText("사진첩에 잠든 여행을,\n영상 한 편으로.", {
    x: 1.0, y: 1.25, w: 8.5, h: 1.5,
    fontFace: F, fontSize: 32, bold: true, color: CREAM,
    lineSpacingMultiple: 1.22, isTextBox: true, margin: 0,
  });

  const colY = 3.25;
  card(s, 1.0, colY, 5.3, 2.2, TERRA);
  s.addText("다음 단계", {
    x: 1.38, y: colY + 0.3, w: 4.5, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: CREAM, charSpacing: 1.5,
    isTextBox: true, margin: 0,
  });
  ["지인이 아닌 실사용자 확보", "두 번째 영상까지 만드는지 관찰", "무료 2회 이후 유료 전환 실험"].forEach((t, i) => {
    const y = colY + 0.78 + i * 0.44;
    dot(s, 1.38, y + 0.1, 0.09, CREAM);
    s.addText(t, {
      x: 1.6, y, w: 4.4, h: 0.3,
      fontFace: F, fontSize: 12.5, color: CREAM,
      isTextBox: true, margin: 0,
    });
  });

  card(s, 6.7, colY, 5.3, 2.2, TERRA);
  s.addText("필요한 지원", {
    x: 7.08, y: colY + 0.3, w: 4.5, h: 0.3,
    fontFace: F, fontSize: 12, bold: true, color: CREAM, charSpacing: 1.5,
    isTextBox: true, margin: 0,
  });
  ["렌더링 서버 인프라 비용", "초기 사용자 확보를 위한 마케팅"].forEach((t, i) => {
    const y = colY + 0.78 + i * 0.44;
    dot(s, 7.08, y + 0.1, 0.09, CREAM);
    s.addText(t, {
      x: 7.3, y, w: 4.4, h: 0.3,
      fontFace: F, fontSize: 12.5, color: CREAM,
      isTextBox: true, margin: 0,
    });
  });

  // The live link is a graded deliverable, so it gets read from the back row.
  card(s, 1.0, 5.75, 11.0, 0.95, CREAM);
  s.addText("지금 접속해 보실 수 있습니다", {
    x: 1.4, y: 5.88, w: 4.4, h: 0.28,
    fontFace: F, fontSize: 11, bold: true, color: MUTED, charSpacing: 1,
    isTextBox: true, margin: 0,
  });
  s.addText("vibe-cut.onrender.com", {
    x: 1.4, y: 6.14, w: 7.2, h: 0.42,
    fontFace: F, fontSize: 21, bold: true, color: ACCENT,
    isTextBox: true, margin: 0,
  });

  s.addNotes(
    "[15초]\n" +
    "사진첩에 잠들어 있는 여행을 영상 한 편으로 남기는 것이 One Pick Story입니다.\n" +
    "다음 단계로는 지인이 아닌 실사용자를 확보하고, 유료 전환을 실험해 보려고 합니다.\n" +
    "감사합니다.\n\n" +
    "※ 발표 직전에 링크를 한 번 열어 서버를 깨워 둘 것 — 무료 플랜은 15분 유휴 후 잠들어 첫 접속이 50초 이상 걸린다."
  );
}

pres.writeFile({ fileName: "/Users/songhayoon/Desktop/vibe_mvp/OnePickStory_demoday.pptx" })
  .then((f) => console.log("saved:", f));
