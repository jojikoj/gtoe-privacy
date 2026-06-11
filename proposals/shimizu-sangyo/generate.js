// 清水産業様 HPリニューアル企画書 v15 生成スクリプト
// 使い方: node generate.js → 清水産業様_HPリニューアル企画書_v15_2026-06.pptx を出力
const pptxgen = require("pptxgenjs");

const pptx = new pptxgen();
pptx.defineLayout({ name: "WIDE", width: 13.33, height: 7.5 });
pptx.layout = "WIDE";

// ---- カラーパレット（インダストリアル・ブルー × 警告イエロー）----
const NAVY = "0F2A52";
const NAVY2 = "1B3A6B";
const BLUE = "2060E0";
const YELLOW = "F5B800";
const LIGHT = "EEF3FA";
const GRAY = "6B7280";
const TEXT = "1A1A1A";
const WHITE = "FFFFFF";
const RED = "C0392B";
const GREEN = "1E8E5A";

const FONT = "Yu Gothic";
const W = 13.33;
const M = 0.6; // 左右マージン
const CW = W - M * 2; // コンテンツ幅

let pageNo = 0;

function baseSlide() {
  const s = pptx.addSlide();
  s.background = { color: WHITE };
  return s;
}

// 各スライド共通ヘッダー＋フッター
function contentSlide(section, title, sub) {
  pageNo++;
  const s = baseSlide();
  s.addShape("rect", { x: 0, y: 0, w: W, h: 0.16, fill: { color: NAVY } });
  s.addText(section, {
    x: M, y: 0.38, w: CW, h: 0.3, fontFace: FONT, fontSize: 12, bold: true,
    color: BLUE, charSpacing: 2,
  });
  s.addText(title, {
    x: M, y: 0.66, w: CW, h: 0.62, fontFace: FONT, fontSize: 27, bold: true, color: TEXT,
  });
  if (sub) {
    s.addText(sub, {
      x: M, y: 1.3, w: CW, h: 0.34, fontFace: FONT, fontSize: 13, color: GRAY,
    });
  }
  s.addText(`株式会社TOE ｜ 清水産業株式会社 様 ご提案 ｜ 2026年6月`, {
    x: M, y: 7.08, w: 8, h: 0.3, fontFace: FONT, fontSize: 9, color: GRAY,
  });
  s.addText(String(pageNo), {
    x: W - 1.0, y: 7.08, w: 0.5, h: 0.3, fontFace: FONT, fontSize: 10, color: GRAY, align: "right",
  });
  return s;
}

function card(s, x, y, w, h, opts = {}) {
  s.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: opts.fill || LIGHT },
    line: { color: opts.line || "D7E2F2", width: 0.75 },
  });
}

// ============================================================
// 1. 表紙
// ============================================================
(() => {
  const s = baseSlide();
  s.background = { color: NAVY };
  s.addShape("rect", { x: 0, y: 0, w: W, h: 0.14, fill: { color: YELLOW } });
  s.addText("SHIMIZU SANGYO ｜ CORPORATE SITE RENEWAL PROPOSAL", {
    x: M, y: 0.7, w: CW, h: 0.35, fontFace: FONT, fontSize: 13, bold: true,
    color: YELLOW, charSpacing: 3,
  });
  s.addText("止まったら、清水。", {
    x: M, y: 2.0, w: CW, h: 1.5, fontFace: FONT, fontSize: 60, bold: true, color: WHITE,
  });
  s.addText(
    "売る・直す・据える・つくる。\n4つの柱で、機械の困りごとを一社で止める“かかりつけ”へ。",
    { x: M, y: 3.6, w: CW, h: 0.95, fontFace: FONT, fontSize: 18, color: "C9D6EC", lineSpacing: 30 }
  );
  s.addShape("line", { x: M, y: 5.1, w: 4.5, h: 0, line: { color: YELLOW, width: 2 } });
  s.addText("清水産業株式会社 御中", {
    x: M, y: 5.45, w: CW, h: 0.45, fontFace: FONT, fontSize: 20, bold: true, color: WHITE,
  });
  s.addText("ホームページ リニューアル ご提案書 ／ 2026年6月\n株式会社TOE", {
    x: M, y: 6.05, w: CW, h: 0.75, fontFace: FONT, fontSize: 14, color: "C9D6EC", lineSpacing: 24,
  });
})();

// ============================================================
// 2. 現状診断
// ============================================================
(() => {
  const s = contentSlide("01 ｜ 現状サイト診断", "いまのサイトの「4つのもったいない」",
    "実力は十分。でも、その良さがWebで伝わっていません。直せば伸びる“伸びしろ”です。");
  const items = [
    ["01", "スマホで見づらい", "閲覧の約8割はスマホ。今のサイトはPC向けで、文字が小さく拡大しないと読めません。", "見づらさで、多くが離脱"],
    ["02", "何の会社か伝わらない", "商社・メンテ・配管・設計製造。4つの強みが「販売／修理」の2区分に埋もれています。", "強みが、選ばれていない"],
    ["03", "問い合わせされにくい", "連絡手段が電話中心。夜間・休日や、メンテ・OEMの引き合いを取りこぼしています。", "夜間・休日を取り逃す"],
    ["04", "「保護されていない通信」", "常時SSL非対応で、ブラウザに警告が表示。信用にも検索順位（SEO）にも実害があります。", "初見の信用を失う"],
  ];
  const cw = (CW - 0.3 * 3) / 4;
  items.forEach((it, i) => {
    const x = M + i * (cw + 0.3);
    card(s, x, 1.85, cw, 3.9);
    s.addText(it[0], { x: x + 0.18, y: 2.0, w: cw - 0.36, h: 0.45, fontFace: FONT, fontSize: 22, bold: true, color: BLUE });
    s.addText(it[1], { x: x + 0.18, y: 2.5, w: cw - 0.36, h: 0.75, fontFace: FONT, fontSize: 15, bold: true, color: TEXT });
    s.addText(it[2], { x: x + 0.18, y: 3.3, w: cw - 0.36, h: 1.6, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 18 });
    s.addText("→ " + it[3], { x: x + 0.18, y: 5.05, w: cw - 0.36, h: 0.55, fontFace: FONT, fontSize: 11.5, bold: true, color: RED });
  });
  s.addShape("roundRect", { x: M, y: 6.05, w: CW, h: 0.75, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("制作から8年。事業の中身は進化したのに、サイトだけが止まっています。── どれも、作り直せば解決します。", {
    x: M + 0.25, y: 6.05, w: CW - 0.5, h: 0.75, fontFace: FONT, fontSize: 14, bold: true, color: WHITE, valign: "middle",
  });
})();

// ============================================================
// 3. Before → After
// ============================================================
(() => {
  const s = contentSlide("02 ｜ リニューアルの効果", "リニューアルで、こう変わります",
    "「見づらい・伝わらない・問い合わせが来ない」を、ぜんぶ逆にします。");
  const rows = [
    ["スマホだと崩れて読みにくい", "どの端末でもキレイに表示。現場や外出先でも快適"],
    ["4つの事業がバラバラで分かりにくい", "「売る・直す・据える・つくる」の総合力が3秒で伝わる"],
    ["メンテの情報がほぼ無い", "整備事例・修理済在庫を前面に。「直せる会社」と一目で伝わる"],
    ["問い合わせは電話番号だけ", "症状から相談できるフォーム＋AIの24時間受付で取りこぼさない"],
    ["日本語のみで海外を取り逃す", "英語・繁体字（台湾）をサイト内できちんと表示"],
    ["更新のたびに業者へ依頼", "お知らせや実績を自分たちで即日更新できる"],
  ];
  let y = 1.9;
  rows.forEach((r) => {
    card(s, M, y, 5.5, 0.72, { fill: "F6F7F9", line: "E2E5EA" });
    s.addText(r[0], { x: M + 0.2, y, w: 5.1, h: 0.72, fontFace: FONT, fontSize: 12.5, color: GRAY, valign: "middle" });
    s.addText("▶", { x: M + 5.55, y, w: 0.5, h: 0.72, fontFace: FONT, fontSize: 14, bold: true, color: YELLOW, align: "center", valign: "middle" });
    card(s, M + 6.1, y, CW - 6.1, 0.72, { fill: LIGHT, line: "C9D9F0" });
    s.addText(r[1], { x: M + 6.3, y, w: CW - 6.5, h: 0.72, fontFace: FONT, fontSize: 12.5, bold: true, color: NAVY, valign: "middle" });
    y += 0.84;
  });
})();

// ============================================================
// 4. コンセプト
// ============================================================
(() => {
  const s = contentSlide("03 ｜ ご提案のコンセプト", "機械の“かかりつけ”、清水産業。",
    "バラバラに見えていた事業を、ひとつのメッセージにまとめます。");
  // 中央コンセプト
  s.addShape("roundRect", { x: M, y: 1.95, w: 3.6, h: 3.6, rectRadius: 0.08, fill: { color: NAVY } });
  s.addText("機械の\n困りごとは、\nぜんぶ清水で\n止める。", {
    x: M + 0.2, y: 2.15, w: 3.2, h: 3.2, fontFace: FONT, fontSize: 22, bold: true, color: WHITE, lineSpacing: 36, valign: "middle",
  });
  const pillars = [
    ["① 売る", "総合機械商社", "500社超のメーカーから機械・部品を調達。「これ扱ってる？」に即答。"],
    ["② 直す", "メンテナンス", "モーター・コンプレッサー等の整備・オーバーホール。今回の打ち出しの主役。"],
    ["③ 据える", "配管・据付工事", "プラント配管から機械の据付・試運転まで、現場を一貫対応。"],
    ["④ つくる", "機械設計・製造", "図面設計からOEM製造まで。「無ければ、つくる」で応える。"],
  ];
  const px = M + 3.9, pw = (CW - 3.9 - 0.25) / 2;
  pillars.forEach((p, i) => {
    const x = px + (i % 2) * (pw + 0.25);
    const y = 1.95 + Math.floor(i / 2) * 1.92;
    card(s, x, y, pw, 1.68);
    s.addText(p[0], { x: x + 0.2, y: y + 0.12, w: pw - 0.4, h: 0.4, fontFace: FONT, fontSize: 16, bold: true, color: BLUE });
    s.addText(p[1], { x: x + 0.2, y: y + 0.52, w: pw - 0.4, h: 0.35, fontFace: FONT, fontSize: 13, bold: true, color: TEXT });
    s.addText(p[2], { x: x + 0.2, y: y + 0.88, w: pw - 0.4, h: 0.72, fontFace: FONT, fontSize: 11, color: GRAY, lineSpacing: 16 });
  });
  s.addShape("roundRect", { x: M, y: 5.85, w: CW, h: 0.95, rectRadius: 0.06, fill: { color: "FFF7E0" }, line: { color: YELLOW, width: 1 } });
  s.addText(
    "ポイント：現サイトにほぼ無い「メンテナンス」を最重要コンテンツに格上げ。整備事例と修理済在庫で、“直せる・止めない会社”を前面に打ち出します。",
    { x: M + 0.25, y: 5.85, w: CW - 0.5, h: 0.95, fontFace: FONT, fontSize: 13, bold: true, color: TEXT, valign: "middle", lineSpacing: 20 }
  );
})();

// ============================================================
// 5. サイト全体構成
// ============================================================
(() => {
  const s = contentSlide("04 ｜ サイト全体構成", "トップで「総合力」、4本柱へ迷わず案内",
    "リクルートページは既存の採用サイトへリンクし、二重管理を防ぎます。");
  // トップ
  s.addShape("roundRect", { x: M + 2.5, y: 1.85, w: CW - 5.0, h: 0.7, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("トップページ（4本柱の入口 ＋ 24時間受付 ＋ 実績）", {
    x: M + 2.5, y: 1.85, w: CW - 5.0, h: 0.7, fontFace: FONT, fontSize: 14, bold: true, color: WHITE, align: "center", valign: "middle",
  });
  const biz = [
    ["機械の販売", "総合商社\n500社超のメーカー\n機械・部品の調達"],
    ["修理・メンテ", "オーバーホール\n整備事例の公開\n修理済在庫＝即納"],
    ["配管・据付", "プラント配管工事\n機械の設置・施工\n試運転まで一貫"],
    ["設計・製造", "図面設計・OEM\n自社オリジナル機械\n「無ければつくる」"],
  ];
  const bw = (CW - 0.3 * 3) / 4;
  biz.forEach((b, i) => {
    const x = M + i * (bw + 0.3);
    s.addShape("line", { x: x + bw / 2, y: 2.55, w: 0, h: 0.3, line: { color: GRAY, width: 1 } });
    card(s, x, 2.85, bw, 1.75, { fill: LIGHT, line: "C9D9F0" });
    s.addText(b[0], { x: x + 0.15, y: 2.98, w: bw - 0.3, h: 0.4, fontFace: FONT, fontSize: 14.5, bold: true, color: NAVY });
    s.addText(b[1], { x: x + 0.15, y: 3.4, w: bw - 0.3, h: 1.1, fontFace: FONT, fontSize: 11, color: TEXT, lineSpacing: 17 });
  });
  const others = ["会社案内", "実績・事例", "取扱メーカー", "採用 → 既存採用サイトへリンク", "お問い合わせ", "言語切替（英語・繁体字）"];
  s.addText("そのほかのページ", { x: M, y: 4.85, w: CW, h: 0.35, fontFace: FONT, fontSize: 13, bold: true, color: TEXT });
  const ow = (CW - 0.25 * 5) / 6;
  others.forEach((o, i) => {
    const x = M + i * (ow + 0.25);
    card(s, x, 5.25, ow, 0.85, { fill: "F6F7F9", line: "E2E5EA" });
    s.addText(o, { x: x + 0.08, y: 5.25, w: ow - 0.16, h: 0.85, fontFace: FONT, fontSize: 10.5, color: TEXT, align: "center", valign: "middle", lineSpacing: 14 });
  });
  s.addText("※ 採用は既存リクルートサイトを活かしてリンク接続。将来のリニューアルはPhase2以降でご相談可能です。", {
    x: M, y: 6.3, w: CW, h: 0.35, fontFace: FONT, fontSize: 11, color: GRAY,
  });
})();

// ============================================================
// 6. メンテを武器に
// ============================================================
(() => {
  const s = contentSlide("05 ｜ 体験① メンテナンスの可視化", "目に見えない“整備の実力”を、証拠にする",
    "症状から相談 → 診断 → オーバーホール → 再稼働まで、Before/Afterで実証します。");
  const steps = ["症状を選ぶ\n（異音／発熱／漏れ）", "型番・写真を\nアップロード", "整備士が\n診断・見積", "オーバーホール\n→ 再稼働"];
  const sw = (CW - 0.5 * 3) / 4;
  steps.forEach((st, i) => {
    const x = M + i * (sw + 0.5);
    card(s, x, 1.9, sw, 1.15, { fill: LIGHT, line: "C9D9F0" });
    s.addText(`${i + 1}`, { x: x + 0.12, y: 1.98, w: 0.5, h: 0.4, fontFace: FONT, fontSize: 18, bold: true, color: BLUE });
    s.addText(st, { x: x + 0.12, y: 2.0, w: sw - 0.24, h: 1.0, fontFace: FONT, fontSize: 11.5, bold: true, color: TEXT, align: "center", valign: "middle", lineSpacing: 16 });
    if (i < 3) s.addText("→", { x: x + sw + 0.05, y: 1.9, w: 0.4, h: 1.15, fontFace: FONT, fontSize: 18, bold: true, color: YELLOW, align: "center", valign: "middle" });
  });
  const feats = [
    ["メーカー問わず対応", "特定ブランドに縛られず、ポンプ・モーター・コンプレッサー等の回転機械を幅広く整備。"],
    ["分解整備の“見える化”", "分解→洗浄→部品交換→再組立を、Before/After写真で実証。カタログに無い説得力。"],
    ["整備事例ページが“営業”になる", "「どんな故障を、どう直したか」を蓄積。同じ悩みの人が検索で見つけ、相談につながる。"],
  ];
  const fw = (CW - 0.3 * 2) / 3;
  feats.forEach((f, i) => {
    const x = M + i * (fw + 0.3);
    card(s, x, 3.4, fw, 1.7);
    s.addText(f[0], { x: x + 0.2, y: 3.55, w: fw - 0.4, h: 0.55, fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY });
    s.addText(f[1], { x: x + 0.2, y: 4.12, w: fw - 0.4, h: 0.9, fontFace: FONT, fontSize: 11, color: TEXT, lineSpacing: 17 });
  });
  s.addShape("roundRect", { x: M, y: 5.45, w: CW, h: 1.25, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("1,892万円", { x: M + 0.35, y: 5.55, w: 3.2, h: 0.6, fontFace: FONT, fontSize: 26, bold: true, color: YELLOW });
  s.addText("製造業の突発停止による年間平均損失（1社あたり）", { x: M + 0.35, y: 6.15, w: 3.6, h: 0.4, fontFace: FONT, fontSize: 10, color: "C9D6EC" });
  s.addText("「止めない・また動かす」価値を数字で語り、メンテの引き合いを生むページ群に。", {
    x: M + 4.3, y: 5.45, w: CW - 4.6, h: 1.25, fontFace: FONT, fontSize: 14, bold: true, color: WHITE, valign: "middle", lineSpacing: 22,
  });
  s.addText("出典：八千代ソリューションズ 突発停止損失調査（500社）", { x: M, y: 6.75, w: CW, h: 0.3, fontFace: FONT, fontSize: 9, color: GRAY });
})();

// ============================================================
// 7. 目玉：修理済在庫品
// ============================================================
(() => {
  const s = contentSlide("06 ｜ 体験② ご提案の目玉", "「修理済在庫品」を、即納カタログに変える",
    "現サイトで空欄のままのページが、御社にしか出せない最大の武器になります。");
  const flow = ["ポンプ・モーター\nが止まった", "在庫を型番・症状\nで検索（AI補助）", "整備可否＋即納可\nをその場で回答", "交換機を即出荷\n外した機械は整備へ"];
  const fw = (CW - 0.5 * 3) / 4;
  flow.forEach((f, i) => {
    const x = M + i * (fw + 0.5);
    card(s, x, 1.9, fw, 1.2, { fill: i === 3 ? "FFF7E0" : LIGHT, line: i === 3 ? YELLOW : "C9D9F0" });
    s.addText(f, { x: x + 0.12, y: 1.9, w: fw - 0.24, h: 1.2, fontFace: FONT, fontSize: 11.5, bold: true, color: TEXT, align: "center", valign: "middle", lineSpacing: 17 });
    if (i < 3) s.addText("→", { x: x + fw + 0.05, y: 1.9, w: 0.4, h: 1.2, fontFace: FONT, fontSize: 18, bold: true, color: YELLOW, align: "center", valign: "middle" });
  });
  card(s, M, 3.45, (CW - 0.3) / 2, 1.85, { fill: "FBF1F0", line: "E5C3BE" });
  s.addText("いま（在庫が見えない）", { x: M + 0.2, y: 3.6, w: 5.5, h: 0.4, fontFace: FONT, fontSize: 13.5, bold: true, color: RED });
  s.addText("在庫一覧ページは“メーカー名―／型番―”の空テンプレのまま。どんな在庫があるか客には分からず、電話で都度確認。夜間・休日は受けられず、止まったお客を取りこぼす。", {
    x: M + 0.2, y: 4.05, w: (CW - 0.3) / 2 - 0.4, h: 1.1, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 18,
  });
  const x2 = M + (CW - 0.3) / 2 + 0.3;
  card(s, x2, 3.45, (CW - 0.3) / 2, 1.85, { fill: "EAF6F0", line: "BCDCCB" });
  s.addText("これから（即納在庫が見える）", { x: x2 + 0.2, y: 3.6, w: 5.5, h: 0.4, fontFace: FONT, fontSize: 13.5, bold: true, color: GREEN });
  s.addText("即納できる整備済機をWebで公開。機種・メーカー・能力で検索でき、24時間取りこぼさない。「在庫、あります」をサイトで言い切り、“止めない会社”を証明する。", {
    x: x2 + 0.2, y: 4.05, w: (CW - 0.3) / 2 - 0.4, h: 1.1, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 18,
  });
  s.addShape("roundRect", { x: M, y: 5.6, w: CW, h: 1.05, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("なぜ清水産業だけができるのか ── 整備力 × 即納在庫 × 検索・AI の3点を同時に持つ会社は他にありません。メーカー特約店にも、街の修理屋にも、ネット通販にも真似できない武器です。", {
    x: M + 0.3, y: 5.6, w: CW - 0.6, h: 1.05, fontFace: FONT, fontSize: 13, bold: true, color: WHITE, valign: "middle", lineSpacing: 21,
  });
})();

// ============================================================
// 8. AI
// ============================================================
(() => {
  const s = contentSlide("07 ｜ 体験③ TOEのAI", "AIが、相談のハードルを下げる",
    "深夜でも、海外からでも。AIがまず受けて、必要なら整備士へつなぎます。");
  const ai = [
    ["AI 24時間 一次受付", "深夜・休日もAIが症状を聞き取り、緊急修理・見積・在庫照会へ自動で振り分け。電話に出られない時間も取りこぼさない。", "目玉"],
    ["AI 部品・在庫特定", "型番や症状、銘板の写真から、即納在庫や後継・代替機の当たりをAIが提示。「型番が分からない」相談にも応えられる。", "目玉"],
    ["AI 多言語一次対応", "英語・繁体字の問い合わせもAIが一次対応。海外・外資系工場からの引き合いに備える。", ""],
  ];
  const aw = (CW - 0.3 * 2) / 3;
  ai.forEach((a, i) => {
    const x = M + i * (aw + 0.3);
    card(s, x, 1.9, aw, 2.45);
    if (a[2]) {
      s.addShape("roundRect", { x: x + aw - 1.0, y: 2.02, w: 0.8, h: 0.32, rectRadius: 0.06, fill: { color: YELLOW } });
      s.addText(a[2], { x: x + aw - 1.0, y: 2.02, w: 0.8, h: 0.32, fontFace: FONT, fontSize: 10, bold: true, color: NAVY, align: "center", valign: "middle" });
    }
    s.addText(a[0], { x: x + 0.2, y: 2.05, w: aw - 1.1, h: 0.7, fontFace: FONT, fontSize: 14.5, bold: true, color: NAVY, lineSpacing: 19 });
    s.addText(a[1], { x: x + 0.2, y: 2.8, w: aw - 0.4, h: 1.45, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 18 });
  });
  card(s, M, 4.6, 7.4, 2.0, { fill: "F6F7F9", line: "E2E5EA" });
  s.addText("安心して任せられるAI設計", { x: M + 0.2, y: 4.72, w: 7.0, h: 0.4, fontFace: FONT, fontSize: 13.5, bold: true, color: TEXT });
  s.addText(
    "・御社専用環境で処理し、入力内容を学習に使わせない\n・AIは候補提示まで。見積・最終判断は必ず整備士が行う\n・分からないことは「分からない」と正直に答えさせる\n・チャット受付 → 部品特定 → 多言語、と段階導入で無理なく",
    { x: M + 0.2, y: 5.12, w: 7.0, h: 1.4, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 19 }
  );
  s.addShape("roundRect", { x: M + 7.7, y: 4.6, w: CW - 7.7, h: 2.0, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("なぜTOEか", { x: M + 7.9, y: 4.72, w: 4.0, h: 0.4, fontFace: FONT, fontSize: 13.5, bold: true, color: YELLOW });
  s.addText("製造業AI「YAI」・介護記録AI「CARE VOICE」を実運用まで作り切った実績。デモで終わらせず、御社の業務に本当に使えるAIを実装します。", {
    x: M + 7.9, y: 5.12, w: CW - 8.1, h: 1.4, fontFace: FONT, fontSize: 11.5, color: WHITE, lineSpacing: 19,
  });
})();

// ============================================================
// 9. 多言語
// ============================================================
(() => {
  const s = contentSlide("08 ｜ 多言語対応", "英語・繁体字（台湾）を、サイト内で正しく表示",
    "Google翻訳まかせをやめて、海外からの引き合いを受け止める体制に。");
  card(s, M, 1.9, 7.4, 2.6);
  s.addText("Weglot による自動多言語化", { x: M + 0.2, y: 2.05, w: 7.0, h: 0.45, fontFace: FONT, fontSize: 15, bold: true, color: NAVY });
  s.addText(
    "・タグを1つ埋め込むだけで、英語・繁体字を自動展開\n・日本語の本体を更新すれば、外国語ページも自動で追従\n・AI自動翻訳＋Web上で人手修正。公開前にネイティブチェック\n・初期費用 ¥0 ／ 月額 ¥7,500〜（プランにより変動）",
    { x: M + 0.2, y: 2.55, w: 7.0, h: 1.8, fontFace: FONT, fontSize: 12.5, color: TEXT, lineSpacing: 22 }
  );
  s.addShape("roundRect", { x: M + 7.7, y: 1.9, w: CW - 7.7, h: 2.6, rectRadius: 0.06, fill: { color: LIGHT }, line: { color: "C9D9F0", width: 0.75 } });
  s.addText("対応言語", { x: M + 7.9, y: 2.05, w: 4.0, h: 0.4, fontFace: FONT, fontSize: 13, bold: true, color: TEXT });
  s.addText("英語（EN）／ 繁体字中国語（台湾）", { x: M + 7.9, y: 2.45, w: CW - 8.1, h: 0.45, fontFace: FONT, fontSize: 14, bold: true, color: NAVY });
  s.addText("専門用語（機種名・部品名）は用語集で固定し、取り違えを防ぎます。必要に応じて言語の追加も可能です。", {
    x: M + 7.9, y: 2.95, w: CW - 8.1, h: 1.4, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 19,
  });
  s.addShape("roundRect", { x: M, y: 4.85, w: CW, h: 1.7, rectRadius: 0.06, fill: { color: "FFF7E0" }, line: { color: YELLOW, width: 1 } });
  s.addText("多言語 × AI受付 × 海外専用フォーム", { x: M + 0.25, y: 5.0, w: CW - 0.5, h: 0.45, fontFace: FONT, fontSize: 14, bold: true, color: TEXT });
  s.addText("ページを翻訳するだけでなく、問い合わせの受け皿（海外専用フォーム＋AIの英語・繁体字一次対応）までセットで設計。OEM・装置設計は海外引き合いが伸びる領域です。台湾・アジア商社からの相談を、取りこぼさず日本語に整えて御社へ届けます。", {
    x: M + 0.25, y: 5.45, w: CW - 0.5, h: 1.0, fontFace: FONT, fontSize: 12, color: TEXT, lineSpacing: 20,
  });
})();

// ============================================================
// 10. デザイン＋撮影
// ============================================================
(() => {
  const s = contentSlide("09 ｜ デザインと撮影", "“プロの現場”が伝わるデザインと写真",
    "デザインの土台は撮影。現場のリアルでしか伝わらない信頼があります。");
  const tone = [
    ["配色", "インダストリアル・ブルー基調。重厚な信頼感と“現場の即応”を両立。"],
    ["写真", "機械の質感・整備士の手元まで伝わる、作り物でないリアル。"],
    ["言葉", "専門用語に逃げず、「止まった」「また動く」で語る平易さ。"],
  ];
  tone.forEach((t, i) => {
    const x = M + i * ((CW - 0.6) / 3 + 0.3);
    const w = (CW - 0.6) / 3;
    card(s, x, 1.9, w, 1.3, { fill: LIGHT, line: "C9D9F0" });
    s.addText(t[0], { x: x + 0.2, y: 2.0, w: w - 0.4, h: 0.4, fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY });
    s.addText(t[1], { x: x + 0.2, y: 2.4, w: w - 0.4, h: 0.75, fontFace: FONT, fontSize: 11, color: TEXT, lineSpacing: 17 });
  });
  s.addText("撮影プラン（1日完結・TOEプロデュース・お見積に込み）", {
    x: M, y: 3.45, w: CW, h: 0.4, fontFace: FONT, fontSize: 14.5, bold: true, color: TEXT,
  });
  const shots = [
    ["現場の人物", "整備士・配管職人の手元と横顔。事業ページ＋信頼の核に"],
    ["ビフォーアフター", "整備前後の対比。実績ページの説得材料に"],
    ["工場・倉庫", "設備・在庫の規模が伝わる広角。商社・在庫ページの根拠に"],
    ["対応機器・部品", "取扱品・OEM製品。対応範囲の広さを可視化"],
    ["経営層・若手", "代表・現場リーダーのポートレート。会社紹介に"],
    ["外観・全景", "本社・拠点の俯瞰。トップのヒーロー画像候補"],
  ];
  const shw = (CW - 0.25 * 2) / 3;
  shots.forEach((sh, i) => {
    const x = M + (i % 3) * (shw + 0.25);
    const y = 3.95 + Math.floor(i / 3) * 1.18;
    card(s, x, y, shw, 1.02, { fill: "F6F7F9", line: "E2E5EA" });
    s.addText(sh[0], { x: x + 0.18, y: y + 0.08, w: shw - 0.36, h: 0.35, fontFace: FONT, fontSize: 12.5, bold: true, color: NAVY });
    s.addText(sh[1], { x: x + 0.18, y: y + 0.43, w: shw - 0.36, h: 0.55, fontFace: FONT, fontSize: 10.5, color: GRAY, lineSpacing: 15 });
  });
  s.addText("※ 撮影しないと「メーカー品を売っているだけの会社」に見えてしまう。現場・人・設備は、写真でしか伝えられません。", {
    x: M, y: 6.4, w: CW, h: 0.35, fontFace: FONT, fontSize: 11, color: GRAY,
  });
})();

// ============================================================
// 11. 体制
// ============================================================
(() => {
  const s = contentSlide("10 ｜ 制作体制", "少人数だから、速い・伝わる・無駄がない",
    "TOE 2名＋必要時パートナー。中間マージンゼロを体制で担保します。");
  s.addShape("roundRect", { x: M + 2.0, y: 1.9, w: CW - 4.0, h: 0.65, rectRadius: 0.06, fill: { color: LIGHT }, line: { color: "C9D9F0", width: 0.75 } });
  s.addText("清水産業株式会社 様（ご担当者様＋経営層）", {
    x: M + 2.0, y: 1.9, w: CW - 4.0, h: 0.65, fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY, align: "center", valign: "middle",
  });
  const team = [
    ["プロデューサー／ディレクター", "顧客窓口、要件定義、デザイン・コピー監修、進行管理、納品まで一気通貫。"],
    ["実装リード", "サイト構築、CMS設計、AI・多言語連携、QA・公開作業、保守運用基盤の整備。"],
  ];
  team.forEach((t, i) => {
    const x = M + 0.8 + i * ((CW - 1.6 - 0.4) / 2 + 0.4);
    const w = (CW - 1.6 - 0.4) / 2;
    s.addShape("line", { x: x + w / 2, y: 2.55, w: 0, h: 0.3, line: { color: GRAY, width: 1 } });
    card(s, x, 2.85, w, 1.5, { fill: NAVY === "0F2A52" ? LIGHT : LIGHT, line: "C9D9F0" });
    s.addText(t[0], { x: x + 0.2, y: 2.98, w: w - 0.4, h: 0.45, fontFace: FONT, fontSize: 13.5, bold: true, color: NAVY });
    s.addText(t[1], { x: x + 0.2, y: 3.45, w: w - 0.4, h: 0.8, fontFace: FONT, fontSize: 11.5, color: TEXT, lineSpacing: 18 });
  });
  s.addText("TOE（2名）＋ 必要時パートナー（TOEがディレクション・品質管理）", {
    x: M, y: 4.55, w: CW, h: 0.4, fontFace: FONT, fontSize: 13, bold: true, color: TEXT,
  });
  const partners = ["撮影カメラマン", "翻訳ネイティブチェック", "ライター（必要時）"];
  partners.forEach((p, i) => {
    const x = M + i * ((CW - 0.6) / 3 + 0.3);
    card(s, x, 5.0, (CW - 0.6) / 3, 0.6, { fill: "F6F7F9", line: "E2E5EA" });
    s.addText(p, { x, y: 5.0, w: (CW - 0.6) / 3, h: 0.6, fontFace: FONT, fontSize: 12, color: TEXT, align: "center", valign: "middle" });
  });
  s.addShape("roundRect", { x: M, y: 5.85, w: CW, h: 0.8, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("① 中間マージンゼロ　② 決裁者直結で判断が速い　③ 伝言ゲームなしで品質劣化なし", {
    x: M + 0.25, y: 5.85, w: CW - 0.5, h: 0.8, fontFace: FONT, fontSize: 13.5, bold: true, color: WHITE, valign: "middle",
  });
})();

// ============================================================
// 12. スケジュール
// ============================================================
(() => {
  const s = contentSlide("11 ｜ スケジュール", "今年つくり、来年さらに伸ばす",
    "2026年内に新サイトを公開し、2027年はAI・在庫検索で“集まるサイト”に育てます。");
  const phases = [
    ["2026年7月", "キックオフ・要件定義", "ヒアリング／現サイト・競合分析／サイト構成・要件の確定", LIGHT],
    ["8〜9月", "設計・デザイン", "ワイヤーフレーム／デザイン制作／撮影（1日）／原稿整理", LIGHT],
    ["10〜11月", "実装・多言語", "CMS構築・コーディング／常時SSL化／英語・繁体字展開／テスト", LIGHT],
    ["12月", "公開", "最終確認・公開作業／SEO・計測初期設定／運用レクチャー", "FFF7E0"],
    ["2027年〜", "Phase 2：育てる", "AI一次受付／修理済在庫検索／顧客マイページ／月次レポートで改善", "EAF6F0"],
  ];
  const pw = (CW - 0.25 * 4) / 5;
  phases.forEach((p, i) => {
    const x = M + i * (pw + 0.25);
    s.addShape("roundRect", { x, y: 1.95, w: pw, h: 0.55, rectRadius: 0.06, fill: { color: i === 4 ? GREEN : NAVY } });
    s.addText(p[0], { x, y: 1.95, w: pw, h: 0.55, fontFace: FONT, fontSize: 12.5, bold: true, color: WHITE, align: "center", valign: "middle" });
    card(s, x, 2.62, pw, 2.5, { fill: p[3], line: "D7E2F2" });
    s.addText(p[1], { x: x + 0.15, y: 2.75, w: pw - 0.3, h: 0.65, fontFace: FONT, fontSize: 13, bold: true, color: TEXT, lineSpacing: 17 });
    s.addText(p[2], { x: x + 0.15, y: 3.45, w: pw - 0.3, h: 1.55, fontFace: FONT, fontSize: 10.5, color: TEXT, lineSpacing: 16 });
  });
  s.addShape("roundRect", { x: M, y: 5.5, w: CW, h: 1.15, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("年内公開 → 来年強化", { x: M + 0.35, y: 5.62, w: 3.6, h: 0.55, fontFace: FONT, fontSize: 18, bold: true, color: YELLOW });
  s.addText("「今年から来年にかけて」のご要望に合わせた二段構え。まず年内に“顔”を一新し、2027年は問い合わせを増やす機能を段階投入。投資を分散しながら、効果を毎月の数字で確認できます。", {
    x: M + 4.1, y: 5.5, w: CW - 4.4, h: 1.15, fontFace: FONT, fontSize: 12, color: WHITE, valign: "middle", lineSpacing: 19,
  });
})();

// ============================================================
// 13. 概算見積
// ============================================================
(() => {
  const s = contentSlide("12 ｜ 概算お見積り", "撮影費込み・Phase 1 一式 ¥2,900,000（税別）",
    "要件確定後に正式お見積りをご提示します。Phase 2は効果を見ながら段階導入できます。");
  const rows = [
    [{ text: "項目", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
     { text: "内容", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
     { text: "概算金額（税別）", options: { bold: true, color: WHITE, fill: { color: NAVY }, align: "right" } }],
    ["企画・設計・ディレクション", "要件定義／サイト構成設計／進行管理・品質管理", "¥400,000"],
    ["デザイン", "トップ＋下層テンプレート／スマホ最適化", "¥600,000"],
    ["CMS実装・コーディング", "全ページ実装／お知らせ・実績の自社更新機能／常時SSL化", "¥1,200,000"],
    ["写真撮影", "1日完結・6カテゴリ（出張費・データ納品込み）", "¥180,000"],
    ["多言語化（英語・繁体字）", "Weglot導入設定／用語集整備／ネイティブチェック", "¥150,000"],
    ["SEO・計測・公開作業", "構造化データ／GA4・Search Console設定／公開・リダイレクト", "¥370,000"],
    [{ text: "Phase 1 合計", options: { bold: true, fill: { color: "FFF7E0" } } },
     { text: "撮影費込み・リニューアル一式", options: { bold: true, fill: { color: "FFF7E0" } } },
     { text: "¥2,900,000", options: { bold: true, align: "right", fill: { color: "FFF7E0" }, color: "C0392B", fontSize: 14 } }],
  ];
  s.addTable(rows, {
    x: M, y: 1.95, w: CW, colW: [3.4, 6.13, 2.6],
    fontFace: FONT, fontSize: 11.5, color: TEXT,
    border: { type: "solid", color: "D7E2F2", pt: 0.75 },
    rowH: 0.42, valign: "middle", margin: 0.06,
  });
  const run = [
    ["保守・運用", "サーバー監視／セキュリティ更新／月次レポート（公開後6ヶ月）", "¥30,000〜/月"],
    ["多言語（Weglot）", "英語・繁体字の自動展開・追従", "¥7,500〜/月"],
    ["Phase 2（2027年〜）", "AI一次受付／修理済在庫検索／顧客マイページ", "効果を見ながら個別お見積り"],
  ];
  let y = 5.45;
  run.forEach((r) => {
    card(s, M, y, CW, 0.5, { fill: "F6F7F9", line: "E2E5EA" });
    s.addText(r[0], { x: M + 0.2, y, w: 3.0, h: 0.5, fontFace: FONT, fontSize: 11.5, bold: true, color: TEXT, valign: "middle" });
    s.addText(r[1], { x: M + 3.3, y, w: 6.0, h: 0.5, fontFace: FONT, fontSize: 11, color: GRAY, valign: "middle" });
    s.addText(r[2], { x: M + 9.4, y, w: 2.55, h: 0.5, fontFace: FONT, fontSize: 11.5, bold: true, color: NAVY, valign: "middle", align: "right" });
    y += 0.56;
  });
})();

// ============================================================
// 14. 公開後
// ============================================================
(() => {
  const s = contentSlide("13 ｜ 公開後の運用", "作って終わりにしない。数字で育てる",
    "効果を毎月の数字で確認し、改善を回し続けます。");
  const kpi = [
    ["月間問い合わせ", "10件", "30件", "公開12ヶ月後の目標"],
    ["海外問い合わせ", "0件", "8件", "英・繁体字＋専用フォーム"],
    ["平均滞在時間", "45秒", "2分15秒", "構成・導線の改善で"],
  ];
  const kw = (CW - 0.3 * 2) / 3;
  kpi.forEach((k, i) => {
    const x = M + i * (kw + 0.3);
    card(s, x, 1.9, kw, 1.85);
    s.addText(k[0], { x: x + 0.2, y: 2.02, w: kw - 0.4, h: 0.4, fontFace: FONT, fontSize: 13, bold: true, color: TEXT });
    s.addText([
      { text: k[1], options: { fontSize: 20, bold: true, color: GRAY } },
      { text: "  →  ", options: { fontSize: 16, color: YELLOW, bold: true } },
      { text: k[2], options: { fontSize: 26, bold: true, color: NAVY } },
    ], { x: x + 0.2, y: 2.45, w: kw - 0.4, h: 0.7 });
    s.addText(k[3], { x: x + 0.2, y: 3.25, w: kw - 0.4, h: 0.4, fontFace: FONT, fontSize: 10.5, color: GRAY });
  });
  s.addText("※ 目標値は業界平均CVR・同業改修事例からの試算。公開後の実数で毎月上書きします。", {
    x: M, y: 3.85, w: CW, h: 0.3, fontFace: FONT, fontSize: 9.5, color: GRAY,
  });
  const ops = [
    ["GA4＋Search Console", "流入経路・検索語・ページ別の成果を月次レビュー"],
    ["月次レポート", "公開後6ヶ月、改善提案を月3項目添えてご報告"],
    ["自社更新（CMS）", "お知らせ・整備事例・在庫情報を月2〜3時間で内製"],
    ["技術コラム・事例蓄積", "「ポンプ 修理 ○○」等の検索流入を育てる資産に"],
  ];
  const ow = (CW - 0.3 * 3) / 4;
  ops.forEach((o, i) => {
    const x = M + i * (ow + 0.3);
    card(s, x, 4.35, ow, 1.6, { fill: "F6F7F9", line: "E2E5EA" });
    s.addText(o[0], { x: x + 0.18, y: 4.48, w: ow - 0.36, h: 0.6, fontFace: FONT, fontSize: 12, bold: true, color: NAVY, lineSpacing: 16 });
    s.addText(o[1], { x: x + 0.18, y: 5.1, w: ow - 0.36, h: 0.75, fontFace: FONT, fontSize: 10.5, color: TEXT, lineSpacing: 16 });
  });
  s.addShape("roundRect", { x: M, y: 6.15, w: CW, h: 0.7, rectRadius: 0.06, fill: { color: NAVY } });
  s.addText("サイトは“もう一人の営業”。雇って終わりではなく、毎月育てて成果を出させます。", {
    x: M + 0.25, y: 6.15, w: CW - 0.5, h: 0.7, fontFace: FONT, fontSize: 13.5, bold: true, color: WHITE, valign: "middle",
  });
})();

// ============================================================
// 15. 次のステップ（裏表紙）
// ============================================================
(() => {
  const s = baseSlide();
  s.background = { color: NAVY };
  s.addShape("rect", { x: 0, y: 0, w: W, h: 0.14, fill: { color: YELLOW } });
  s.addText("NEXT STEP", { x: M, y: 0.8, w: CW, h: 0.4, fontFace: FONT, fontSize: 14, bold: true, color: YELLOW, charSpacing: 3 });
  s.addText("まず、年内公開へ。", { x: M, y: 1.3, w: CW, h: 0.9, fontFace: FONT, fontSize: 40, bold: true, color: WHITE });
  const steps = [
    ["1", "ご提案内容のご確認", "本企画書と概算見積の方向性をご確認ください。構成・予算の調整も柔軟に承ります。"],
    ["2", "キックオフ日程の調整", "7月中の着手で年内公開に間に合います。お打ち合わせはオンラインでも伺ってでも。"],
    ["3", "整備事例3件の素材ヒアリング", "写真＋ひと言でOK。御社の負担は素材提供と最終確認だけ。実装はTOEが全責任を持ちます。"],
  ];
  const sw = (CW - 0.4 * 2) / 3;
  steps.forEach((st, i) => {
    const x = M + i * (sw + 0.4);
    s.addShape("roundRect", { x, y: 2.7, w: sw, h: 2.3, rectRadius: 0.08, fill: { color: NAVY2 }, line: { color: "3A5688", width: 0.75 } });
    s.addText(st[0], { x: x + 0.25, y: 2.9, w: 1.0, h: 0.6, fontFace: FONT, fontSize: 28, bold: true, color: YELLOW });
    s.addText(st[1], { x: x + 0.25, y: 3.5, w: sw - 0.5, h: 0.5, fontFace: FONT, fontSize: 14.5, bold: true, color: WHITE });
    s.addText(st[2], { x: x + 0.25, y: 4.0, w: sw - 0.5, h: 0.9, fontFace: FONT, fontSize: 11, color: "C9D6EC", lineSpacing: 17 });
  });
  s.addShape("line", { x: M, y: 5.6, w: CW, h: 0, line: { color: "3A5688", width: 1 } });
  s.addText("株式会社TOE", { x: M, y: 5.85, w: CW, h: 0.5, fontFace: FONT, fontSize: 20, bold: true, color: WHITE });
  s.addText(
    "Web制作 ／ 業務システム ／ AI開発　｜　IT業界25年・1,000案件（福岡）\n本社：福岡県福岡市中央区渡辺通1-9-3 一丁目ビル203　｜　joe@gtoe.info　｜　www.gtoe.info",
    { x: M, y: 6.35, w: CW, h: 0.75, fontFace: FONT, fontSize: 12, color: "C9D6EC", lineSpacing: 20 }
  );
})();

pptx.writeFile({ fileName: "清水産業様_HPリニューアル企画書_v15_2026-06.pptx" }).then((f) => {
  console.log("生成完了:", f);
});
