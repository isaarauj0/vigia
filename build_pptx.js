/**
 * vigIA — Apresentação PowerPoint
 * 2º Hackathon Sintehctica 2026 — Linha 1
 */
const pptx = require('pptxgenjs');
const pres = new pptx();

// Canvas: LAYOUT_16x9 = 10" × 5.625"
const W = 10, H = 5.625;

// Brand colors (NO # prefix — pptxgenjs rule)
const C = {
  bg:      '070A12',
  blue:    '3861FF',
  violet:  '7C3BFF',
  white:   'F0F4FF',
  muted:   '8899BB',
  danger:  'FF4757',
  success: '2ED573',
  warn:    'FFC107',
  orange:  'FF8C00',
  card:    '0D1323',
  border:  '1A2340',
};

// ── Helpers ──────────────────────────────────────────
function addBg(s) {
  s.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: W, h: H,
    fill: { color: C.bg },
    line: { color: C.bg, pt: 0 }
  });
}

function badge(s, text) {
  s.addText(text, {
    x: 0.5, y: 0.32, w: 9, h: 0.22,
    fontSize: 8, bold: true, color: C.muted,
    fontFace: 'Arial', charSpacing: 1.5,
    align: 'center', isTextBox: true, margin: 0
  });
}

function divBar(s, x, y, w) {
  s.addShape(pres.ShapeType.rect, {
    x, y, w, h: 0.04,
    fill: { color: C.blue },
    line: { color: C.blue, pt: 0 }
  });
}

function card(s, x, y, w, h, fillColor, lineColor, radius) {
  s.addShape(pres.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color: fillColor || C.card },
    line: { color: lineColor || C.border, pt: 1 },
    rectRadius: radius || 0.08
  });
}

// ═══════════════════════════════════════════════════════
// SLIDE 1 — ABERTURA
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);

  // Logo: vig (white) + IA (blue)
  s.addText([
    { text: 'vig', options: { color: C.white } },
    { text: 'IA',  options: { color: C.blue  } }
  ], {
    x: 0.5, y: 0.42, w: 9, h: 1.35,
    fontSize: 88, bold: true, align: 'center',
    fontFace: 'Arial', isTextBox: true, margin: 0
  });

  // Tagline
  s.addText('A IA que chega antes do aluno desistir', {
    x: 0.5, y: 1.78, w: 9, h: 0.42,
    fontSize: 18, bold: true, color: C.white,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  // Divider
  divBar(s, 4.1, 2.28, 1.8);

  // Institutions
  s.addText('São Judas Polo Santo Amaro  ·  FADERGS  ·  Ecossistema Ânima', {
    x: 0.5, y: 2.4, w: 9, h: 0.28,
    fontSize: 11, color: C.muted,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  s.addText('Equipe vigIA  ·  Orientação: Prof.ª Mariana', {
    x: 0.5, y: 2.72, w: 9, h: 0.25,
    fontSize: 10, color: C.muted,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  // Event card
  card(s, 2.3, 3.1, 5.4, 0.98, C.card, C.border, 0.1);
  s.addText([
    { text: '2º Hackathon Sintehctica 2026', options: { color: C.blue, bold: true } },
    { text: '\nLinha 1 — Aprendizagem Hiperpersonalizada com IA e Neurociência', options: { color: C.muted } }
  ], {
    x: 2.3, y: 3.1, w: 5.4, h: 0.98,
    fontSize: 11.5, fontFace: 'Calibri', align: 'center',
    valign: 'middle', isTextBox: true, margin: 8
  });

  // Bottom credits
  s.addText('vigIA  ·  Pitch 2026', {
    x: 0.5, y: 5.2, w: 9, h: 0.22,
    fontSize: 8, color: '3A4A6A',
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  s.addNotes('Slide 1 — Abertura. Apresentar a equipe vigIA e o contexto do hackathon.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 2 — O PROBLEMA
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '02 — O PROBLEMA');

  // Big danger stat
  s.addText('1 em 4', {
    x: 0.5, y: 0.55, w: 9, h: 1.35,
    fontSize: 96, bold: true, color: C.danger,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  s.addText('alunos abandona o ensino superior no Brasil', {
    x: 0.5, y: 1.9, w: 9, h: 0.4,
    fontSize: 19, bold: true, color: C.white,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  s.addText('Fonte: INEP 2025', {
    x: 0.5, y: 2.32, w: 9, h: 0.2,
    fontSize: 9, color: C.muted, italic: true,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  // Story text
  s.addText('As instituições investem muito em captação — campanhas, bolsas, processos seletivos.\nMas é a retenção que fecha o ciclo. Um aluno retido gera receita recorrente, indica outros\ne dilui o custo de aquisição.', {
    x: 1.2, y: 2.6, w: 7.6, h: 0.72,
    fontSize: 11.5, color: C.muted,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  // Maria story card
  card(s, 1.2, 3.42, 7.6, 0.85, C.card, C.border, 0.08);
  s.addText('"Conheça Maria. Entrou motivada, passou na primeira prova, mas começou a faltar.\nMandou mensagem — não teve resposta. No terceiro mês, parou de aparecer."', {
    x: 1.2, y: 3.42, w: 7.6, h: 0.85,
    fontSize: 11, color: C.muted, italic: true,
    fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 8
  });

  // Quote
  s.addText('O que faltou para Maria?  Alguém chegar antes.', {
    x: 0.5, y: 4.42, w: 9, h: 0.45,
    fontSize: 15, bold: true, color: C.blue,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  s.addNotes('Slide 2 — O problema. 1 em 4 alunos abandona. A história da Maria ilustra a falha no sistema.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 3 — A SOLUÇÃO
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '03 — A SOLUÇÃO');

  s.addText('Plataforma modular com 5 frentes', {
    x: 0.5, y: 0.55, w: 9, h: 0.52,
    fontSize: 22, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  const modules = [
    { n: 'MÓD. 01', nm: 'Diagnóstico\nCognitivo',       d: 'Perfil do aluno no\nprimeiro acesso',   hl: false },
    { n: 'MÓD. 02', nm: 'Trilha\nPersonalizada',         d: 'Itinerário adaptado\nao perfil',         hl: false },
    { n: 'MÓD. 03', nm: 'Microlearning\nAdaptativo',     d: 'Conteúdo em\nmicro-doses precisas', hl: false },
    { n: 'MÓD. 04', nm: 'Simulados\nInteligentes',       d: 'Avaliação contextual\ncontinua',         hl: false },
    { n: 'MÓD. 05', nm: '⚡ Vigilância\nInstitucional',  d: 'Radar · Alertas\nautomáticos · Dia 1', hl: true  },
  ];

  const cw = 1.68, cgap = 0.16;
  const totalW = modules.length * cw + (modules.length - 1) * cgap;
  const sx0 = (W - totalW) / 2;
  const cy = 1.22, ch = 4.0;

  modules.forEach((m, i) => {
    const cx = sx0 + i * (cw + cgap);
    card(s, cx, cy, cw, ch,
      m.hl ? '14082B' : C.card,
      m.hl ? C.violet  : C.border, 0.1);

    s.addText(m.n, {
      x: cx + 0.05, y: cy + 0.18, w: cw - 0.1, h: 0.22,
      fontSize: 7.5, bold: true, color: C.muted,
      fontFace: 'Arial', charSpacing: 1, align: 'center', isTextBox: true, margin: 0
    });

    s.addText(m.nm, {
      x: cx + 0.05, y: cy + 0.45, w: cw - 0.1, h: 0.7,
      fontSize: 11, bold: true, color: m.hl ? C.violet : C.white,
      fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
    });

    s.addText(m.d, {
      x: cx + 0.08, y: cy + 1.3, w: cw - 0.16, h: 0.75,
      fontSize: 9.5, color: C.muted,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  s.addNotes('Slide 3 — A solução vigIA com 5 módulos. O Módulo 5 (Vigilância Institucional) é o diferencial do pitch.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 4 — COMO O RADAR FUNCIONA
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '04 — COMO O RADAR FUNCIONA');

  s.addText('O aluno não vê. Não preenche. Não faz nada.', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 20, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  const steps = [
    { icon: '🎓', label: 'MATRÍCULA',    sub: 'Ativa o\nmonitoramento',     bc: C.blue,    ec: '0B1530' },
    { icon: '📊', label: 'LMS/ERP/AVA',  sub: 'Dados alimentados\nautomaticamente', bc: C.violet, ec: '14082B' },
    { icon: '⚙️', label: 'SCORE',         sub: '5 dimensões\nem tempo real',   bc: C.warn,    ec: '2B2200' },
    { icon: '🚨', label: 'ALERTA',       sub: 'Limiar atingido\nautomaticamente', bc: C.danger, ec: '2A0A0E' },
    { icon: '✅', label: 'CANAL CERTO',  sub: 'Roteamento\nsem triagem',       bc: C.success, ec: '0D2B1A' },
  ];

  const sw = 1.38, sarr = 0.3, sgap = 0.12;
  const totalSW = steps.length * sw + (steps.length - 1) * (sarr + sgap);
  const sx0 = (W - totalSW) / 2;
  const sy = 1.2;

  steps.forEach((st, i) => {
    const sx = sx0 + i * (sw + sarr + sgap);

    // Circle
    s.addShape(pres.ShapeType.ellipse, {
      x: sx + (sw - 0.58) / 2, y: sy, w: 0.58, h: 0.58,
      fill: { color: st.ec },
      line: { color: st.bc, pt: 2 }
    });
    s.addText(st.icon, {
      x: sx + (sw - 0.58) / 2, y: sy + 0.02, w: 0.58, h: 0.52,
      fontSize: 18, align: 'center', valign: 'middle',
      isTextBox: true, margin: 0
    });

    // Arrow (not after last)
    if (i < steps.length - 1) {
      s.addText('→', {
        x: sx + sw + 0.02, y: sy + 0.06, w: sarr + sgap - 0.02, h: 0.46,
        fontSize: 14, color: C.muted,
        fontFace: 'Arial', align: 'center', valign: 'middle',
        isTextBox: true, margin: 0
      });
    }

    // Label
    s.addText(st.label, {
      x: sx, y: sy + 0.68, w: sw, h: 0.26,
      fontSize: 8.5, bold: true, color: C.white,
      fontFace: 'Arial', charSpacing: 0.5, align: 'center', isTextBox: true, margin: 0
    });

    // Sub
    s.addText(st.sub, {
      x: sx, y: sy + 0.95, w: sw, h: 0.42,
      fontSize: 8, color: C.muted,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  // Info box
  card(s, 0.75, 2.6, 8.5, 0.45, '0B1530', '1E3870', 0.07);
  s.addText('🔒  O radar é da instituição — o aluno não sabe que está sendo monitorado.', {
    x: 0.75, y: 2.6, w: 8.5, h: 0.45,
    fontSize: 12, bold: true, color: C.blue,
    fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });

  // Users chips
  s.addText('Quem usa o vigIA:', {
    x: 0.5, y: 3.18, w: 9, h: 0.24,
    fontSize: 9.5, color: C.muted,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  const chips = [
    { t: '👨‍🏫 Professor',       fill: C.card,   line: C.border, col: C.white  },
    { t: '🧑‍💼 Tutor',            fill: C.card,   line: C.border, col: C.white  },
    { t: '🏛 Coordenação',      fill: C.card,   line: C.border, col: C.white  },
    { t: '💳 Setor Financeiro', fill: C.card,   line: C.border, col: C.white  },
    { t: '✕ Nunca: o aluno',   fill: '2A0A0E', line: 'AA1122', col: C.danger },
  ];
  const chipW = 1.55, chipGap = 0.15;
  const totalChipW = chips.length * chipW + (chips.length - 1) * chipGap;
  const chipX0 = (W - totalChipW) / 2;

  chips.forEach((ch, i) => {
    const cx = chipX0 + i * (chipW + chipGap);
    s.addShape(pres.ShapeType.roundRect, {
      x: cx, y: 3.52, w: chipW, h: 0.34,
      fill: { color: ch.fill }, line: { color: ch.line, pt: 1 }, rectRadius: 0.17
    });
    s.addText(ch.t, {
      x: cx, y: 3.52, w: chipW, h: 0.34,
      fontSize: 9, bold: true, color: ch.col,
      fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
    });
  });

  s.addNotes('Slide 4 — O radar funciona a partir da matrícula, 100% automático, sem intervenção do aluno.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 5 — PAINEL INSTITUCIONAL
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '05 — PAINEL INSTITUCIONAL');

  s.addText('6 momentos críticos — detectados automaticamente', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 20, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  const sinais = [
    { icon: '🎓', t: 'Ingresso no curso',       d: 'Monitoramento ativado na matrícula' },
    { icon: '📝', t: 'Primeira avaliação',       d: 'Baseline de desempenho estabelecido' },
    { icon: '📅', t: 'Mudança de semestre',      d: 'Transição com alto risco de evasão' },
    { icon: '📉', t: 'Queda de desempenho',      d: 'Score cai abaixo do limiar definido' },
    { icon: '💤', t: 'Baixa participação',       d: 'Ausência no LMS, AVA e entregas' },
    { icon: '🔄', t: 'Retorno após ausência',    d: 'Reengajamento monitorado' },
  ];

  const cols = 3, rows = 2;
  const cw = 2.85, ch = 1.52, gapX = 0.17, gapY = 0.17;
  const totalW = cols * cw + (cols - 1) * gapX;
  const x0 = (W - totalW) / 2;
  const y0 = 1.22;

  sinais.forEach((sn, i) => {
    const col = i % cols, row = Math.floor(i / cols);
    const sx = x0 + col * (cw + gapX);
    const sy = y0 + row * (ch + gapY);

    card(s, sx, sy, cw, ch, C.card, C.border, 0.09);

    s.addText(sn.icon, {
      x: sx, y: sy + 0.18, w: cw, h: 0.4,
      fontSize: 20, align: 'center', valign: 'middle', isTextBox: true, margin: 0
    });

    s.addText(sn.t, {
      x: sx + 0.1, y: sy + 0.65, w: cw - 0.2, h: 0.35,
      fontSize: 11, bold: true, color: C.white,
      fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
    });

    s.addText(sn.d, {
      x: sx + 0.1, y: sy + 1.02, w: cw - 0.2, h: 0.4,
      fontSize: 9.5, color: C.muted,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  s.addNotes('Slide 5 — Os 6 momentos críticos de risco de evasão, detectados automaticamente pelo vigIA.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 6 — RÉGUA DE MEDIDA
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '06 — RÉGUA DE MEDIDA');

  s.addText('Score Composto em 5 Dimensões', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 20, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  // 5 dimension chips
  const dims = ['📊 Notas', '📅 Faltas', '💳 Mensalidades', '💻 Acesso LMS/AVA', '📋 Entregas'];
  const dw = 1.62, dg = 0.18;
  const totalDW = dims.length * dw + (dims.length - 1) * dg;
  const dx0 = (W - totalDW) / 2;

  dims.forEach((d, i) => {
    const dx = dx0 + i * (dw + dg);
    s.addShape(pres.ShapeType.roundRect, {
      x: dx, y: 1.22, w: dw, h: 0.37,
      fill: { color: '0C1A3A' }, line: { color: '1E3870', pt: 1 }, rectRadius: 0.06
    });
    s.addText(d, {
      x: dx, y: 1.22, w: dw, h: 0.37,
      fontSize: 10, bold: true, color: C.blue,
      fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
    });
  });

  // 4 color zones
  const zones = [
    { label: 'VERDE',    range: '76–100', note: 'Monitoramento padrão',    fill: '0D2B1A', line: '1A5C32', col: C.success },
    { label: 'AMARELO',  range: '51–75',  note: 'Atenção preventiva',       fill: '2B2200', line: '5C4800', col: C.warn   },
    { label: 'LARANJA',  range: '41–50',  note: 'Intervenção recomendada',  fill: '2B1500', line: '5C2E00', col: C.orange },
    { label: 'VERMELHO', range: '0–40',   note: 'Intervenção imediata',     fill: '2A0A0E', line: 'AA1122', col: C.danger },
  ];

  const zw = 2.05, zg = 0.12;
  const totalZW = zones.length * zw + (zones.length - 1) * zg;
  const zx0 = (W - totalZW) / 2;

  zones.forEach((z, i) => {
    const zx = zx0 + i * (zw + zg);
    card(s, zx, 1.72, zw, 0.92, z.fill, z.line, 0.07);

    s.addText(z.label, {
      x: zx, y: 1.78, w: zw, h: 0.2,
      fontSize: 7.5, bold: true, color: z.col,
      fontFace: 'Arial', charSpacing: 1, align: 'center', isTextBox: true, margin: 0
    });
    s.addText(z.range, {
      x: zx, y: 1.98, w: zw, h: 0.38,
      fontSize: 24, bold: true, color: z.col,
      fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
    });
    s.addText(z.note, {
      x: zx + 0.05, y: 2.38, w: zw - 0.1, h: 0.22,
      fontSize: 8, color: z.col,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  // 2 threshold cards
  const thrs = [
    { score: '≤ 40', label: 'ALERTA IMEDIATO',         note: 'Canal responsável notificado automaticamente', fill: '2B2200', line: C.warn,   col: C.warn   },
    { score: '≤ 25', label: 'INTERVENÇÃO PRIORITÁRIA', note: 'Todos os canais ativados simultaneamente',     fill: '2A0A0E', line: C.danger, col: C.danger },
  ];

  const tw = 4.3, tg = 0.3;
  const totalTW = thrs.length * tw + tg;
  const tx0 = (W - totalTW) / 2;

  thrs.forEach((t, i) => {
    const tx = tx0 + i * (tw + tg);
    card(s, tx, 2.78, tw, 1.22, t.fill, t.line, 0.09);

    s.addText(t.score, {
      x: tx, y: 2.88, w: tw, h: 0.56,
      fontSize: 38, bold: true, color: t.col,
      fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
    });
    s.addText(t.label, {
      x: tx + 0.1, y: 3.48, w: tw - 0.2, h: 0.26,
      fontSize: 10, bold: true, color: t.col,
      fontFace: 'Arial', charSpacing: 0.5, align: 'center', isTextBox: true, margin: 0
    });
    s.addText(t.note, {
      x: tx + 0.1, y: 3.76, w: tw - 0.2, h: 0.22,
      fontSize: 9, color: t.col,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  s.addNotes('Slide 6 — Régua de medida: score composto em 5 dimensões com 4 zonas de risco.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 7 — ROTEAMENTO INTELIGENTE
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '07 — ROTEAMENTO INTELIGENTE');

  s.addText('O problema define o canal — triagem automática, sem intervenção humana', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 17, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  const channels = [
    {
      ico: '💳', nm: 'CANAL FINANCEIRO', col: C.warn, fill: '2B2200', line: '5C4800',
      items: [
        ['⚠ Mensalidade 1 mês em atraso', '→ Consultor de Cobranças'],
        ['⚠ Mensalidade 2+ meses',         '→ Gerência Financeira'],
        ['⚠ Score financeiro < 40',        '→ Negociação prioritária'],
      ]
    },
    {
      ico: '📚', nm: 'CANAL ACADÊMICO', col: C.blue, fill: '0B1530', line: '1E3870',
      items: [
        ['⚠ Nota abaixo de 5,5',     '→ Professor da disciplina'],
        ['⚠ Faltas acima de 25%',    '→ Professor + Coordenação'],
        ['⚠ Entregas >25% faltando', '→ Tutor + Professor'],
      ]
    },
    {
      ico: '🎯', nm: 'CANAL DE RETENÇÃO', col: C.violet, fill: '170929', line: '3A1560',
      items: [
        ['⚠ Acesso LMS ausente',            '→ Tutor EAD'],
        ['⚠ Score geral < 25',              '→ Coord. + todos os canais'],
        ['⚠ Múltiplos alertas simultâneos', '→ Marketing de Retenção'],
      ]
    },
  ];

  const cw = 2.88, cg = 0.17;
  const totalCW = channels.length * cw + (channels.length - 1) * cg;
  const cx0 = (W - totalCW) / 2;
  const cy = 1.18, ch = 3.3;

  channels.forEach((ch, i) => {
    const cx = cx0 + i * (cw + cg);

    card(s, cx, cy, cw, ch, ch.fill, ch.line, 0.1);

    // Channel header
    s.addText(ch.ico + '  ' + ch.nm, {
      x: cx + 0.12, y: cy + 0.1, w: cw - 0.24, h: 0.36,
      fontSize: 9.5, bold: true, color: ch.col,
      fontFace: 'Arial', charSpacing: 0.5, align: 'left', isTextBox: true, margin: 0
    });

    // Header divider
    s.addShape(pres.ShapeType.rect, {
      x: cx, y: cy + 0.52, w: cw, h: 0.014,
      fill: { color: ch.line }, line: { color: ch.line, pt: 0 }
    });

    // Items
    ch.items.forEach((item, j) => {
      const iy = cy + 0.62 + j * 0.84;
      if (j > 0) {
        s.addShape(pres.ShapeType.rect, {
          x: cx + 0.1, y: iy - 0.06, w: cw - 0.2, h: 0.01,
          fill: { color: C.border }, line: { color: C.border, pt: 0 }
        });
      }
      s.addText(item[0], {
        x: cx + 0.12, y: iy + 0.04, w: cw - 0.24, h: 0.3,
        fontSize: 8.5, bold: true, color: C.white,
        fontFace: 'Calibri', align: 'left', isTextBox: true, margin: 0
      });
      s.addText(item[1], {
        x: cx + 0.12, y: iy + 0.36, w: cw - 0.24, h: 0.26,
        fontSize: 8.5, bold: false, color: ch.col,
        fontFace: 'Calibri', align: 'left', isTextBox: true, margin: 0
      });
    });
  });

  // Footer
  card(s, 0.5, 4.62, 9, 0.5, '170929', '3A1560', 0.07);
  s.addText('Um único score  ·  Três canais  ·  Triagem automática\nO responsável certo recebe o alerta certo, no momento certo.', {
    x: 0.5, y: 4.62, w: 9, h: 0.5,
    fontSize: 10, bold: true, color: C.violet,
    fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });

  s.addNotes('Slide 7 — Roteamento inteligente: um score, três canais, triagem 100% automática.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 8 — ALINHAMENTO PEDAGÓGICO
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '08 — ALINHAMENTO PEDAGÓGICO');

  s.addText('Taxonomia Ânima — 5 Princípios de Neuroaprendizagem', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 19, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  s.addText('O vigIA não é IA genérica — foi construído sobre a ciência da própria instituição.', {
    x: 0.5, y: 1.05, w: 9, h: 0.28,
    fontSize: 11.5, color: C.muted,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  const principles = [
    { icon: '🔁', n: 'Repetição\nEspaçada', d: 'Revisões nos momentos\nideais de consolidação' },
    { icon: '🧠', n: 'Aprendizagem\nAtiva',   d: 'Engajamento real,\nnão passivo' },
    { icon: '⚡', n: 'Microlearning',          d: 'Conteúdo em\ndoses precisas' },
    { icon: '💬', n: 'Feedback\nImediato',     d: 'Retorno no\nmomento exato' },
    { icon: '🎯', n: 'Personalização',         d: 'Cada aluno,\numa trilha única' },
  ];

  const pw = 1.65, pg = 0.18, ph = 2.22;
  const totalPW = principles.length * pw + (principles.length - 1) * pg;
  const px0 = (W - totalPW) / 2;
  const py = 1.42;

  principles.forEach((p, i) => {
    const px = px0 + i * (pw + pg);
    card(s, px, py, pw, ph, C.card, C.border, 0.1);

    s.addText(p.icon, {
      x: px, y: py + 0.25, w: pw, h: 0.5,
      fontSize: 22, align: 'center', valign: 'middle', isTextBox: true, margin: 0
    });
    s.addText(p.n, {
      x: px + 0.06, y: py + 0.82, w: pw - 0.12, h: 0.5,
      fontSize: 11, bold: true, color: C.white,
      fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
    });
    s.addText(p.d, {
      x: px + 0.06, y: py + 1.4, w: pw - 0.12, h: 0.6,
      fontSize: 9.5, color: C.muted,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  s.addText('Cada alerta, cada intervenção, cada trilha — segue a ciência da própria instituição.', {
    x: 0.5, y: 4.78, w: 9, h: 0.28,
    fontSize: 11, bold: true, color: C.blue,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  s.addNotes('Slide 8 — Alinhamento pedagógico com a Taxonomia Ânima e os 5 princípios de neuroaprendizagem.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 9 — TECNOLOGIA E ESCALA
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '09 — TECNOLOGIA E ESCALA');

  s.addText('Arquitetura universal — pronta para escalar', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 20, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  // Tech chips
  const techs = ['🤖 LLM Adaptativo', '📊 Learning Analytics\nTempo Real', '🔒 LGPD Nativo', '🔗 API Aberta'];
  const tw = 2.1, tg = 0.16;
  const totalTW = techs.length * tw + (techs.length - 1) * tg;
  const tx0 = (W - totalTW) / 2;

  techs.forEach((t, i) => {
    const tx = tx0 + i * (tw + tg);
    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: 1.25, w: tw, h: 0.55,
      fill: { color: '0C1A3A' }, line: { color: '1E3870', pt: 1 }, rectRadius: 0.06
    });
    s.addText(t, {
      x: tx, y: 1.25, w: tw, h: 0.55,
      fontSize: 10, bold: true, color: C.blue,
      fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 4
    });
  });

  // Data flow
  const ds = ['LMS', 'ERP Acadêmico', 'AVA', 'vigIA ✦'];
  const dsHL = [false, false, false, true];
  const dsw = 1.62, dsarr = 0.3, dsg = 0.1;
  const totalDSW = ds.length * dsw + (ds.length - 1) * (dsarr + dsg);
  const dsx0 = (W - totalDSW) / 2;

  ds.forEach((d, i) => {
    const dx = dsx0 + i * (dsw + dsarr + dsg);
    const hl = dsHL[i];
    s.addShape(pres.ShapeType.roundRect, {
      x: dx, y: 1.98, w: dsw, h: 0.4,
      fill: { color: hl ? '0C1A3A' : C.card },
      line: { color: hl ? C.blue : C.border, pt: 1 }, rectRadius: 0.05
    });
    s.addText(d, {
      x: dx, y: 1.98, w: dsw, h: 0.4,
      fontSize: 11, bold: hl, color: hl ? C.blue : C.muted,
      fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
    });
    if (i < ds.length - 1) {
      s.addText('→', {
        x: dx + dsw, y: 2.0, w: dsarr + dsg, h: 0.36,
        fontSize: 15, color: C.blue,
        fontFace: 'Arial', align: 'center', valign: 'middle', isTextBox: true, margin: 0
      });
    }
  });

  // No-bias note
  s.addText('O aluno não preenche absolutamente nada — eliminando viés de autorrelato,\nmonitorando inclusive quem está em crise mais profunda.', {
    x: 1, y: 2.55, w: 8, h: 0.55,
    fontSize: 11, color: C.muted,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  // Scale box
  card(s, 1, 3.2, 8, 0.5, '0D2B1A', '1A5C32', 0.07);
  s.addText('🌐  Ânima = piloto  ·  Expansão para qualquer instituição com LMS', {
    x: 1, y: 3.2, w: 8, h: 0.5,
    fontSize: 13, bold: true, color: C.success,
    fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });

  // Pilot info card
  card(s, 1, 3.85, 8, 1.15, '0C1A3A', C.blue, 0.1);
  s.addText([
    { text: '🎓  Piloto: ', options: { bold: true, color: C.white } },
    { text: 'São Judas Polo Santo Amaro + FADERGS  ', options: { color: C.muted } },
    { text: '→  Escala para todo ecossistema Ânima', options: { bold: true, color: C.blue } }
  ], {
    x: 1, y: 3.9, w: 8, h: 0.46,
    fontSize: 11.5, fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });
  s.addText('Stack: Python · Node.js · React · PostgreSQL · API REST · LGPD Compliance', {
    x: 1, y: 4.4, w: 8, h: 0.52,
    fontSize: 10, color: C.muted,
    fontFace: 'Calibri', align: 'center', valign: 'middle', isTextBox: true, margin: 0
  });

  s.addNotes('Slide 9 — Tecnologia universal com LGPD nativo. Piloto Ânima com escala para todo o ecossistema.');
}

// ═══════════════════════════════════════════════════════
// SLIDE 10 — IMPACTO ESPERADO
// ═══════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  addBg(s);
  badge(s, '10 — IMPACTO ESPERADO');

  s.addText('Três números que mudam a equação', {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 20, bold: true, color: C.white,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  const metrics = [
    { n: '−30%', l: 'Taxa de evasão\nnos primeiros semestres', fill: '2A0A0E', line: 'AA1122', col: C.danger  },
    { n: '+40%', l: 'Engajamento dos\nalunos em risco',         fill: '0D2B1A', line: '1A5C32', col: C.success },
    { n: '6×',   l: 'Detecção de risco\navante dos sistemas',   fill: '0B1530', line: '1E3870', col: C.blue    },
  ];

  const mw = 2.82, mg = 0.32, mh = 2.2;
  const totalMW = metrics.length * mw + (metrics.length - 1) * mg;
  const mx0 = (W - totalMW) / 2;

  metrics.forEach((m, i) => {
    const mx = mx0 + i * (mw + mg);
    card(s, mx, 1.22, mw, mh, m.fill, m.line, 0.12);
    s.addText(m.n, {
      x: mx, y: 1.38, w: mw, h: 1.12,
      fontSize: 56, bold: true, color: m.col,
      fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
    });
    s.addText(m.l, {
      x: mx + 0.1, y: 2.55, w: mw - 0.2, h: 0.7,
      fontSize: 11, bold: true, color: m.col,
      fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
    });
  });

  // Quote
  s.addText('"Cada aluno que chega ao fim do curso é um futuro que a sociedade não perdeu."', {
    x: 0.8, y: 3.62, w: 8.4, h: 0.55,
    fontSize: 13, italic: true, bold: true, color: C.violet,
    fontFace: 'Calibri', align: 'center', isTextBox: true, margin: 0
  });

  // Thank you line
  s.addText('Equipe vigIA. Obrigado.', {
    x: 0.5, y: 4.3, w: 9, h: 0.4,
    fontSize: 18, bold: true, color: C.blue,
    fontFace: 'Arial', align: 'center', isTextBox: true, margin: 0
  });

  // vigIA mini-logo
  s.addText([
    { text: 'vig', options: { color: C.white } },
    { text: 'IA',  options: { color: C.blue  } }
  ], {
    x: 0.5, y: 4.78, w: 9, h: 0.35,
    fontSize: 20, bold: true, fontFace: 'Arial', align: 'center',
    isTextBox: true, margin: 0
  });

  s.addNotes('Slide 10 — Impacto esperado: -30% evasão, +40% engajamento, 6x detecção antecipada. Encerramento.');
}

// ── Write file ──────────────────────────────────────────
const OUT = '/tmp/claude-0/-home-claude/0f3da570-6e0f-5003-a59a-a6b7c1e79178/scratchpad/vigIA_Apresentacao.pptx';

pres.writeFile({ fileName: OUT })
  .then(() => {
    const fs = require('fs');
    const sz = (fs.statSync(OUT).size / 1024 / 1024).toFixed(1);
    console.log(`\n✅ PPTX gerado: ${OUT}`);
    console.log(`   Tamanho: ${sz} MB`);
    console.log(`   Slides: 10`);
  })
  .catch(e => {
    console.error('ERRO:', e.message);
    process.exit(1);
  });
