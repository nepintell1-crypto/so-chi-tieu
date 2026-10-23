/* Sổ chi tiêu cá nhân
   Bước 2: nhập / sửa / xoá khoản thu chi, xem theo tháng, tự tính tổng.
   Bước 3: ngân sách tháng + cảnh báo, báo cáo năm có biểu đồ, xuất file Excel.
   Bước 4: sao lưu / khôi phục dữ liệu, chạy offline khi cài lên điện thoại.
   Danh mục theo file Excel cũ: + thu nhập, A ăn chơi, C thiết yếu cá nhân, G gia đình, K chi khác; thêm T công ty. */
(function () {
  'use strict';

  var CATS = {
    '+': { name: 'Thu nhập', short: 'Thu nhập', color: 'var(--cat-plus)' },
    A: { name: 'Ăn chơi', short: 'Ăn chơi', color: 'var(--cat-A)' },
    C: { name: 'Thiết yếu cá nhân', short: 'Cá nhân', color: 'var(--cat-C)' },
    G: { name: 'Gia đình', short: 'Gia đình', color: 'var(--cat-G)' },
    K: { name: 'Chi khác', short: 'Khác', color: 'var(--cat-K)' },
    T: { name: 'Công ty', short: 'Công ty', color: 'var(--cat-T)' }
  };
  var SPEND_CATS = ['A', 'C', 'G', 'K', 'T'];
  var WEEKDAYS = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
  var MAX_AMOUNT = 999999999999;
  var WARN_AT = 0.8; // cảnh báo khi chi đạt 80% hạn mức
  // Bản cài trên điện thoại dùng file kèm theo (chạy offline); không có thì tải từ CDN
  var XLSX_URLS = ['vendor/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];

  // ---------- Lưu trữ ----------
  // Dữ liệu lưu ngay trên máy (localStorage). Nếu trình duyệt chặn bộ nhớ, app vẫn chạy nhưng chỉ giữ trong phiên.
  var Store = {
    key: 'sochitieu.v1',
    memory: null,
    persistent: true,
    load: function () {
      try {
        var raw = localStorage.getItem(this.key);
        if (raw) return JSON.parse(raw);
      } catch (e) { this.persistent = false; }
      return this.memory;
    },
    save: function (data) {
      this.memory = data;
      try { localStorage.setItem(this.key, JSON.stringify(data)); this.persistent = true; }
      catch (e) { this.persistent = false; }
    }
  };

  // 5 khoản tháng 5 lấy từ file Excel "quanlychitieuhangngay2025" của anh Sơn.
  function seedData() {
    var d = '2026-05-29', note = 'Nhập từ file Excel';
    return [
      { type: 'chi', cat: 'C', content: 'Cafe', amount: 18000 },
      { type: 'thu', cat: '+', content: 'Lương tháng 5', amount: 9334000 },
      { type: 'chi', cat: 'K', content: 'Hụi', amount: 1000000 },
      { type: 'chi', cat: 'K', content: 'Trả tiền Thiện', amount: 1000000 },
      { type: 'chi', cat: 'K', content: 'Trả tiền A.Tuấn', amount: 1000000 }
    ].map(function (t, i) {
      return { id: 'seed' + i, date: d, type: t.type, cat: t.cat, content: t.content, amount: t.amount, note: note, createdAt: i };
    });
  }

  var db = Store.load();
  if (!db || !Array.isArray(db.items)) {
    db = { version: 1, items: seedData() };
    Store.save(db);
  }
  if (!db.budgets || typeof db.budgets !== 'object') db.budgets = {}; // { 'YYYY-MM': { A: 500000, ... } }

  // ---------- Tiện ích ----------
  function money(n) { return Math.round(n).toLocaleString('vi-VN').replace(/,/g, '.'); }
  function signed(t) { return (t.type === 'thu' ? '+' : '−') + money(t.amount); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function fold(s) {
    return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd');
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function todayISO() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function monthOf(iso) { return iso.slice(0, 7); }
  function shiftMonth(ym, delta) {
    var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + delta;
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    return y + '-' + pad(m + 1);
  }
  function monthLabel(ym) { return 'Tháng ' + (+ym.slice(5, 7)) + '/' + ym.slice(0, 4); }
  function dayLabel(iso) {
    var p = iso.split('-'), d = new Date(+p[0], +p[1] - 1, +p[2]);
    return WEEKDAYS[d.getDay()] + ', ' + p[2] + '/' + p[1];
  }
  function validDate(iso) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
    var p = iso.split('-'), d = new Date(+p[0], +p[1] - 1, +p[2]);
    return d.getFullYear() === +p[0] && d.getMonth() === +p[1] - 1 && d.getDate() === +p[2];
  }
  function digits(s) { return String(s).replace(/\D/g, ''); }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }

  function sortDesc(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return (b.createdAt || 0) - (a.createdAt || 0);
  }
  function itemsOfMonth(ym) { return db.items.filter(function (t) { return monthOf(t.date) === ym; }).sort(sortDesc); }
  function itemsOfYear(y) { return db.items.filter(function (t) { return t.date.slice(0, 4) === y; }); }
  function budgetOf(ym) { return db.budgets[ym] || {}; }
  // level: none (chưa đặt) | ok | warn (>= 80%) | over (vượt)
  function budgetStatus(spent, limit) {
    if (!limit) return { level: 'none', pct: 0 };
    var r = spent / limit;
    return { level: r > 1 ? 'over' : r >= WARN_AT ? 'warn' : 'ok', pct: Math.round(r * 100) };
  }
  function budgetLine(k, spent, limit) {
    var st = budgetStatus(spent, limit);
    if (st.level === 'over') return CATS[k].name + ' đã vượt hạn mức ' + money(spent - limit) + ' đ';
    if (st.level === 'warn') return CATS[k].name + ' đã dùng ' + st.pct + '% hạn mức, còn ' + money(limit - spent) + ' đ';
    return '';
  }
  function totals(list) {
    var r = { thu: 0, chi: 0 };
    SPEND_CATS.forEach(function (k) { r[k] = 0; });
    list.forEach(function (t) {
      if (t.type === 'thu') r.thu += t.amount;
      else { r.chi += t.amount; r[t.cat] += t.amount; }
    });
    r.con = r.thu - r.chi;
    return r;
  }

  // ---------- Trạng thái màn hình ----------
  function latestMonth() {
    if (!db.items.length) return monthOf(todayISO());
    return monthOf(db.items.slice().sort(sortDesc)[0].date);
  }
  var ui = {
    view: 'home', month: latestMonth(), filter: 'all', query: '',
    draft: null, errors: {}, confirmDelete: false, back: 'home', toast: '',
    year: null, exporting: false, exportMsg: '', restore: null, dataMsg: ''
  };
  ui.year = ui.month.slice(0, 4);

  function newDraft() {
    var today = todayISO();
    // Nếu đang xem tháng khác tháng hiện tại, gợi ý ngày đầu của tháng đang xem
    var date = monthOf(today) === ui.month ? today : ui.month + '-01';
    return { id: null, type: 'chi', cat: '', amount: '', date: date, content: '', note: '' };
  }

  // ---------- Icon ----------
  var I = {
    prev: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M15 5l-7 7 7 7"/></svg>',
    next: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M9 5l7 7-7 7"/></svg>',
    close: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    home: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z"/></svg>',
    list: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M8 6h13M8 12h13M8 18h13M3 6h1M3 12h1M3 18h1"/></svg>',
    plus: '<svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>',
    budget: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>',
    chart: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg>',
    warn: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M12 3l10 18H2z"/><path d="M12 10v4M12 17v.5"/></svg>',
    download: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M12 4v11M7 10l5 5 5-5M5 20h14"/></svg>',
    search: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg>'
  };

  // ---------- Các phần giao diện ----------
  function monthHeader(title) {
    return '<header class="top">' +
      '<button class="icon-btn" data-action="month" data-delta="-1" aria-label="Tháng trước">' + I.prev + '</button>' +
      '<h1>' + esc(title) + '</h1>' +
      '<button class="icon-btn" data-action="month" data-delta="1" aria-label="Tháng sau">' + I.next + '</button>' +
      '</header>';
  }

  function txRow(t, showDate) {
    var c = CATS[t.cat];
    var sub = showDate ? t.date.slice(8, 10) + '/' + t.date.slice(5, 7) + (t.note ? ' · ' + t.note : '') : (t.note || c.name);
    return '<button class="tx" data-action="edit" data-id="' + esc(t.id) + '" aria-label="Sửa ' + esc(t.content) + '">' +
      '<span class="tag" style="background:' + c.color + '">' + esc(t.cat) + '</span>' +
      '<span class="mid"><span>' + esc(t.content) + '</span><span class="sub">' + esc(sub) + '</span></span>' +
      '<span class="amt num ' + (t.type === 'thu' ? 'in' : 'out') + '">' + signed(t) + '</span>' +
      '</button>';
  }

  function emptyMonth() {
    return '<div class="empty"><p>Chưa có khoản thu chi nào trong ' + esc(monthLabel(ui.month).toLowerCase()) + '.</p>' +
      '<button class="btn" data-action="add">Thêm khoản đầu tiên</button></div>';
  }

  function viewHome() {
    var list = itemsOfMonth(ui.month), s = totals(list), bud = budgetOf(ui.month);
    var maxCat = Math.max.apply(null, [1].concat(SPEND_CATS.map(function (k) { return s[k]; })));
    var warnings = SPEND_CATS.map(function (k) { return budgetLine(k, s[k], bud[k]); }).filter(Boolean);
    var cats = SPEND_CATS.map(function (k) {
      var pct = s.chi ? Math.round(s[k] / s.chi * 1000) / 10 : 0;
      var w = s[k] ? Math.max(2, s[k] / maxCat * 100) : 0;
      var lim = bud[k] ? ' / ' + money(bud[k]) : '';
      return '<div class="cat-row"><div class="t"><span>' + k + ' · ' + CATS[k].name + '</span>' +
        '<span class="num">' + money(s[k]) + lim + ' đ <span style="color:var(--muted)">(' + String(pct).replace('.', ',') + '%)</span></span></div>' +
        '<div class="bar" role="img" aria-label="' + CATS[k].name + ' ' + String(pct).replace('.', ',') + '% tổng chi"><span style="width:' + w + '%;background:' + CATS[k].color + '"></span></div></div>';
    }).join('');
    var recent = list.slice(0, 5).map(function (t) { return txRow(t, true); }).join('');

    return monthHeader(monthLabel(ui.month)) +
      '<section class="hero" aria-label="Tóm tắt tháng">' +
        '<div><div class="lbl">Còn lại</div><div class="big num">' + money(s.con) + ' đ</div></div>' +
        '<div class="pair"><div><div class="lbl">Tổng thu (+)</div><b class="num">' + money(s.thu) + ' đ</b></div>' +
        '<div><div class="lbl">Tổng chi (−)</div><b class="num">' + money(s.chi) + ' đ</b></div></div>' +
      '</section>' +
      (warnings.length ? '<section class="alert" role="status" aria-label="Cảnh báo ngân sách">' + warnings.map(function (w) {
        return '<div>' + I.warn + '<span>' + esc(w) + '</span></div>'; }).join('') +
        '<button class="link-btn" data-action="go" data-view="budget">Xem ngân sách</button></section>' : '') +
      '<section class="card"><h2>Chi theo danh mục</h2>' + cats + '</section>' +
      '<section class="card"><h2>Gần đây' + (list.length ? '<button class="link-btn" data-action="go" data-view="list">Xem tất cả (' + list.length + ')</button>' : '') + '</h2>' +
        (list.length ? '<div>' + recent + '</div>' : emptyMonth()) +
      '</section>' +
      (Store.persistent ? '' : '<p class="note">Trình duyệt đang chặn bộ nhớ, dữ liệu sẽ mất khi đóng trang. Bản cài lên điện thoại ở Bước 4 sẽ khắc phục việc này.</p>');
  }

  function listBody() {
    var list = itemsOfMonth(ui.month);
    var q = fold(ui.query.trim());
    var shown = list.filter(function (t) {
      if (ui.filter === 'thu' && t.type !== 'thu') return false;
      if (SPEND_CATS.indexOf(ui.filter) >= 0 && t.cat !== ui.filter) return false;
      if (q && fold(t.content + ' ' + (t.note || '')).indexOf(q) < 0) return false;
      return true;
    });
    if (!list.length) return emptyMonth();
    if (!shown.length) return '<div class="empty"><p>Không có khoản nào khớp bộ lọc.</p><button class="btn ghost" data-action="clear-filter">Bỏ lọc</button></div>';

    var groups = '', curDay = null, buf = [], dayNet = 0;
    function flush() {
      if (!curDay) return;
      groups += '<div class="day-head"><b>' + esc(dayLabel(curDay)) + '</b><span class="num">' + (dayNet >= 0 ? '+' : '−') + money(Math.abs(dayNet)) + '</span></div>' + buf.join('');
    }
    shown.forEach(function (t) {
      if (t.date !== curDay) { flush(); curDay = t.date; buf = []; dayNet = 0; }
      buf.push(txRow(t, false));
      dayNet += t.type === 'thu' ? t.amount : -t.amount;
    });
    flush();
    return groups;
  }

  function viewList() {
    var list = itemsOfMonth(ui.month), s = totals(list);
    var chips = [['all', 'Tất cả'], ['thu', 'Thu'], ['A', 'A · Ăn chơi'], ['C', 'C · Cá nhân'], ['G', 'G · Gia đình'], ['K', 'K · Khác'], ['T', 'T · Công ty']]
      .map(function (c) {
        return '<button class="chip" data-action="filter" data-f="' + c[0] + '" aria-pressed="' + (ui.filter === c[0]) + '">' + c[1] + '</button>';
      }).join('');

    return monthHeader('Sổ thu chi · ' + monthLabel(ui.month)) +
      '<label class="search">' + I.search +
        '<input id="q" type="search" placeholder="Tìm theo nội dung hoặc ghi chú" aria-label="Tìm theo nội dung hoặc ghi chú" value="' + esc(ui.query) + '" data-input="query"></label>' +
      '<div class="chips" role="group" aria-label="Lọc theo danh mục">' + chips + '</div>' +
      '<div class="mini"><div>Thu<b class="num" style="color:var(--in)">' + money(s.thu) + '</b></div>' +
        '<div>Chi<b class="num" style="color:var(--out)">' + money(s.chi) + '</b></div>' +
        '<div>Còn lại<b class="num">' + money(s.con) + '</b></div></div>' +
      '<section class="card" style="gap:0" id="list-body" aria-label="Danh sách khoản thu chi">' + listBody() + '</section>' +
      (list.length ? '<p style="margin:0;text-align:center;font-size:12px;color:var(--muted)">Chạm vào một dòng để sửa hoặc xoá.</p>' : '');
  }

  function viewForm() {
    var d = ui.draft, e = ui.errors, isEdit = !!d.id;
    var catBtns = d.type === 'thu'
      ? '<p style="margin:0;font-size:13px;color:var(--muted)">Khoản thu được ghi mã <b style="color:var(--in)">+</b> (thu nhập).</p>'
      : '<div class="cats" role="group" aria-label="Danh mục">' + SPEND_CATS.map(function (k) {
          return '<button type="button" class="cat" data-action="cat" data-cat="' + k + '" aria-pressed="' + (d.cat === k) + '">' +
            '<b style="color:' + CATS[k].color + '">' + k + '</b>' + CATS[k].short + '</button>';
        }).join('') + '</div>';
    var amountShown = d.amount ? money(+d.amount) : '';

    var confirm = ui.confirmDelete
      ? '<div class="confirm" role="alert"><span>Xoá khoản "' + esc(d.content) + '"? Không thể hoàn tác.</span>' +
        '<div class="row"><button type="button" class="btn ghost" data-action="cancel-delete">Giữ lại</button>' +
        '<button type="button" class="btn danger solid" data-action="do-delete">Xoá</button></div></div>'
      : '';

    return '<header class="top">' +
        '<button class="icon-btn" data-action="close-form" aria-label="Đóng">' + I.close + '</button>' +
        '<h1>' + (isEdit ? 'Sửa khoản' : 'Thêm khoản mới') + '</h1><span style="width:44px"></span></header>' +
      '<form class="form" id="txform" novalidate>' +
        '<div class="seg" role="group" aria-label="Loại">' +
          '<button type="button" data-action="type" data-type="chi" aria-pressed="' + (d.type === 'chi') + '">Chi (−)</button>' +
          '<button type="button" data-action="type" data-type="thu" aria-pressed="' + (d.type === 'thu') + '">Thu (+)</button></div>' +
        '<label class="field" for="f-amount">Số tiền (đồng)' +
          '<input id="f-amount" class="amount num" inputmode="numeric" autocomplete="off" placeholder="0" value="' + amountShown + '" data-input="amount"' + (e.amount ? ' aria-invalid="true" aria-describedby="e-amount"' : '') + '>' +
          (e.amount ? '<span class="err" id="e-amount">' + e.amount + '</span>' : '') + '</label>' +
        '<div class="field"><span>Danh mục (KH)</span>' + catBtns + (e.cat ? '<span class="err">' + e.cat + '</span>' : '') + '</div>' +
        '<label class="field" for="f-date">Ngày' +
          '<input id="f-date" type="date" value="' + esc(d.date) + '" data-input="date"' + (e.date ? ' aria-invalid="true"' : '') + '>' +
          (e.date ? '<span class="err">' + e.date + '</span>' : '') + '</label>' +
        '<label class="field" for="f-content">Nội dung thu chi' +
          '<input id="f-content" maxlength="120" autocomplete="off" placeholder="' + (d.type === 'thu' ? 'Ví dụ: Lương tháng 6' : 'Ví dụ: Ăn sáng, đổ xăng') + '" value="' + esc(d.content) + '" data-input="content"' + (e.content ? ' aria-invalid="true"' : '') + '>' +
          (e.content ? '<span class="err">' + e.content + '</span>' : '') + '</label>' +
        '<label class="field" for="f-note">Ghi chú (không bắt buộc)' +
          '<input id="f-note" maxlength="200" autocomplete="off" value="' + esc(d.note) + '" data-input="note"></label>' +
        '<div class="form-actions">' +
          '<button type="submit" class="btn">' + (isEdit ? 'Lưu thay đổi' : (d.type === 'thu' ? 'Lưu khoản thu' : 'Lưu khoản chi')) + '</button>' +
          (isEdit && !ui.confirmDelete ? '<button type="button" class="btn danger" data-action="ask-delete">Xoá khoản này</button>' : '') +
          confirm +
        '</div>' +
      '</form>';
  }

  function viewBudget() {
    var list = itemsOfMonth(ui.month), s = totals(list), bud = budgetOf(ui.month);
    var prev = budgetOf(shiftMonth(ui.month, -1));
    var hasPrev = SPEND_CATS.some(function (k) { return prev[k]; });
    var limitSum = 0, spentInBudget = 0;
    SPEND_CATS.forEach(function (k) { if (bud[k]) { limitSum += bud[k]; spentInBudget += s[k]; } });

    var cards = SPEND_CATS.map(function (k) {
      var st = budgetStatus(s[k], bud[k]);
      var color = st.level === 'over' ? 'var(--out)' : st.level === 'warn' ? 'var(--warn-bar)' : CATS[k].color;
      var text;
      if (st.level === 'none') text = 'Đã chi ' + money(s[k]) + ' đ · chưa đặt hạn mức';
      else if (st.level === 'over') text = 'Đã chi ' + money(s[k]) + ' đ · vượt ' + money(s[k] - bud[k]) + ' đ';
      else text = 'Đã chi ' + money(s[k]) + ' đ (' + st.pct + '%) · còn ' + money(bud[k] - s[k]) + ' đ';
      return '<div class="card budget">' +
        '<div class="bhead"><span class="tag" style="background:' + CATS[k].color + '">' + k + '</span>' +
          '<label for="b-' + k + '" class="bname">' + CATS[k].name + '</label>' +
          '<input id="b-' + k + '" class="bin num" inputmode="numeric" autocomplete="off" placeholder="Nhập hạn mức" data-budget="' + k + '" value="' + (bud[k] ? money(bud[k]) : '') + '"></div>' +
        '<div class="bar" role="img" aria-label="' + esc(text) + '"><span style="width:' + Math.min(100, st.pct) + '%;background:' + color + '"></span></div>' +
        '<div class="bstat' + (st.level === 'over' ? ' over' : st.level === 'warn' ? ' warn' : '') + '">' + text + '</div>' +
      '</div>';
    }).join('');

    return monthHeader('Ngân sách · ' + monthLabel(ui.month)) +
      '<p class="lead">Đặt hạn mức chi cho từng danh mục. App báo khi chi đạt 80% và khi vượt hạn mức. Để trống nếu không cần theo dõi.</p>' +
      (limitSum ? '<div class="mini" style="grid-template-columns:repeat(2,minmax(0,1fr))"><div>Tổng hạn mức<b class="num">' + money(limitSum) + ' đ</b></div>' +
        '<div>Đã chi (các mục có hạn mức)<b class="num" style="color:' + (spentInBudget > limitSum ? 'var(--out)' : 'var(--fg)') + '">' + money(spentInBudget) + ' đ</b></div></div>' : '') +
      cards +
      (hasPrev ? '<button class="btn ghost" data-action="copy-budget">Chép hạn mức từ ' + monthLabel(shiftMonth(ui.month, -1)).toLowerCase() + '</button>' : '');
  }

  // Bước nhảy trục đẹp: 1, 2, 2.5, 5 x 10^n
  function niceStep(raw) {
    var p = Math.pow(10, Math.floor(Math.log10(raw))), f = raw / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
  }
  function shortMoney(v) {
    if (v === 0) return '0';
    if (v >= 1e6) return String(Math.round(v / 1e5) / 10).replace('.', ',') + ' tr';
    return String(Math.round(v / 100) / 10).replace('.', ',') + ' k';
  }

  function yearChart(months) {
    var max = 0;
    months.forEach(function (m) { max = Math.max(max, m.thu, m.chi); });
    var W = 340, H = 210, L = 44, R = 6, T = 10, B = 26;
    var step = niceStep(max / 4), top = Math.ceil(max / step) * step;
    var y = function (v) { return T + (H - T - B) * (1 - v / top); };
    var slot = (W - L - R) / 12, bw = Math.min(10, slot / 2 - 2);
    var g = '';
    for (var v = 0; v <= top + 1; v += step) {
      g += '<line x1="' + L + '" x2="' + (W - R) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="var(--line)" stroke-width="1"/>' +
        '<text x="' + (L - 6) + '" y="' + (y(v) + 4) + '" text-anchor="end" font-size="10" fill="var(--muted)">' + shortMoney(v) + '</text>';
    }
    months.forEach(function (m, i) {
      var cx = L + slot * i + slot / 2;
      if (m.thu) g += '<rect x="' + (cx - bw - 1) + '" y="' + y(m.thu) + '" width="' + bw + '" height="' + (y(0) - y(m.thu)) + '" rx="2" fill="var(--in)"><title>Tháng ' + (i + 1) + ': thu ' + money(m.thu) + ' đ</title></rect>';
      if (m.chi) g += '<rect x="' + (cx + 1) + '" y="' + y(m.chi) + '" width="' + bw + '" height="' + (y(0) - y(m.chi)) + '" rx="2" fill="var(--bar-out)"><title>Tháng ' + (i + 1) + ': chi ' + money(m.chi) + ' đ</title></rect>';
      g += '<text x="' + cx + '" y="' + (H - 8) + '" text-anchor="middle" font-size="10" fill="' + (m.thu || m.chi ? 'var(--fg)' : 'var(--muted)') + '">' + (i + 1) + '</text>';
    });
    var desc = months.map(function (m, i) { return m.thu || m.chi ? 'tháng ' + (i + 1) + ' thu ' + money(m.thu) + ', chi ' + money(m.chi) : null; }).filter(Boolean).join('; ');
    return '<svg viewBox="0 0 ' + W + ' ' + H + '" width="100%" role="img" aria-label="Thu chi theo tháng: ' + esc(desc) + '">' + g + '</svg>';
  }

  function dataSection() {
    var r = ui.restore;
    var confirm = r
      ? '<div class="confirm" role="alert"><span>Thay toàn bộ dữ liệu hiện tại (' + db.items.length + ' khoản) bằng ' + r.items.length +
        ' khoản trong file sao lưu ngày ' + esc(r.savedAt && validDate(r.savedAt) ? r.savedAt.split('-').reverse().join('/') : '?') + '? Không thể hoàn tác.</span>' +
        '<div class="row"><button type="button" class="btn ghost" data-action="cancel-restore">Huỷ</button>' +
        '<button type="button" class="btn danger solid" data-action="do-restore">Khôi phục</button></div></div>'
      : '';
    return '<section class="card"><h2>Dữ liệu của anh</h2>' +
      '<p class="lead">Dữ liệu chỉ lưu trên máy này. Nên sao lưu mỗi tháng một lần và cất file ở Google Drive hoặc Zalo "Cloud của tôi".</p>' +
      '<button class="btn ghost with-icon" data-action="backup">' + I.download + 'Sao lưu dữ liệu (.json)</button>' +
      '<label class="btn ghost with-icon file-btn" for="restore-file">Khôi phục từ file sao lưu' +
        '<input id="restore-file" type="file" accept=".json,application/json" data-restore="1"></label>' +
      confirm +
      (ui.dataMsg ? '<p class="lead" role="status">' + esc(ui.dataMsg) + '</p>' : '') +
      '</section>';
  }

  function viewReport() {
    var yr = ui.year, items = itemsOfYear(yr), ys = totals(items);
    var months = [];
    for (var i = 1; i <= 12; i++) months.push(totals(items.filter(function (t) { return +t.date.slice(5, 7) === i; })));

    var rows = months.map(function (m, i) {
      var empty = !m.thu && !m.chi;
      return '<tr' + (empty ? ' class="dim"' : '') + '><th scope="row"><button class="link-btn" data-action="open-month" data-m="' + yr + '-' + pad(i + 1) + '">Tháng ' + (i + 1) + '</button></th>' +
        '<td>' + (empty ? '–' : money(m.thu)) + '</td><td>' + (empty ? '–' : money(m.chi)) + '</td><td>' + (empty ? '–' : money(m.con)) + '</td></tr>';
    }).join('');

    var cats = SPEND_CATS.map(function (k) {
      var pct = ys.chi ? Math.round(ys[k] / ys.chi * 1000) / 10 : 0;
      return '<div class="cat-row"><div class="t"><span>' + k + ' · ' + CATS[k].name + '</span><span class="num">' + money(ys[k]) + ' đ <span style="color:var(--muted)">(' + String(pct).replace('.', ',') + '%)</span></span></div>' +
        '<div class="bar"><span style="width:' + pct + '%;background:' + CATS[k].color + '"></span></div></div>';
    }).join('');

    var head = '<header class="top">' +
      '<button class="icon-btn" data-action="year" data-delta="-1" aria-label="Năm trước">' + I.prev + '</button>' +
      '<h1>Báo cáo năm ' + yr + '</h1>' +
      '<button class="icon-btn" data-action="year" data-delta="1" aria-label="Năm sau">' + I.next + '</button></header>';

    if (!items.length) {
      return head + '<section class="card"><div class="empty"><p>Năm ' + yr + ' chưa có khoản thu chi nào.</p></div></section>' + dataSection();
    }
    return head +
      '<section class="hero" aria-label="Tổng cả năm">' +
        '<div><div class="lbl">Tiền còn lại cả năm</div><div class="big num">' + money(ys.con) + ' đ</div></div>' +
        '<div class="pair"><div><div class="lbl">Tổng thu</div><b class="num">' + money(ys.thu) + ' đ</b></div>' +
        '<div><div class="lbl">Tổng chi</div><b class="num">' + money(ys.chi) + ' đ</b></div></div></section>' +
      '<section class="card"><h2>So sánh các tháng<span class="legend"><span><i style="background:var(--in)"></i>Thu</span><span><i style="background:var(--bar-out)"></i>Chi</span></span></h2>' + yearChart(months) + '</section>' +
      '<section class="card"><h2>Chi tiết từng tháng</h2><div class="tbl-wrap"><table class="tbl num"><thead><tr><th scope="col">Tháng</th><th scope="col">Thu</th><th scope="col">Chi</th><th scope="col">Còn lại</th></tr></thead><tbody>' + rows +
        '</tbody><tfoot><tr><th scope="row">Cả năm</th><td>' + money(ys.thu) + '</td><td>' + money(ys.chi) + '</td><td>' + money(ys.con) + '</td></tr></tfoot></table></div></section>' +
      '<section class="card"><h2>Chi theo danh mục cả năm</h2>' + cats + '</section>' +
      '<button class="btn ghost with-icon" data-action="export"' + (ui.exporting ? ' disabled' : '') + '>' + I.download + (ui.exporting ? 'Đang tạo file...' : 'Xuất file Excel năm ' + yr) + '</button>' +
      (ui.exportMsg ? '<p class="lead" role="status">' + esc(ui.exportMsg) + '</p>' : '') +
      dataSection();
  }

  // ---------- Xuất Excel ----------
  var xlsxPromise = null;
  function loadXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) {
      xlsxPromise = new Promise(function (ok, fail) {
        (function attempt(i) {
          if (i >= XLSX_URLS.length) { xlsxPromise = null; fail(new Error('load failed')); return; }
          var sc = document.createElement('script');
          sc.src = XLSX_URLS[i];
          sc.onload = function () { window.XLSX ? ok(window.XLSX) : attempt(i + 1); };
          sc.onerror = function () { sc.remove(); attempt(i + 1); };
          document.head.appendChild(sc);
        })(0);
      });
    }
    return xlsxPromise;
  }
  function excelDate(iso) {
    var p = iso.split('-');
    return (Date.UTC(+p[0], +p[1] - 1, +p[2]) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  // Bố cục giống file Excel cũ: sheet "Bao cao" + 12 sheet T1..T12
  function buildWorkbook(X, yr) {
    var items = itemsOfYear(yr), wb = X.utils.book_new();
    var money0 = '#,##0';
    var catCols = SPEND_CATS.map(function (k) { return k + ' - ' + CATS[k].name; });
    var rep = [['BÁO CÁO THU CHI NĂM ' + yr], [], ['Tháng', 'Tổng thu', 'Tổng chi', 'Còn lại'].concat(catCols)];
    function rowOf(label, m) { return [label, m.thu, m.chi, m.con].concat(SPEND_CATS.map(function (k) { return m[k]; })); }
    for (var i = 1; i <= 12; i++) {
      rep.push(rowOf('Tháng ' + i, totals(items.filter(function (t) { return +t.date.slice(5, 7) === i; }))));
    }
    rep.push(rowOf('Cả năm', totals(items)));
    var ws = X.utils.aoa_to_sheet(rep);
    var numCols = 'BCDEFGHIJKLMNOP'.slice(0, 3 + SPEND_CATS.length).split('');
    // Còn lại = Thu - Chi; dòng Cả năm = SUM, để anh sửa số trong Excel vẫn tự tính lại
    for (var r = 4; r <= 15; r++) ws['D' + r] = { t: 'n', v: rep[r - 1][3], f: 'B' + r + '-C' + r };
    numCols.forEach(function (c, j) { ws[c + '16'] = { t: 'n', v: rep[15][j + 1], f: 'SUM(' + c + '4:' + c + '15)' }; });
    for (r = 4; r <= 16; r++) numCols.forEach(function (c) { if (ws[c + r]) ws[c + r].z = money0; });
    ws['!cols'] = [{ wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 14 }].concat(catCols.map(function (h) { return { wch: Math.max(13, h.length + 2) }; }));
    X.utils.book_append_sheet(wb, ws, 'Bao cao');

    for (i = 1; i <= 12; i++) {
      var list = items.filter(function (t) { return +t.date.slice(5, 7) === i; })
        .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.createdAt || 0) - (b.createdAt || 0); });
      var aoa = [['Ngày', 'KH', 'Nội dung thu chi', 'Thu + / Chi -', 'Số tiền', 'Ghi chú']];
      list.forEach(function (t) { aoa.push([0, t.cat, t.content, t.type === 'thu' ? '+' : '-', t.amount, t.note || '']); });
      var n = list.length, last = n + 1;
      aoa.push([], ['', '', 'Tổng thu', '', 0], ['', '', 'Tổng chi', '', 0], ['', '', 'Còn lại', '', 0]);
      var sh = X.utils.aoa_to_sheet(aoa);
      list.forEach(function (t, j) {
        sh['A' + (j + 2)] = { t: 'n', v: excelDate(t.date), z: 'dd/mm/yyyy' };
        sh['E' + (j + 2)].z = money0;
      });
      var m2 = totals(list), rT = last + 2;
      sh['E' + rT] = { t: 'n', v: m2.thu, z: money0, f: n ? 'SUMIF(D2:D' + last + ',"+",E2:E' + last + ')' : undefined };
      sh['E' + (rT + 1)] = { t: 'n', v: m2.chi, z: money0, f: n ? 'SUMIF(D2:D' + last + ',"-",E2:E' + last + ')' : undefined };
      sh['E' + (rT + 2)] = { t: 'n', v: m2.con, z: money0, f: 'E' + rT + '-E' + (rT + 1) };
      sh['!cols'] = [{ wch: 11 }, { wch: 5 }, { wch: 32 }, { wch: 12 }, { wch: 14 }, { wch: 28 }];
      X.utils.book_append_sheet(wb, sh, 'T' + i);
    }
    return X.write(wb, { bookType: 'xlsx', type: 'array' });
  }

  function saveFile(name, buf, mime) {
    var c = window.claude;
    var viaClaude = c && typeof c.use === 'function' ? c.use('downloads') : Promise.resolve(null);
    return viaClaude.then(function (dl) {
      if (dl) return dl.save({ filename: name, data: new Blob([buf]) }).then(function () { return 'saved'; });
      // Bản cài trên điện thoại / trình duyệt thường: tải trực tiếp
      var url = URL.createObjectURL(new Blob([buf], { type: mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      var a = document.createElement('a');
      a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
      return 'saved';
    });
  }

  function exportYear() {
    var yr = ui.year, name = 'SoChiTieu_' + yr + '.xlsx';
    ui.exporting = true; ui.exportMsg = ''; render();
    loadXLSX().then(function (X) {
      return saveFile(name, buildWorkbook(X, yr));
    }).then(function () {
      ui.exportMsg = 'Đã tạo file ' + name + '.';
    }, function (err) {
      var code = err && err.code;
      ui.exportMsg = code === 'declined' ? 'Anh đã huỷ lưu file.'
        : code === 'rate_limited' ? 'Đang có một hộp thoại lưu file mở, thử lại sau giây lát.'
        : err && err.message === 'load failed' ? 'Không tải được bộ tạo file Excel. Kiểm tra kết nối mạng rồi thử lại.'
        : 'Không lưu được file ở chế độ xem này.';
    }).then(function () {
      ui.exporting = false;
      if (ui.view === 'report') render();
    });
  }

  function tabs() {
    var cur = ui.view === 'form' ? ui.back : ui.view;
    function t(v, label, icon) {
      return '<button class="tab" data-action="go" data-view="' + v + '"' + (cur === v ? ' aria-current="page"' : '') + '>' + icon + label + '</button>';
    }
    return '<div class="tabs-in">' + t('home', 'Tổng quan', I.home) + t('list', 'Sổ thu chi', I.list) +
      '<button class="tab add" data-action="add" aria-label="Thêm khoản thu chi">' + I.plus + '</button>' +
      t('budget', 'Ngân sách', I.budget) + t('report', 'Báo cáo', I.chart) + '</div>';
  }

  // ---------- Vẽ lại ----------
  var main = document.getElementById('app'), tabBar = document.getElementById('tabs'), toastEl = document.getElementById('toast');
  var toastTimer = null;

  function render(focusId) {
    var html;
    if (ui.view === 'home') html = viewHome();
    else if (ui.view === 'list') html = viewList();
    else if (ui.view === 'form') html = viewForm();
    else if (ui.view === 'budget') html = viewBudget();
    else html = viewReport();
    main.innerHTML = html;
    tabBar.innerHTML = tabs();
    tabBar.hidden = ui.view === 'form';
    if (focusId) {
      var el = document.getElementById(focusId);
      if (el) { el.focus(); if (el.setSelectionRange && el.type !== 'date' && el.type !== 'search') { var n = el.value.length; el.setSelectionRange(n, n); } }
    }
  }

  function showToast(msg) {
    toastEl.textContent = msg; toastEl.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.hidden = true; }, 2200);
  }

  function go(view) {
    if (view === 'report' && ui.view !== 'report') { ui.year = ui.month.slice(0, 4); ui.exportMsg = ''; ui.dataMsg = ''; ui.restore = null; }
    ui.view = view; ui.confirmDelete = false; ui.errors = {};
    render(); window.scrollTo(0, 0);
  }

  function openForm(item) {
    ui.back = ui.view === 'form' ? ui.back : ui.view;
    ui.draft = item ? { id: item.id, type: item.type, cat: item.cat === '+' ? '' : item.cat, amount: String(item.amount), date: item.date, content: item.content, note: item.note || '' } : newDraft();
    ui.errors = {}; ui.confirmDelete = false; ui.view = 'form';
    render(item ? null : 'f-amount'); window.scrollTo(0, 0);
  }

  function validate(d) {
    var e = {}, amt = +d.amount;
    if (!d.amount || !amt) e.amount = 'Nhập số tiền lớn hơn 0.';
    else if (amt > MAX_AMOUNT) e.amount = 'Số tiền quá lớn, kiểm tra lại số chữ số.';
    if (d.type === 'chi' && SPEND_CATS.indexOf(d.cat) < 0) e.cat = 'Chọn một danh mục A, C, G, K hoặc T.';
    if (!validDate(d.date)) e.date = 'Chọn ngày hợp lệ (ngày/tháng/năm).';
    if (!d.content.trim()) e.content = 'Ghi nội dung để sau này dễ tra lại.';
    return e;
  }

  function saveDraft() {
    var d = ui.draft, e = validate(d);
    ui.errors = e;
    if (Object.keys(e).length) {
      render(e.amount ? 'f-amount' : e.date ? 'f-date' : e.content ? 'f-content' : null);
      return;
    }
    var rec = {
      id: d.id || uid(), date: d.date, type: d.type, cat: d.type === 'thu' ? '+' : d.cat,
      content: d.content.trim(), amount: +d.amount, note: d.note.trim()
    };
    if (d.id) {
      db.items = db.items.map(function (t) { return t.id === d.id ? Object.assign({}, t, rec) : t; });
    } else {
      rec.createdAt = Date.now();
      db.items.push(rec);
    }
    Store.save(db);
    ui.month = monthOf(rec.date);
    var msg = d.id ? 'Đã lưu thay đổi' : 'Đã thêm ' + (rec.type === 'thu' ? 'khoản thu ' : 'khoản chi ') + money(rec.amount) + ' đ';
    if (rec.type === 'chi') {
      var w = budgetLine(rec.cat, totals(itemsOfMonth(ui.month))[rec.cat], budgetOf(ui.month)[rec.cat]);
      if (w) msg += '. Chú ý: ' + w.charAt(0).toLowerCase() + w.slice(1);
    }
    go(ui.back === 'list' ? 'list' : 'home');
    showToast(msg);
  }

  // ---------- Sao lưu / khôi phục ----------
  function backup() {
    var name = 'SoChiTieu_saoluu_' + todayISO() + '.json';
    var payload = JSON.stringify({ app: 'so-chi-tieu', version: 1, savedAt: todayISO(), items: db.items, budgets: db.budgets }, null, 1);
    ui.dataMsg = '';
    saveFile(name, payload, 'application/json').then(function () {
      db.lastBackup = todayISO(); Store.save(db);
      ui.dataMsg = 'Đã tạo file ' + name + ' (' + db.items.length + ' khoản).';
    }, function (err) {
      ui.dataMsg = err && err.code === 'declined' ? 'Anh đã huỷ lưu file.' : 'Không lưu được file ở chế độ xem này.';
    }).then(function () { if (ui.view === 'report') render(); });
  }

  function validItem(t) {
    return t && typeof t.id === 'string' && validDate(t.date) && (t.type === 'thu' || t.type === 'chi') &&
      CATS[t.cat] && (t.type === 'thu') === (t.cat === '+') && typeof t.content === 'string' &&
      typeof t.amount === 'number' && t.amount > 0 && t.amount <= MAX_AMOUNT;
  }

  function readRestoreFile(file) {
    ui.restore = null; ui.dataMsg = '';
    if (!file) return;
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var data = JSON.parse(fr.result);
        if (!data || data.app !== 'so-chi-tieu' || !Array.isArray(data.items)) throw new Error('format');
        var bad = data.items.filter(function (t) { return !validItem(t); }).length;
        if (bad) throw new Error('items');
        ui.restore = { items: data.items, budgets: data.budgets && typeof data.budgets === 'object' ? data.budgets : {}, savedAt: data.savedAt };
      } catch (e) {
        ui.dataMsg = 'File này không phải file sao lưu của Sổ chi tiêu, hoặc đã bị sửa hỏng. Chọn lại file .json do app tạo ra.';
      }
      render();
    };
    fr.onerror = function () { ui.dataMsg = 'Không đọc được file.'; render(); };
    fr.readAsText(file);
  }

  // Xin trình duyệt giữ dữ liệu lâu dài (không tự xoá khi máy thiếu dung lượng)
  try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) {}

  // ---------- Sự kiện ----------
  document.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-action]');
    if (!b) return;
    var a = b.getAttribute('data-action');
    if (a === 'month') { ui.month = shiftMonth(ui.month, +b.getAttribute('data-delta')); render(); }
    else if (a === 'go') go(b.getAttribute('data-view'));
    else if (a === 'year') { ui.year = String(+ui.year + +b.getAttribute('data-delta')); ui.exportMsg = ''; render(); }
    else if (a === 'open-month') { ui.month = b.getAttribute('data-m'); go('home'); }
    else if (a === 'export') { if (!ui.exporting) exportYear(); }
    else if (a === 'backup') backup();
    else if (a === 'cancel-restore') { ui.restore = null; render(); }
    else if (a === 'do-restore') {
      var n = ui.restore.items.length;
      db.items = ui.restore.items; db.budgets = ui.restore.budgets; ui.restore = null;
      Store.save(db); ui.month = latestMonth(); ui.year = ui.month.slice(0, 4);
      ui.dataMsg = 'Đã khôi phục ' + n + ' khoản.'; render();
    }
    else if (a === 'copy-budget') {
      db.budgets[ui.month] = Object.assign({}, budgetOf(shiftMonth(ui.month, -1)));
      Store.save(db); render(); showToast('Đã chép hạn mức tháng trước');
    }
    else if (a === 'add') openForm(null);
    else if (a === 'edit') {
      var it = db.items.filter(function (t) { return t.id === b.getAttribute('data-id'); })[0];
      if (it) openForm(it);
    }
    else if (a === 'close-form') go(ui.back);
    else if (a === 'type') {
      ui.draft.type = b.getAttribute('data-type');
      if (ui.draft.type === 'thu') ui.draft.cat = '';
      delete ui.errors.cat; render();
    }
    else if (a === 'cat') { ui.draft.cat = b.getAttribute('data-cat'); delete ui.errors.cat; render(); }
    else if (a === 'filter') { ui.filter = b.getAttribute('data-f'); render(); }
    else if (a === 'clear-filter') { ui.filter = 'all'; ui.query = ''; render(); }
    else if (a === 'ask-delete') { ui.confirmDelete = true; render(); }
    else if (a === 'cancel-delete') { ui.confirmDelete = false; render(); }
    else if (a === 'do-delete') {
      var id = ui.draft.id, name = ui.draft.content;
      db.items = db.items.filter(function (t) { return t.id !== id; });
      Store.save(db);
      go(ui.back === 'list' ? 'list' : 'home');
      showToast('Đã xoá "' + name + '"');
    }
  });

  // Hạn mức: lưu ngay khi gõ, vẽ lại khi rời ô (để không ngắt việc gõ)
  function setBudget(k, raw) {
    var b = db.budgets[ui.month] || (db.budgets[ui.month] = {});
    if (raw) b[k] = +raw; else delete b[k];
    if (!Object.keys(b).length) delete db.budgets[ui.month];
    Store.save(db);
  }
  document.addEventListener('change', function (ev) {
    if (ev.target.getAttribute && ev.target.getAttribute('data-restore')) { readRestoreFile(ev.target.files[0]); return; }
    var k = ev.target.getAttribute && ev.target.getAttribute('data-budget');
    if (!k) return;
    render(); showToast('Đã lưu hạn mức ' + CATS[k].name);
  });

  document.addEventListener('input', function (ev) {
    var el = ev.target, k = el.getAttribute && el.getAttribute('data-input');
    var bk = el.getAttribute && el.getAttribute('data-budget');
    if (bk) {
      var braw = digits(el.value).replace(/^0+/, '').slice(0, 13);
      el.value = braw ? money(+braw) : '';
      setBudget(bk, braw);
      return;
    }
    if (!k) return;
    if (k === 'query') {
      // chỉ vẽ lại danh sách, không đụng ô tìm kiếm để gõ tiếng Việt (Telex/VNI) không bị ngắt
      ui.query = el.value;
      var lb = document.getElementById('list-body');
      if (lb) lb.innerHTML = listBody();
      return;
    }
    if (k === 'amount') {
      var raw = digits(el.value).replace(/^0+/, '').slice(0, 13);
      ui.draft.amount = raw;
      el.value = raw ? money(+raw) : '';
      if (ui.errors.amount) { delete ui.errors.amount; }
      return;
    }
    ui.draft[k] = el.value;
  });

  document.addEventListener('submit', function (ev) {
    if (ev.target.id === 'txform') { ev.preventDefault(); saveDraft(); }
  });

  render();
})();
