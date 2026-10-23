/* MyLife Hub (nâng cấp từ Sổ chi tiêu cá nhân)
   Giai đoạn 1: tab Hôm nay, Lịch tháng, Việc (thẻ, ưu tiên, hạn, lặp lại, đính kèm ảnh/tài liệu),
   Mua sắm (tự xếp nhóm, chuyển thành khoản chi: ghi riêng / gộp lần đi chợ / không ghi),
   nút (+) ghi nhanh, danh mục và nhóm hàng tự thêm, hoàn tiền khoản Công ty (T).
   Giữ nguyên các phần của Sổ chi tiêu: sổ thu chi, ngân sách 80%, báo cáo năm, xuất Excel theo mẫu cũ, sao lưu.
   Dữ liệu: xem store.js. */
(function () {
  'use strict';

  var Store = window.MyLifeStore, S = null;
  var WEEKDAYS = ['Chủ nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'];
  var WD_SHORT = ['T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'CN'];
  var MAX_AMOUNT = 999999999999;
  var BIG_FILE = 50 * 1024 * 1024; // tệp lớn hơn 50 MB: hỏi lại trước khi lưu
  var XLSX_URLS = ['vendor/xlsx.full.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];
  var TAGS = { work: 'Công việc', personal: 'Cá nhân', home: 'Việc nhà' };
  var TAG_CAT = { work: 'T', personal: 'C', home: 'G' }; // danh mục gợi ý cho chi phí của việc
  var PRIO = { 1: 'Cao', 2: 'Vừa', 3: 'Thấp' };
  var REPEATS = [['', 'Không lặp'], ['day', 'Hằng ngày'], ['week', 'Hằng tuần'], ['month', 'Hằng tháng'], ['monthEnd', 'Ngày cuối tháng']];
  var EXTRA_COLORS = ['#E8C27A', '#8FD694', '#F2A65A', '#A0B4F0', '#E58FB0', '#7FD1C7', '#C9A0E8', '#F07C7C'];
  var DOC_ACCEPT = '.pdf,.doc,.docx,.xls,.xlsx,.ppt,.pptx,.txt,.csv,.zip,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  // ---------- Danh mục (đọc từ dữ liệu, anh thêm được mục mới) ----------
  var CATS = {}, SPEND_CATS = [], ALL_SPEND = [];
  function refreshCats() {
    CATS = {};
    S.categories.sort(function (a, b) { return a.sort - b.sort; }).forEach(function (c) { CATS[c.code] = c; });
    ALL_SPEND = S.categories.filter(function (c) { return c.kind === 'chi'; }).map(function (c) { return c.code; });
    SPEND_CATS = ALL_SPEND.filter(function (k) { return !CATS[k].hidden; });
  }
  function catOf(k) { return CATS[k] || { code: k, name: k, short: k, color: 'var(--muted)', kind: 'chi' }; }

  // ---------- Tiện ích ----------
  function money(n) { return Math.round(n).toLocaleString('vi-VN').replace(/,/g, '.'); }
  function signed(t) { return (t.type === 'thu' ? '+' : '−') + money(t.amount); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function fold(s) {
    return String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/\s+/g, ' ').trim();
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function isoOf(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function todayISO() { return isoOf(new Date()); }
  function parseISO(iso) { var p = iso.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
  function addDays(iso, n) { var d = parseISO(iso); d.setDate(d.getDate() + n); return isoOf(d); }
  function monthOf(iso) { return iso.slice(0, 7); }
  function shiftMonth(ym, delta) {
    var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + delta;
    y += Math.floor(m / 12); m = ((m % 12) + 12) % 12;
    return y + '-' + pad(m + 1);
  }
  function monthLabel(ym) { return 'Tháng ' + (+ym.slice(5, 7)) + '/' + ym.slice(0, 4); }
  function dayLabel(iso) { return WEEKDAYS[parseISO(iso).getDay()] + ', ' + iso.slice(8, 10) + '/' + iso.slice(5, 7); }
  function shortDate(iso) { return iso.slice(8, 10) + '/' + iso.slice(5, 7); }
  function validDate(iso) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return false;
    var d = parseISO(iso), p = iso.split('-');
    return d.getFullYear() === +p[0] && d.getMonth() === +p[1] - 1 && d.getDate() === +p[2];
  }
  function digits(s) { return String(s).replace(/\D/g, ''); }
  function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function byId(list, id) { for (var i = 0; i < list.length; i++) if (list[i].id === id) return list[i]; return null; }
  function nowHHMM() { var d = new Date(); return pad(d.getHours()) + ':' + pad(d.getMinutes()); }
  function fmtSize(b) { return b >= 1048576 ? String(Math.round(b / 104857.6) / 10).replace('.', ',') + ' MB' : Math.max(1, Math.round(b / 1024)) + ' KB'; }
  function shortMoney(v) {
    if (v === 0) return '0';
    if (v >= 1e6) return String(Math.round(v / 1e5) / 10).replace('.', ',') + ' tr';
    return String(Math.round(v / 100) / 10).replace('.', ',') + ' k';
  }
  function cut(s, n) { return s.length > n ? s.slice(0, n - 1) + '…' : s; }

  // ---------- Ghi dữ liệu ----------
  function persist(ops) {
    return Store.commit(ops).catch(function () {
      showToast('Không lưu được vào bộ nhớ máy. Kiểm tra dung lượng trống rồi thử lại.');
    });
  }
  function upsert(store, rec) {
    var list = S[store], i = list.indexOf(byId(list, rec.id));
    if (i >= 0) list[i] = rec; else list.push(rec);
    return { store: store, put: rec };
  }
  function drop(store, id) {
    S[store] = S[store].filter(function (r) { return r.id !== id; });
    return { store: store, del: id };
  }
  function metaOp(key) { return { store: 'meta', key: key }; }

  // ---------- Khoản thu chi ----------
  function sortDesc(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return (b.createdAt || 0) - (a.createdAt || 0);
  }
  function itemsOfMonth(ym) { return S.expenses.filter(function (t) { return monthOf(t.date) === ym; }).sort(sortDesc); }
  function itemsOfYear(y) { return S.expenses.filter(function (t) { return t.date.slice(0, 4) === y; }); }
  function budgetOf(ym) { return S.budgets[ym] || {}; }
  function warnAt() { return S.settings.budgetWarn || 0.8; }
  function budgetStatus(spent, limit) {
    if (!limit) return { level: 'none', pct: 0 };
    var r = spent / limit;
    return { level: r > 1 ? 'over' : r >= warnAt() ? 'warn' : 'ok', pct: Math.round(r * 100) };
  }
  function budgetLine(k, spent, limit) {
    var st = budgetStatus(spent, limit);
    if (st.level === 'over') return catOf(k).name + ' đã vượt hạn mức ' + money(spent - limit) + ' đ';
    if (st.level === 'warn') return catOf(k).name + ' đã dùng ' + st.pct + '% hạn mức, còn ' + money(limit - spent) + ' đ';
    return '';
  }
  function totals(list) {
    var r = { thu: 0, chi: 0 };
    ALL_SPEND.forEach(function (k) { r[k] = 0; });
    list.forEach(function (t) {
      if (t.type === 'thu') r.thu += t.amount;
      else { r.chi += t.amount; r[t.cat] = (r[t.cat] || 0) + t.amount; }
    });
    r.con = r.thu - r.chi;
    return r;
  }
  function owedList() { return S.expenses.filter(function (t) { return t.reimburse === 'cho'; }).sort(sortDesc); }
  function attsOf(type, id) { return S.attachments.filter(function (a) { return a.ownerType === type && a.ownerId === id; }); }
  function latestMonth() {
    var t = monthOf(todayISO());
    if (S.expenses.some(function (x) { return monthOf(x.date) === t; }) || !S.expenses.length) return t;
    return monthOf(S.expenses.slice().sort(sortDesc)[0].date);
  }

  // ---------- Việc ----------
  function isOpen(t) { return t.status !== 'done'; }
  function taskSort(a, b) {
    var ad = a.due || '9999', bd = b.due || '9999';
    if (ad !== bd) return ad < bd ? -1 : 1;
    if (a.priority !== b.priority) return a.priority - b.priority;
    return (a.time || '99') < (b.time || '99') ? -1 : 1;
  }
  function dueText(t) {
    if (!t.due) return 'Không hạn';
    var today = todayISO(), s;
    if (t.due === today) s = 'Hôm nay';
    else if (t.due === addDays(today, 1)) s = 'Ngày mai';
    else if (t.due === addDays(today, -1)) s = 'Hôm qua';
    else s = shortDate(t.due) + (t.due.slice(0, 4) !== today.slice(0, 4) ? '/' + t.due.slice(0, 4) : '');
    return s + (t.time ? ' ' + t.time : '');
  }
  function lastDay(y, m) { return new Date(y, m + 1, 0).getDate(); }
  function stepDue(iso, rep, day) {
    var d = parseISO(iso), y = d.getFullYear(), m = d.getMonth();
    if (rep === 'day') return addDays(iso, 1);
    if (rep === 'week') return addDays(iso, 7);
    if (rep === 'month') return isoOf(new Date(y, m + 1, Math.min(day || d.getDate(), lastDay(y, m + 1))));
    if (rep === 'monthEnd') return isoOf(new Date(y, m + 2, 0));
    return iso;
  }
  // Lần tới của việc lặp lại: luôn từ hôm nay trở đi, kể cả khi làm trễ nhiều kỳ
  function nextDue(t) {
    var today = todayISO(), n = stepDue(t.due, t.repeat, t.repeatDay);
    while (n < today) n = stepDue(n, t.repeat, t.repeatDay);
    return n;
  }
  function repeatText(t) {
    if (!t.repeat) return '';
    if (t.repeat === 'week' && t.due) return 'Mỗi ' + WEEKDAYS[parseISO(t.due).getDay()];
    if (t.repeat === 'month' && t.due) return 'Ngày ' + (t.repeatDay || +t.due.slice(8, 10)) + ' hằng tháng';
    if (t.repeat === 'monthEnd') return 'Cuối mỗi tháng';
    return 'Hằng ngày';
  }

  // ---------- Mua sắm ----------
  function groupOf(id) { return byId(S.groups, id) || byId(S.groups, 'g-other') || S.groups[0]; }
  function groupFor(name) {
    var f = ' ' + fold(name) + ' ', best = null, len = 0;
    S.groups.forEach(function (g) {
      if (g.hidden) return;
      (g.keywords || []).forEach(function (kw) {
        if (kw.length > len && f.indexOf(' ' + kw + ' ') >= 0) { best = g; len = kw.length; }
      });
    });
    return (best || groupOf('g-other')).id;
  }
  function estOf(it) { return (it.price || 0) * (it.qty || 1); }
  function pendingShop() {
    return S.shopping.filter(function (it) { return it.status === 'bought' && !it.expenseId && it.choice !== 'skip'; });
  }

  // ---------- Trạng thái màn hình ----------
  var ui = {
    view: 'today', back: 'today', moneyTab: 'overview', month: null, filter: 'all', query: '',
    taskFilter: 'all', showDone: false, calMonth: null, calDay: null,
    draft: null, tdraft: null, sdraft: null, errors: {}, confirmDelete: false,
    sheet: null, viewer: null, attBusy: 0, lastSaved: null,
    year: null, exporting: false, exportMsg: '', restore: null, dataMsg: '', storageInfo: '', setMsg: ''
  };

  // ---------- Icon ----------
  function svg(d, size, sw) {
    return '<svg width="' + (size || 22) + '" height="' + (size || 22) + '" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="' + (sw || 2) + '" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + d + '</svg>';
  }
  var I = {
    prev: svg('<path d="M15 5l-7 7 7 7"/>'),
    next: svg('<path d="M9 5l7 7-7 7"/>'),
    close: svg('<path d="M6 6l12 12M18 6L6 18"/>'),
    today: svg('<circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'),
    cal: svg('<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 10h18M8 3v4M16 3v4"/>'),
    tasks: svg('<path d="M9 6h12M9 12h12M9 18h12"/><path d="M3 6l1.5 1.5L7 5M3 12l1.5 1.5L7 11M3 18l1.5 1.5L7 17"/>'),
    wallet: svg('<rect x="3" y="6" width="18" height="14" rx="2"/><path d="M3 10h18M16 15h2"/><path d="M6 6l9-3 1 3"/>'),
    plus: svg('<path d="M12 5v14M5 12h14"/>', 26, 2.5),
    small: svg('<path d="M12 5v14M5 12h14"/>', 18, 2.2),
    gear: svg('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>'),
    warn: svg('<path d="M12 3l10 18H2z"/><path d="M12 10v4M12 17v.5"/>', 18),
    download: svg('<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>', 20),
    search: svg('<circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/>', 18),
    clip: svg('<path d="M21 11l-8.6 8.6a5 5 0 0 1-7-7L14 4a3.3 3.3 0 0 1 4.7 4.7L10 17.3a1.7 1.7 0 0 1-2.4-2.4l8-8"/>', 14),
    repeat: svg('<path d="M17 2l4 4-4 4"/><path d="M3 11V9a3 3 0 0 1 3-3h15M7 22l-4-4 4-4"/><path d="M21 13v2a3 3 0 0 1-3 3H3"/>', 14),
    camera: svg('<path d="M4 8h3l2-3h6l2 3h3a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="13" r="3.5"/>', 20),
    image: svg('<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-9 9"/>', 20),
    file: svg('<path d="M14 3H6a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8z"/><path d="M14 3v5h5"/>', 20),
    check: svg('<path d="M5 12l5 5 9-10"/>', 16, 3),
    cart: svg('<circle cx="9" cy="20" r="1.5"/><circle cx="18" cy="20" r="1.5"/><path d="M2 3h3l2.5 12h11L21 7H6.2"/>', 18)
  };

  // ---------- Phần dùng chung ----------
  function topBar(title, left, right) {
    return '<header class="top">' + (left || '<span style="width:44px"></span>') + '<h1>' + title + '</h1>' + (right || '<span style="width:44px"></span>') + '</header>';
  }
  function monthHeader(title, action) {
    action = action || 'month';
    return '<header class="top">' +
      '<button class="icon-btn" data-action="' + action + '" data-delta="-1" aria-label="Tháng trước">' + I.prev + '</button>' +
      '<h1>' + esc(title) + '</h1>' +
      '<button class="icon-btn" data-action="' + action + '" data-delta="1" aria-label="Tháng sau">' + I.next + '</button>' +
      '</header>';
  }
  function txRow(t, showDate) {
    var c = catOf(t.cat), n = attsOf('expense', t.id).length;
    var bits = [];
    if (showDate) bits.push(shortDate(t.date));
    bits.push(t.note || c.name);
    if (t.reimburse === 'cho') bits.push('Chờ hoàn');
    return '<button class="tx" data-action="edit" data-id="' + esc(t.id) + '" aria-label="Sửa ' + esc(t.content) + '">' +
      '<span class="tag" style="background:' + c.color + '">' + esc(t.cat) + '</span>' +
      '<span class="mid"><span>' + esc(t.content) + (n ? ' <span class="clip">' + I.clip + '</span>' : '') + '</span><span class="sub">' + esc(bits.join(' · ')) + '</span></span>' +
      '<span class="amt num ' + (t.type === 'thu' ? 'in' : 'out') + '">' + signed(t) + '</span>' +
      '</button>';
  }
  function taskRow(t, opts) {
    opts = opts || {};
    var today = todayISO(), done = !isOpen(t), late = !done && t.due && t.due < today;
    var n = attsOf('task', t.id).length, sub = [];
    sub.push('<span class="' + (late ? 'late' : '') + '">' + esc(dueText(t)) + '</span>');
    if (!opts.hideTag) sub.push(esc(TAGS[t.tag] || ''));
    if (t.cost) sub.push(money(t.cost) + ' đ');
    if (t.repeat) sub.push('<span class="ico">' + I.repeat + '</span>');
    if (n) sub.push('<span class="ico">' + I.clip + n + '</span>');
    return '<div class="task-row' + (done ? ' done' : '') + '">' +
      '<button class="check p' + t.priority + '" data-action="task-toggle" data-id="' + esc(t.id) + '" aria-pressed="' + done + '" aria-label="' + (done ? 'Bỏ đánh dấu xong: ' : 'Đánh dấu xong: ') + esc(t.title) + '">' + (done ? I.check : '') + '</button>' +
      '<button class="task-main" data-action="task-edit" data-id="' + esc(t.id) + '"><span class="t">' + esc(t.title) + '</span><span class="sub">' + sub.join(' · ') + '</span></button>' +
      (opts.later && !done && t.due && t.due <= today ? '<button class="mini-btn" data-action="task-later" data-id="' + esc(t.id) + '">Để mai</button>' : '') +
      '</div>';
  }
  function emptyBox(text, btn) {
    return '<div class="empty"><p>' + text + '</p>' + (btn || '') + '</div>';
  }
  function attachBlock(type, id, title) {
    var list = attsOf(type, id);
    var items = list.map(function (a) {
      var img = /^image\//.test(a.mime);
      var ext = (a.name.split('.').pop() || 'tệp').slice(0, 4).toUpperCase();
      return '<div class="att">' +
        '<button type="button" class="att-open' + (img ? '' : ' doc') + '" data-action="att-open" data-id="' + esc(a.id) + '" aria-label="Mở ' + esc(a.name) + '">' +
          (img ? '<img data-att="' + esc(a.id) + '" alt="' + esc(a.name) + '">' : '<span class="ext">' + esc(ext) + '</span><span class="nm">' + esc(cut(a.name, 22)) + '</span>') +
        '</button>' +
        '<span class="att-size">' + fmtSize(a.size) + '</span>' +
        '<button type="button" class="att-del" data-action="att-del" data-id="' + esc(a.id) + '" aria-label="Xoá ' + esc(a.name) + '">' + svg('<path d="M6 6l12 12M18 6L6 18"/>', 14, 2.5) + '</button>' +
      '</div>';
    }).join('');
    function picker(kind, label, icon, accept, extra) {
      return '<label class="att-btn file-btn">' + icon + '<span>' + label + '</span>' +
        '<input type="file" accept="' + accept + '"' + extra + ' data-attach="' + kind + '" data-owner-type="' + type + '" data-owner-id="' + esc(id) + '"></label>';
    }
    return '<div class="field"><span>' + title + (list.length ? ' (' + list.length + ')' : '') + '</span>' +
      (list.length ? '<div class="att-grid">' + items + '</div>' : '') +
      (ui.attBusy ? '<p class="lead" role="status">Đang lưu ' + ui.attBusy + ' tệp…</p>' : '') +
      '<div class="att-btns">' +
        picker('camera', 'Chụp ảnh', I.camera, 'image/*', ' capture="environment"') +
        picker('image', 'Chọn ảnh', I.image, 'image/*', ' multiple') +
        (type === 'task' ? picker('doc', 'Tài liệu', I.file, DOC_ACCEPT, ' multiple') : '') +
      '</div></div>';
  }

  // ---------- Tab Hôm nay ----------
  function viewToday() {
    var today = todayISO(), ym = monthOf(today);
    var todays = S.expenses.filter(function (t) { return t.date === today; }).sort(sortDesc);
    var spentToday = 0, outCount = 0;
    todays.forEach(function (t) { if (t.type === 'chi') { spentToday += t.amount; outCount++; } });
    var s = totals(itemsOfMonth(ym)), bud = budgetOf(ym), limitSum = 0, spentIn = 0;
    SPEND_CATS.forEach(function (k) { if (bud[k]) { limitSum += bud[k]; spentIn += s[k]; } });
    var warnings = SPEND_CATS.map(function (k) { return budgetLine(k, s[k], bud[k]); }).filter(Boolean);
    var due = S.tasks.filter(function (t) { return isOpen(t) && t.due && t.due <= today; }).sort(taskSort);
    var prompts = S.tasks.filter(function (t) { return !isOpen(t) && t.cost && !t.expenseId && t.prompt === 'pending'; });
    var pend = pendingShop();
    var doneToday = S.tasks.filter(function (t) { return !isOpen(t) && t.doneAt && isoOf(new Date(t.doneAt)) === today; }).length;
    var tomorrow = S.tasks.filter(function (t) { return isOpen(t) && t.due === addDays(today, 1); }).length;

    var summary = nowHHMM() >= (S.settings.summaryTime || '22:00')
      ? '<section class="card summary" aria-label="Tổng kết ngày"><h2>Tổng kết ngày</h2>' +
        '<p>Đã chi <b class="num">' + money(spentToday) + ' đ</b> (' + outCount + ' khoản), xong <b>' + doneToday + '</b> việc.' +
        (due.length ? ' Còn <b>' + due.length + '</b> việc đến hạn chưa xong.' : '') +
        ' Ngày mai có <b>' + tomorrow + '</b> việc đến hạn.</p></section>'
      : '';

    var reminders = [];
    prompts.forEach(function (t) {
      reminders.push('<div class="rem"><span>Ghi <b class="num">' + money(t.cost) + ' đ</b> cho “' + esc(t.title) + '”?</span>' +
        '<span class="rem-btns"><button class="mini-btn gold" data-action="task-log" data-id="' + esc(t.id) + '">Ghi</button>' +
        '<button class="mini-btn" data-action="task-skip" data-id="' + esc(t.id) + '">Bỏ qua</button></span></div>');
    });
    if (pend.length) {
      var est = pend.reduce(function (a, it) { return a + estOf(it); }, 0);
      reminders.push('<div class="rem"><span>Đã mua <b>' + pend.length + '</b> món chưa ghi chi' + (est ? ' (ước tính ' + money(est) + ' đ)' : '') + '.</span>' +
        '<span class="rem-btns"><button class="mini-btn gold" data-action="shop-log">Ghi chi</button></span></div>');
    }
    var lb = S.settings.lastBackup;
    if (S.expenses.length > 20 && (!lb || lb < addDays(today, -30))) {
      reminders.push('<div class="rem"><span>' + (lb ? 'Lần sao lưu gần nhất: ' + shortDate(lb) + '/' + lb.slice(0, 4) + '.' : 'Anh chưa sao lưu dữ liệu lần nào.') + '</span>' +
        '<span class="rem-btns"><button class="mini-btn" data-action="go" data-view="settings">Sao lưu</button></span></div>');
    }

    var head = '<header class="top"><div class="day-title"><span class="lbl">Hôm nay</span><h1>' + esc(dayLabel(today)) + '</h1></div>' +
      '<button class="icon-btn" data-action="go" data-view="settings" aria-label="Cài đặt">' + I.gear + '</button></header>';

    return head + summary +
      '<section class="hero" aria-label="Tiền hôm nay">' +
        '<div><div class="lbl">Đã chi hôm nay</div><div class="big num">' + money(spentToday) + ' đ</div></div>' +
        '<div class="pair"><div><div class="lbl">Chi tháng này</div><b class="num">' + money(s.chi) + ' đ</b></div>' +
        (limitSum
          ? '<div><div class="lbl">Còn được chi</div><b class="num" style="color:' + (spentIn > limitSum ? 'var(--out)' : 'var(--fg)') + '">' + (spentIn > limitSum ? '−' : '') + money(Math.abs(limitSum - spentIn)) + ' đ</b></div>'
          : '<div><div class="lbl">Thu tháng này</div><b class="num">' + money(s.thu) + ' đ</b></div>') +
        '</div></section>' +
      (Store.warning ? '<p class="note">' + esc(Store.warning) + '</p>' : '') +
      (warnings.length ? '<section class="alert" role="status" aria-label="Cảnh báo ngân sách">' + warnings.map(function (w) {
        return '<div>' + I.warn + '<span>' + esc(w) + '</span></div>'; }).join('') +
        '<button class="link-btn" data-action="money-tab" data-tab="budget">Xem ngân sách</button></section>' : '') +
      (reminders.length ? '<section class="card"><h2>Nhắc nhở</h2>' + reminders.join('') + '</section>' : '') +
      '<section class="card"><h2>Việc đến hạn<button class="link-btn" data-action="go" data-view="tasks">Tất cả việc</button></h2>' +
        (due.length ? due.slice(0, 8).map(function (t) { return taskRow(t, { later: true }); }).join('') +
          (due.length > 8 ? '<p class="lead">Còn ' + (due.length - 8) + ' việc nữa trong tab Việc.</p>' : '')
          : emptyBox('Hôm nay không có việc nào đến hạn.', '<button class="btn ghost" data-action="task-new">Thêm việc</button>')) +
      '</section>' +
      '<section class="card"><h2>Thu chi hôm nay' + (todays.length ? '<span class="num muted">' + todays.length + ' khoản</span>' : '') + '</h2>' +
        (todays.length ? todays.map(function (t) { return txRow(t, false); }).join('') : emptyBox('Chưa ghi khoản nào hôm nay. Bấm nút + để ghi nhanh.')) +
      '</section>';
  }

  // ---------- Tab Lịch ----------
  function viewCalendar() {
    var ym = ui.calMonth, today = todayISO();
    var first = parseISO(ym + '-01'), offset = (first.getDay() + 6) % 7;
    var days = lastDay(first.getFullYear(), first.getMonth());
    var spend = {}, tasksBy = {};
    S.expenses.forEach(function (t) { if (t.type === 'chi' && monthOf(t.date) === ym) spend[t.date] = (spend[t.date] || 0) + t.amount; });
    S.tasks.forEach(function (t) { if (t.due && monthOf(t.due) === ym) (tasksBy[t.due] = tasksBy[t.due] || []).push(t); });
    var cells = WD_SHORT.map(function (w) { return '<span class="wd">' + w + '</span>'; }).join('');
    for (var i = 0; i < offset; i++) cells += '<span></span>';
    for (var d = 1; d <= days; d++) {
      var iso = ym + '-' + pad(d), tl = tasksBy[iso] || [];
      var open = tl.filter(isOpen).length;
      var cls = 'cd' + (iso === today ? ' today' : '') + (iso === ui.calDay ? ' sel' : '');
      var label = dayLabel(iso) + (spend[iso] ? ', chi ' + money(spend[iso]) + ' đồng' : '') + (tl.length ? ', ' + tl.length + ' việc' : '');
      cells += '<button class="' + cls + '" data-action="cal-day" data-d="' + iso + '" aria-label="' + esc(label) + '"' + (iso === ui.calDay ? ' aria-current="date"' : '') + '>' +
        '<span class="dn">' + d + '</span>' +
        (spend[iso] ? '<span class="ds num">' + shortMoney(spend[iso]).replace(' ', '') + '</span>' : '<span class="ds"></span>') +
        (tl.length ? '<i class="dot' + (open ? '' : ' ok') + '"></i>' : '') + '</button>';
    }
    var sel = ui.calDay;
    var dayTasks = S.tasks.filter(function (t) { return t.due === sel; }).sort(taskSort);
    var dayTx = S.expenses.filter(function (t) { return t.date === sel; }).sort(sortDesc);
    var daySpent = dayTx.reduce(function (a, t) { return a + (t.type === 'chi' ? t.amount : 0); }, 0);
    var monthSpent = Object.keys(spend).reduce(function (a, k) { return a + spend[k]; }, 0);

    return monthHeader(monthLabel(ym), 'cal-month') +
      '<p class="lead">Chi cả tháng: <b class="num" style="color:var(--fg)">' + money(monthSpent) + ' đ</b>. Ô ngày ghi số tiền đã chi; chấm vàng là có việc chưa xong, chấm xanh là việc đã xong.</p>' +
      '<section class="card cal" aria-label="Lịch ' + esc(monthLabel(ym)) + '"><div class="cal-grid">' + cells + '</div></section>' +
      '<section class="card"><h2>' + esc(dayLabel(sel)) + (daySpent ? '<span class="num muted">Chi ' + money(daySpent) + ' đ</span>' : '') + '</h2>' +
        (dayTasks.length ? dayTasks.map(function (t) { return taskRow(t, {}); }).join('') : '') +
        (dayTx.length ? dayTx.map(function (t) { return txRow(t, false); }).join('') : '') +
        (!dayTasks.length && !dayTx.length ? '<p class="lead">Ngày này chưa có việc hay khoản thu chi nào.</p>' : '') +
        '<div class="row2"><button class="btn ghost" data-action="task-new" data-due="' + sel + '">Thêm việc</button>' +
        '<button class="btn ghost" data-action="add" data-date="' + sel + '">Ghi khoản</button></div>' +
      '</section>' +
      '<p class="lead center">Lịch hẹn và hoá đơn định kỳ sẽ có ở giai đoạn 2.</p>';
  }

  // ---------- Tab Việc ----------
  function viewTasks() {
    var today = todayISO();
    var f = ui.taskFilter;
    var openCount = function (tag) { return S.tasks.filter(function (t) { return isOpen(t) && (!tag || t.tag === tag); }).length; };
    var chips = [['all', 'Tất cả', openCount()], ['work', 'Công việc', openCount('work')], ['personal', 'Cá nhân', openCount('personal')],
      ['home', 'Việc nhà', openCount('home')], ['shop', 'Mua sắm', S.shopping.filter(function (i) { return i.status === 'need'; }).length]]
      .map(function (c) {
        return '<button class="chip" data-action="task-filter" data-f="' + c[0] + '" aria-pressed="' + (f === c[0]) + '">' + c[1] + (c[2] ? ' <span class="cnt">' + c[2] + '</span>' : '') + '</button>';
      }).join('');
    var head = topBar(f === 'shop' ? 'Mua sắm' : 'Việc', null,
      '<button class="icon-btn" data-action="' + (f === 'shop' ? 'shop-focus' : 'task-new') + '" aria-label="' + (f === 'shop' ? 'Thêm món cần mua' : 'Thêm việc') + '">' + I.small + '</button>') +
      '<div class="chips scroll" role="group" aria-label="Lọc việc">' + chips + '</div>';
    if (f === 'shop') return head + viewShopping();

    var list = S.tasks.filter(function (t) { return f === 'all' || t.tag === f; });
    var open = list.filter(isOpen).sort(taskSort);
    var groups = [
      ['Quá hạn', open.filter(function (t) { return t.due && t.due < today; })],
      ['Hôm nay', open.filter(function (t) { return t.due === today; })],
      ['Sắp tới', open.filter(function (t) { return t.due && t.due > today; })],
      ['Chưa có hạn', open.filter(function (t) { return !t.due; })]
    ];
    var html = groups.filter(function (g) { return g[1].length; }).map(function (g) {
      return '<section class="card"><h2>' + g[0] + '<span class="num muted">' + g[1].length + '</span></h2>' +
        g[1].map(function (t) { return taskRow(t, { hideTag: f !== 'all', later: g[0] === 'Quá hạn' }); }).join('') + '</section>';
    }).join('');
    var done = list.filter(function (t) { return !isOpen(t); }).sort(function (a, b) { return (b.doneAt || 0) - (a.doneAt || 0); });
    if (!open.length) html = '<section class="card">' + emptyBox(list.length ? 'Đã xong hết việc. Tuyệt!' : 'Chưa có việc nào. Ghi việc cần làm, đặt hạn và mức ưu tiên.', '<button class="btn ghost" data-action="task-new">Thêm việc</button>') + '</section>';
    if (done.length) {
      html += '<section class="card"><h2>Đã xong<button class="link-btn" data-action="toggle-done">' + (ui.showDone ? 'Ẩn' : 'Hiện ' + Math.min(done.length, 30)) + '</button></h2>' +
        (ui.showDone ? done.slice(0, 30).map(function (t) { return taskRow(t, { hideTag: f !== 'all' }); }).join('') : '') + '</section>';
    }
    return head + html;
  }

  function viewShopping() {
    var need = S.shopping.filter(function (i) { return i.status === 'need'; });
    var pend = pendingShop();
    var since = addDays(todayISO(), -14);
    var bought = S.shopping.filter(function (i) { return i.status === 'bought' && i.boughtOn >= since; })
      .sort(function (a, b) { return a.boughtOn < b.boughtOn ? 1 : -1; });
    var html = '<form class="shop-add" id="shopadd"><label class="sr" for="shop-new">Thêm món cần mua</label>' +
      '<input id="shop-new" autocomplete="off" maxlength="300" placeholder="Ví dụ: thịt heo, rau muống, 2 sữa tươi">' +
      '<button class="btn" type="submit">Thêm</button></form>';
    if (pend.length) {
      var est = pend.reduce(function (a, it) { return a + estOf(it); }, 0);
      html += '<section class="pend" aria-label="Món đã mua chưa ghi chi"><div><b>Đã mua ' + pend.length + ' món</b>' + (est ? ' · Ước tính <span class="num">' + money(est) + ' đ</span>' : '') +
        '<div class="sub">' + esc(cut(pend.map(function (i) { return i.name; }).join(', '), 90)) + '</div></div>' +
        '<button class="btn" data-action="shop-log">Ghi chi</button></section>';
    }
    var gids = [];
    S.groups.forEach(function (g) { if (need.some(function (i) { return i.group === g.id; })) gids.push(g.id); });
    need.forEach(function (i) { if (gids.indexOf(i.group) < 0 && !byId(S.groups, i.group)) gids.push(i.group); });
    if (!need.length) html += '<section class="card">' + emptyBox('Chưa có món nào cần mua. Gõ tên món ở ô trên, nhiều món cách nhau bằng dấu phẩy.') + '</section>';
    gids.forEach(function (gid) {
      var g = groupOf(gid), items = need.filter(function (i) { return groupOf(i.group).id === g.id; });
      var est = items.reduce(function (a, it) { return a + estOf(it); }, 0);
      html += '<section class="card"><h2><span>' + esc(g.name) + ' <span class="tag sm" style="background:' + catOf(g.cat).color + '">' + esc(g.cat) + '</span></span>' +
        (est ? '<span class="num muted">~' + money(est) + ' đ</span>' : '') + '</h2>' +
        items.map(shopRow).join('') + '</section>';
    });
    if (bought.length) {
      html += '<section class="card"><h2>Mua gần đây<span class="num muted">14 ngày</span></h2>' + bought.map(shopRow).join('') + '</section>';
    }
    html += '<button class="link-btn center" data-action="go" data-view="settings">Quản lý nhóm hàng</button>';
    return html;
  }

  function shopRow(it) {
    var done = it.status === 'bought', sub = [];
    if (it.qty > 1) sub.push('× ' + it.qty);
    if (it.price) sub.push('~' + money(estOf(it)) + ' đ');
    if (done) {
      sub.unshift(shortDate(it.boughtOn));
      sub.push(it.expenseId ? 'đã ghi chi' : it.choice === 'skip' ? 'không ghi chi' : 'chưa ghi chi');
    }
    return '<div class="task-row' + (done ? ' done' : '') + '">' +
      '<button class="check" data-action="shop-toggle" data-id="' + esc(it.id) + '" aria-pressed="' + done + '" aria-label="' + (done ? 'Bỏ đánh dấu đã mua: ' : 'Đánh dấu đã mua: ') + esc(it.name) + '">' + (done ? I.check : '') + '</button>' +
      '<button class="task-main" data-action="shop-edit" data-id="' + esc(it.id) + '"><span class="t">' + esc(it.name) + '</span>' + (sub.length ? '<span class="sub">' + esc(sub.join(' · ')) + '</span>' : '') + '</button>' +
      (done ? '<button class="mini-btn" data-action="shop-again" data-id="' + esc(it.id) + '">Mua lại</button>' : '') +
      '</div>';
  }

  // ---------- Form Việc ----------
  function kindSeg(cur) {
    return '<div class="seg seg3" role="group" aria-label="Loại cần thêm">' +
      '<button type="button" data-action="add-kind" data-k="expense" aria-pressed="' + (cur === 'expense') + '">Khoản thu chi</button>' +
      '<button type="button" data-action="add-kind" data-k="task" aria-pressed="' + (cur === 'task') + '">Việc</button>' +
      '<button type="button" data-action="add-kind" data-k="shop" aria-pressed="' + (cur === 'shop') + '">Món cần mua</button></div>';
  }

  function viewTaskForm() {
    var d = ui.tdraft, e = ui.errors, isEdit = !d.isNew;
    var tagBtns = Object.keys(TAGS).map(function (k) {
      return '<button type="button" data-action="t-set" data-f="tag" data-v="' + k + '" aria-pressed="' + (d.tag === k) + '">' + TAGS[k] + '</button>';
    }).join('');
    var prioBtns = [1, 2, 3].map(function (p) {
      return '<button type="button" class="pr' + p + '" data-action="t-set" data-f="priority" data-v="' + p + '" aria-pressed="' + (+d.priority === p) + '">' + PRIO[p] + '</button>';
    }).join('');
    var today = todayISO();
    var dueChips = [['Hôm nay', today], ['Ngày mai', addDays(today, 1)], ['Không hạn', '']].map(function (c) {
      return '<button type="button" class="chip" data-action="t-set" data-f="due" data-v="' + c[1] + '" aria-pressed="' + (d.due === c[1]) + '">' + c[0] + '</button>';
    }).join('');
    var repOpts = REPEATS.map(function (r) { return '<option value="' + r[0] + '"' + (d.repeat === r[0] ? ' selected' : '') + '>' + r[1] + '</option>'; }).join('');
    var catOpts = SPEND_CATS.map(function (k) { return '<option value="' + k + '"' + (d.costCat === k ? ' selected' : '') + '>' + k + ' · ' + esc(catOf(k).name) + '</option>'; }).join('');
    var orig = byId(S.tasks, d.id), ex = orig && orig.expenseId ? byId(S.expenses, orig.expenseId) : null;
    var confirm = ui.confirmDelete
      ? '<div class="confirm" role="alert"><span>Xoá việc “' + esc(d.title) + '”' + (attsOf('task', d.id).length ? ' cùng ' + attsOf('task', d.id).length + ' tệp đính kèm' : '') + '? Không thể hoàn tác.</span>' +
        '<div class="row"><button type="button" class="btn ghost" data-action="cancel-delete">Giữ lại</button>' +
        '<button type="button" class="btn danger solid" data-action="task-delete">Xoá</button></div></div>' : '';

    return topBar(isEdit ? 'Sửa việc' : 'Thêm việc', '<button class="icon-btn" data-action="close-form" aria-label="Đóng">' + I.close + '</button>') +
      '<form class="form" id="taskform" novalidate>' +
        (isEdit ? '' : kindSeg('task')) +
        '<label class="field" for="t-title">Tên việc' +
          '<input id="t-title" maxlength="140" autocomplete="off" placeholder="Ví dụ: Báo cáo sản lượng nước tháng 10" value="' + esc(d.title) + '" data-tinput="title"' + (e.title ? ' aria-invalid="true"' : '') + '>' +
          (e.title ? '<span class="err">' + e.title + '</span>' : '') + '</label>' +
        '<div class="field"><span>Thẻ</span><div class="seg seg3" role="group" aria-label="Thẻ">' + tagBtns + '</div></div>' +
        '<div class="field"><span>Mức ưu tiên</span><div class="seg seg3 prio" role="group" aria-label="Mức ưu tiên">' + prioBtns + '</div></div>' +
        '<div class="field"><span>Hạn</span><div class="chips">' + dueChips + '</div>' +
          '<div class="row2"><input id="t-due" type="date" aria-label="Ngày hạn" value="' + esc(d.due) + '" data-tinput="due"' + (e.due ? ' aria-invalid="true"' : '') + '>' +
          '<input id="t-time" type="time" aria-label="Giờ (không bắt buộc)" value="' + esc(d.time) + '" data-tinput="time"></div>' +
          (e.due ? '<span class="err">' + e.due + '</span>' : '') + '</div>' +
        '<label class="field" for="t-repeat">Lặp lại' +
          '<select id="t-repeat" data-tinput="repeat">' + repOpts + '</select>' +
          (d.repeat && d.due ? '<span class="hint">' + esc(repeatText({ repeat: d.repeat, due: d.repeat === 'monthEnd' ? d.due : d.due, repeatDay: +d.due.slice(8, 10) })) + '. Tích xong thì app tự đặt lần tới.</span>' : '') + '</label>' +
        '<label class="field" for="t-notes">Ghi chú' +
          '<textarea id="t-notes" rows="3" maxlength="2000" data-tinput="notes">' + esc(d.notes) + '</textarea></label>' +
        '<div class="field"><span>Chi phí dự kiến (không bắt buộc)</span>' +
          '<div class="row2"><input id="t-cost" class="num" inputmode="numeric" autocomplete="off" placeholder="0 đ" aria-label="Chi phí dự kiến" value="' + (d.cost ? money(+d.cost) : '') + '" data-tinput="cost"' + (e.cost ? ' aria-invalid="true"' : '') + '>' +
          '<select id="t-costcat" aria-label="Danh mục chi" data-tinput="costCat">' + catOpts + '</select></div>' +
          (e.cost ? '<span class="err">' + e.cost + '</span>' : '<span class="hint">Có chi phí thì khi tích xong, app hỏi ghi thành khoản chi.</span>') + '</div>' +
        attachBlock('task', d.id, 'Tài liệu, ảnh đính kèm') +
        (ex ? '<p class="src">Đã ghi chi ' + money(ex.amount) + ' đ ngày ' + shortDate(ex.date) + '.</p>' : '') +
        '<div class="form-actions">' +
          '<button type="submit" class="btn">' + (isEdit ? 'Lưu thay đổi' : 'Lưu việc') + '</button>' +
          (isEdit && orig ? '<button type="button" class="btn ghost" data-action="task-toggle" data-id="' + esc(d.id) + '" data-from="form">' + (isOpen(orig) ? 'Đánh dấu đã xong' : 'Chuyển về chưa xong') + '</button>' : '') +
          (isEdit && !ui.confirmDelete ? '<button type="button" class="btn danger" data-action="ask-delete">Xoá việc này</button>' : '') +
          confirm +
        '</div></form>';
  }

  // ---------- Form món cần mua ----------
  function viewShopItem() {
    var d = ui.sdraft, e = ui.errors;
    var opts = S.groups.filter(function (g) { return !g.hidden || g.id === d.group; }).map(function (g) {
      return '<option value="' + esc(g.id) + '"' + (d.group === g.id ? ' selected' : '') + '>' + esc(g.name) + ' (' + esc(g.cat) + ')</option>';
    }).join('') + '<option value="__new"' + (d.group === '__new' ? ' selected' : '') + '>+ Nhóm mới…</option>';
    var catOpts = SPEND_CATS.map(function (k) { return '<option value="' + k + '"' + (d.newCat === k ? ' selected' : '') + '>' + k + ' · ' + esc(catOf(k).name) + '</option>'; }).join('');
    var confirm = ui.confirmDelete
      ? '<div class="confirm" role="alert"><span>Xoá “' + esc(d.name) + '” khỏi danh sách?</span>' +
        '<div class="row"><button type="button" class="btn ghost" data-action="cancel-delete">Giữ lại</button>' +
        '<button type="button" class="btn danger solid" data-action="shop-delete">Xoá</button></div></div>' : '';
    return topBar('Sửa món', '<button class="icon-btn" data-action="close-form" aria-label="Đóng">' + I.close + '</button>') +
      '<form class="form" id="shopform" novalidate>' +
        '<label class="field" for="s-name">Tên món<input id="s-name" maxlength="80" autocomplete="off" value="' + esc(d.name) + '" data-sinput="name"' + (e.name ? ' aria-invalid="true"' : '') + '>' +
          (e.name ? '<span class="err">' + e.name + '</span>' : '') + '</label>' +
        '<div class="row2"><label class="field" for="s-qty">Số lượng<input id="s-qty" inputmode="numeric" class="num" value="' + esc(d.qty) + '" data-sinput="qty"></label>' +
        '<label class="field" for="s-price">Giá ước tính 1 món<input id="s-price" inputmode="numeric" class="num" placeholder="0 đ" value="' + (d.price ? money(+d.price) : '') + '" data-sinput="price"></label></div>' +
        '<label class="field" for="s-group">Nhóm hàng<select id="s-group" data-sinput="group">' + opts + '</select>' +
          '<span class="hint">Đổi nhóm thì lần sau app tự xếp món này vào nhóm mới.</span></label>' +
        (d.group === '__new' ? '<div class="row2"><label class="field" for="s-gname">Tên nhóm mới<input id="s-gname" maxlength="40" value="' + esc(d.newName) + '" data-sinput="newName"' + (e.newName ? ' aria-invalid="true"' : '') + '></label>' +
          '<label class="field" for="s-gcat">Ghi chi vào<select id="s-gcat" data-sinput="newCat">' + catOpts + '</select></label></div>' +
          (e.newName ? '<span class="err">' + e.newName + '</span>' : '') : '') +
        '<div class="form-actions"><button type="submit" class="btn">Lưu</button>' +
          (!ui.confirmDelete ? '<button type="button" class="btn danger" data-action="ask-delete">Xoá món này</button>' : '') + confirm + '</div>' +
      '</form>';
  }

  // ---------- Form ghi nhanh khoản thu chi ----------
  function viewForm() {
    var d = ui.draft, e = ui.errors, isEdit = !d.isNew;
    var shown = SPEND_CATS.slice();
    if (d.cat && CATS[d.cat] && CATS[d.cat].hidden && shown.indexOf(d.cat) < 0) shown.push(d.cat);
    shown.push('+');
    var catBtns = '<div class="cats" role="group" aria-label="Danh mục">' + shown.map(function (k) {
      return '<button type="button" class="cat" data-action="cat" data-cat="' + esc(k) + '" aria-pressed="' + (d.cat === k) + '">' +
        '<b style="color:' + catOf(k).color + '">' + esc(k) + '</b>' + esc(catOf(k).short || catOf(k).name) + '</button>';
    }).join('') + '</div>';
    var amountShown = d.amountText || (d.amount ? money(+d.amount) : '');
    var today = todayISO(), yest = addDays(today, -1);
    var src = '';
    if (d.source === 'shopping') src = 'Từ Mua sắm: ' + d.sourceIds.length + ' món';
    else if (d.source === 'task') { var tk = byId(S.tasks, d.sourceIds[0]); src = 'Từ việc: ' + (tk ? tk.title : '(đã xoá)'); }
    var reimb = CATS[d.cat] && CATS[d.cat].reimbursable;
    var confirm = ui.confirmDelete
      ? '<div class="confirm" role="alert"><span>Xoá khoản “' + esc(d.content || catOf(d.cat).name) + '”? Không thể hoàn tác.</span>' +
        '<div class="row"><button type="button" class="btn ghost" data-action="cancel-delete">Giữ lại</button>' +
        '<button type="button" class="btn danger solid" data-action="do-delete">Xoá</button></div></div>' : '';

    return topBar(isEdit ? 'Sửa khoản' : 'Ghi nhanh', '<button class="icon-btn" data-action="close-form" aria-label="Đóng">' + I.close + '</button>') +
      '<form class="form" id="txform" novalidate>' +
        (!isEdit && d.source === 'manual' ? kindSeg('expense') : '') +
        (src ? '<p class="src">' + esc(src) + '</p>' : '') +
        '<label class="field" for="f-amount">Số tiền (đồng)' +
          '<input id="f-amount" class="amount num" inputmode="decimal" autocomplete="off" placeholder="0" value="' + esc(amountShown) + '" data-input="amount"' + (e.amount ? ' aria-invalid="true" aria-describedby="e-amount"' : '') + '>' +
          (e.amount ? '<span class="err" id="e-amount">' + e.amount + '</span>' : '') + '</label>' +
        '<div class="quick" role="group" aria-label="Gõ nhanh số tiền">' +
          '<button type="button" data-action="amt" data-op="000">000</button>' +
          '<button type="button" data-action="amt" data-op="k">× nghìn</button>' +
          '<button type="button" data-action="amt" data-op="m">× triệu</button></div>' +
        '<div class="field"><span>Danh mục (KH)</span>' + catBtns + (e.cat ? '<span class="err">' + e.cat + '</span>' : '') + '</div>' +
        '<label class="field" for="f-content">Nội dung' +
          '<input id="f-content" maxlength="120" autocomplete="off" placeholder="Không bắt buộc. Ví dụ: Cafe, đổ xăng" value="' + esc(d.content) + '" data-input="content"></label>' +
        '<div id="sugg" class="sugg" aria-live="polite">' + suggHTML() + '</div>' +
        '<div class="field"><span>Ngày</span><div class="chips">' +
          '<button type="button" class="chip" data-action="f-date" data-v="' + today + '" aria-pressed="' + (d.date === today) + '">Hôm nay</button>' +
          '<button type="button" class="chip" data-action="f-date" data-v="' + yest + '" aria-pressed="' + (d.date === yest) + '">Hôm qua</button>' +
          '<input id="f-date" type="date" aria-label="Chọn ngày khác" value="' + esc(d.date) + '" data-input="date"' + (e.date ? ' aria-invalid="true"' : '') + '></div>' +
          (e.date ? '<span class="err">' + e.date + '</span>' : '') + '</div>' +
        (reimb ? '<label class="check-line"><input type="checkbox" data-input="reimbursed"' + (d.reimbursed ? ' checked' : '') + '> Công ty đã hoàn tiền khoản này</label>' : '') +
        '<label class="field" for="f-note">Ghi chú thêm (không bắt buộc)' +
          '<input id="f-note" maxlength="200" autocomplete="off" value="' + esc(d.note) + '" data-input="note"></label>' +
        attachBlock('expense', d.id, 'Ảnh hoá đơn, chứng từ') +
        '<div class="form-actions">' +
          '<button type="submit" class="btn">' + (isEdit ? 'Lưu thay đổi' : 'Lưu') + '</button>' +
          (isEdit && !ui.confirmDelete ? '<button type="button" class="btn danger" data-action="ask-delete">Xoá khoản này</button>' : '') +
          confirm +
        '</div></form>';
  }

  // Gợi ý nội dung cũ: chọn là điền luôn danh mục và số tiền hay dùng
  var suggList = [];
  function suggHTML() {
    var d = ui.draft, q = fold(d.content || '');
    suggList = [];
    if (!q) return '';
    var seen = {};
    S.expenses.slice().sort(sortDesc).some(function (t) {
      var f = fold(t.content);
      if (seen[f] || f === q || f.indexOf(q) < 0) return false;
      seen[f] = 1; suggList.push(t);
      return suggList.length >= 5;
    });
    return suggList.map(function (t, i) {
      return '<button type="button" class="chip" data-action="pick-sugg" data-i="' + i + '">' + esc(t.content) + ' <span class="muted num">' + t.cat + ' · ' + money(t.amount) + '</span></button>';
    }).join('');
  }

  // ---------- Tab Tiền ----------
  function moneyNav() {
    var tabs = [['overview', 'Tổng quan'], ['list', 'Sổ thu chi'], ['budget', 'Ngân sách'], ['report', 'Báo cáo']];
    return '<nav class="subnav" aria-label="Mục trong tab Tiền">' + tabs.map(function (t) {
      return '<button data-action="money-tab" data-tab="' + t[0] + '"' + (ui.moneyTab === t[0] ? ' aria-current="page"' : '') + '>' + t[1] + '</button>';
    }).join('') + '</nav>';
  }
  function emptyMonth() {
    return emptyBox('Chưa có khoản thu chi nào trong ' + esc(monthLabel(ui.month).toLowerCase()) + '.', '<button class="btn" data-action="add">Thêm khoản đầu tiên</button>');
  }

  function viewOverview() {
    var list = itemsOfMonth(ui.month), s = totals(list), bud = budgetOf(ui.month);
    var show = ALL_SPEND.filter(function (k) { return !CATS[k].hidden || s[k]; });
    var maxCat = Math.max.apply(null, [1].concat(show.map(function (k) { return s[k]; })));
    var warnings = SPEND_CATS.map(function (k) { return budgetLine(k, s[k], bud[k]); }).filter(Boolean);
    var cats = show.map(function (k) {
      var pct = s.chi ? Math.round(s[k] / s.chi * 1000) / 10 : 0;
      var w = s[k] ? Math.max(2, s[k] / maxCat * 100) : 0;
      var lim = bud[k] ? ' / ' + money(bud[k]) : '';
      return '<div class="cat-row"><div class="t"><span>' + esc(k) + ' · ' + esc(catOf(k).name) + '</span>' +
        '<span class="num">' + money(s[k]) + lim + ' đ <span class="muted">(' + String(pct).replace('.', ',') + '%)</span></span></div>' +
        '<div class="bar" role="img" aria-label="' + esc(catOf(k).name) + ' ' + String(pct).replace('.', ',') + '% tổng chi"><span style="width:' + w + '%;background:' + catOf(k).color + '"></span></div></div>';
    }).join('');
    var recent = list.slice(0, 5).map(function (t) { return txRow(t, true); }).join('');
    var owed = owedList(), owedSum = owed.reduce(function (a, t) { return a + t.amount; }, 0);

    return monthHeader(monthLabel(ui.month)) +
      '<section class="hero" aria-label="Tóm tắt tháng">' +
        '<div><div class="lbl">Còn lại</div><div class="big num">' + money(s.con) + ' đ</div></div>' +
        '<div class="pair"><div><div class="lbl">Tổng thu (+)</div><b class="num">' + money(s.thu) + ' đ</b></div>' +
        '<div><div class="lbl">Tổng chi (−)</div><b class="num">' + money(s.chi) + ' đ</b></div></div>' +
      '</section>' +
      (owed.length ? '<section class="card owed"><h2>Công ty còn nợ<span class="num">' + money(owedSum) + ' đ</span></h2>' +
        '<p class="lead">' + owed.length + ' khoản mục T đang chờ công ty hoàn tiền.</p>' +
        '<button class="btn ghost" data-action="show-owed">Xem các khoản</button></section>' : '') +
      (warnings.length ? '<section class="alert" role="status" aria-label="Cảnh báo ngân sách">' + warnings.map(function (w) {
        return '<div>' + I.warn + '<span>' + esc(w) + '</span></div>'; }).join('') +
        '<button class="link-btn" data-action="money-tab" data-tab="budget">Xem ngân sách</button></section>' : '') +
      '<section class="card"><h2>Chi theo danh mục</h2>' + cats + '</section>' +
      '<section class="card"><h2>Gần đây' + (list.length ? '<button class="link-btn" data-action="money-tab" data-tab="list">Xem tất cả (' + list.length + ')</button>' : '') + '</h2>' +
        (list.length ? '<div>' + recent + '</div>' : emptyMonth()) +
      '</section>';
  }

  function listItems() {
    if (ui.filter === 'owed') return owedList();
    return itemsOfMonth(ui.month);
  }
  function listBody() {
    var list = listItems();
    var q = fold(ui.query);
    var shown = list.filter(function (t) {
      if (ui.filter === 'thu' && t.type !== 'thu') return false;
      if (ALL_SPEND.indexOf(ui.filter) >= 0 && t.cat !== ui.filter) return false;
      if (q && fold(t.content + ' ' + (t.note || '')).indexOf(q) < 0) return false;
      return true;
    });
    if (!list.length) return ui.filter === 'owed' ? emptyBox('Không còn khoản nào chờ hoàn tiền.') : emptyMonth();
    if (!shown.length) return emptyBox('Không có khoản nào khớp bộ lọc.', '<button class="btn ghost" data-action="clear-filter">Bỏ lọc</button>');

    var groups = '', curDay = null, buf = [], dayNet = 0;
    function flush() {
      if (!curDay) return;
      groups += '<div class="day-head"><b>' + esc(dayLabel(curDay)) + (ui.filter === 'owed' ? '/' + curDay.slice(0, 4) : '') + '</b><span class="num">' + (dayNet >= 0 ? '+' : '−') + money(Math.abs(dayNet)) + '</span></div>' + buf.join('');
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
    var owedMode = ui.filter === 'owed';
    var list = itemsOfMonth(ui.month), s = totals(list);
    var chipDefs = [['all', 'Tất cả'], ['thu', 'Thu']].concat(ALL_SPEND.filter(function (k) { return !CATS[k].hidden || s[k]; }).map(function (k) {
      return [k, k + ' · ' + (catOf(k).short || catOf(k).name)];
    }));
    if (owedList().length || owedMode) chipDefs.push(['owed', 'Chờ hoàn tiền']);
    var chips = chipDefs.map(function (c) {
      return '<button class="chip" data-action="filter" data-f="' + esc(c[0]) + '" aria-pressed="' + (ui.filter === c[0]) + '">' + esc(c[1]) + '</button>';
    }).join('');
    var owed = owedList(), owedSum = owed.reduce(function (a, t) { return a + t.amount; }, 0);
    return (owedMode ? topBar('Khoản chờ công ty hoàn tiền') : monthHeader('Sổ thu chi · ' + monthLabel(ui.month))) +
      '<label class="search">' + I.search +
        '<input id="q" type="search" placeholder="Tìm theo nội dung hoặc ghi chú" aria-label="Tìm theo nội dung hoặc ghi chú" value="' + esc(ui.query) + '" data-input="query"></label>' +
      '<div class="chips" role="group" aria-label="Lọc theo danh mục">' + chips + '</div>' +
      (owedMode
        ? (owed.length ? '<div class="mini two"><div>Tổng chờ hoàn<b class="num" style="color:var(--warn-fg)">' + money(owedSum) + ' đ</b></div><div>Số khoản<b class="num">' + owed.length + '</b></div></div>' +
          '<button class="btn ghost" data-action="ask-reimburse-all">Đánh dấu tất cả đã hoàn</button>' : '')
        : '<div class="mini"><div>Thu<b class="num" style="color:var(--in)">' + money(s.thu) + '</b></div>' +
          '<div>Chi<b class="num" style="color:var(--out)">' + money(s.chi) + '</b></div>' +
          '<div>Còn lại<b class="num">' + money(s.con) + '</b></div></div>') +
      '<section class="card" style="gap:0" id="list-body" aria-label="Danh sách khoản thu chi">' + listBody() + '</section>' +
      (list.length || owedMode ? '<p class="lead center">Chạm vào một dòng để sửa hoặc xoá.</p>' : '');
  }

  function viewBudget() {
    var list = itemsOfMonth(ui.month), s = totals(list), bud = budgetOf(ui.month);
    var prev = budgetOf(shiftMonth(ui.month, -1));
    var hasPrev = SPEND_CATS.some(function (k) { return prev[k]; });
    var limitSum = 0, spentInBudget = 0;
    SPEND_CATS.forEach(function (k) { if (bud[k]) { limitSum += bud[k]; spentInBudget += s[k]; } });

    var cards = SPEND_CATS.map(function (k) {
      var st = budgetStatus(s[k], bud[k]);
      var color = st.level === 'over' ? 'var(--out)' : st.level === 'warn' ? 'var(--warn-bar)' : catOf(k).color;
      var text;
      if (st.level === 'none') text = 'Đã chi ' + money(s[k]) + ' đ · chưa đặt hạn mức';
      else if (st.level === 'over') text = 'Đã chi ' + money(s[k]) + ' đ · vượt ' + money(s[k] - bud[k]) + ' đ';
      else text = 'Đã chi ' + money(s[k]) + ' đ (' + st.pct + '%) · còn ' + money(bud[k] - s[k]) + ' đ';
      return '<div class="card budget">' +
        '<div class="bhead"><span class="tag" style="background:' + catOf(k).color + '">' + esc(k) + '</span>' +
          '<label for="b-' + esc(k) + '" class="bname">' + esc(catOf(k).name) + '</label>' +
          '<input id="b-' + esc(k) + '" class="bin num" inputmode="numeric" autocomplete="off" placeholder="Nhập hạn mức" data-budget="' + esc(k) + '" value="' + (bud[k] ? money(bud[k]) : '') + '"></div>' +
        '<div class="bar" role="img" aria-label="' + esc(text) + '"><span style="width:' + Math.min(100, st.pct) + '%;background:' + color + '"></span></div>' +
        '<div class="bstat' + (st.level === 'over' ? ' over' : st.level === 'warn' ? ' warn' : '') + '">' + text + '</div>' +
      '</div>';
    }).join('');

    return monthHeader('Ngân sách · ' + monthLabel(ui.month)) +
      '<p class="lead">Đặt hạn mức chi cho từng danh mục. App báo khi chi đạt 80% và khi vượt hạn mức. Để trống nếu không cần theo dõi.</p>' +
      (limitSum ? '<div class="mini two"><div>Tổng hạn mức<b class="num">' + money(limitSum) + ' đ</b></div>' +
        '<div>Đã chi (các mục có hạn mức)<b class="num" style="color:' + (spentInBudget > limitSum ? 'var(--out)' : 'var(--fg)') + '">' + money(spentInBudget) + ' đ</b></div></div>' : '') +
      cards +
      (hasPrev ? '<button class="btn ghost" data-action="copy-budget">Chép hạn mức từ ' + monthLabel(shiftMonth(ui.month, -1)).toLowerCase() + '</button>' : '');
  }

  // Bước nhảy trục đẹp: 1, 2, 2.5, 5 x 10^n
  function niceStep(raw) {
    var p = Math.pow(10, Math.floor(Math.log10(raw))), f = raw / p;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p;
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

  function viewReport() {
    var yr = ui.year, items = itemsOfYear(yr), ys = totals(items);
    var months = [];
    for (var i = 1; i <= 12; i++) months.push(totals(items.filter(function (t) { return +t.date.slice(5, 7) === i; })));

    var rows = months.map(function (m, i) {
      var empty = !m.thu && !m.chi;
      return '<tr' + (empty ? ' class="dim"' : '') + '><th scope="row"><button class="link-btn" data-action="open-month" data-m="' + yr + '-' + pad(i + 1) + '">Tháng ' + (i + 1) + '</button></th>' +
        '<td>' + (empty ? '–' : money(m.thu)) + '</td><td>' + (empty ? '–' : money(m.chi)) + '</td><td>' + (empty ? '–' : money(m.con)) + '</td></tr>';
    }).join('');

    var cats = ALL_SPEND.filter(function (k) { return !CATS[k].hidden || ys[k]; }).map(function (k) {
      var pct = ys.chi ? Math.round(ys[k] / ys.chi * 1000) / 10 : 0;
      return '<div class="cat-row"><div class="t"><span>' + esc(k) + ' · ' + esc(catOf(k).name) + '</span><span class="num">' + money(ys[k]) + ' đ <span class="muted">(' + String(pct).replace('.', ',') + '%)</span></span></div>' +
        '<div class="bar"><span style="width:' + pct + '%;background:' + catOf(k).color + '"></span></div></div>';
    }).join('');

    var head = '<header class="top">' +
      '<button class="icon-btn" data-action="year" data-delta="-1" aria-label="Năm trước">' + I.prev + '</button>' +
      '<h1>Báo cáo năm ' + yr + '</h1>' +
      '<button class="icon-btn" data-action="year" data-delta="1" aria-label="Năm sau">' + I.next + '</button></header>';
    var dataLink = '<button class="link-btn center" data-action="go" data-view="settings">Sao lưu / khôi phục dữ liệu: vào Cài đặt</button>';

    if (!items.length) {
      return head + '<section class="card">' + emptyBox('Năm ' + yr + ' chưa có khoản thu chi nào.') + '</section>' + dataLink;
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
      dataLink;
  }

  function viewMoney() {
    var body = ui.moneyTab === 'list' ? viewList() : ui.moneyTab === 'budget' ? viewBudget() : ui.moneyTab === 'report' ? viewReport() : viewOverview();
    return moneyNav() + body;
  }

  // ---------- Cài đặt ----------
  function viewSettings() {
    var used = {};
    S.expenses.forEach(function (t) { used[t.cat] = 1; });
    var catRows = S.categories.map(function (c) {
      var canDel = !c.builtin && !used[c.code] && !S.groups.some(function (g) { return g.cat === c.code; });
      return '<div class="srow' + (c.hidden ? ' off' : '') + '"><span class="tag" style="background:' + c.color + '">' + esc(c.code) + '</span>' +
        '<input class="sin" aria-label="Tên danh mục ' + esc(c.code) + '" maxlength="30" value="' + esc(c.name) + '" data-catname="' + esc(c.code) + '">' +
        (c.code === '+' ? '<span class="muted small">Thu nhập</span>' :
          (canDel ? '<button class="mini-btn" data-action="cat-del" data-code="' + esc(c.code) + '">Xoá</button>'
            : '<button class="mini-btn" data-action="cat-hide" data-code="' + esc(c.code) + '">' + (c.hidden ? 'Hiện' : 'Ẩn') + '</button>')) +
        '</div>';
    }).join('');
    var catOpts = function (sel) {
      return SPEND_CATS.map(function (k) { return '<option value="' + esc(k) + '"' + (sel === k ? ' selected' : '') + '>' + esc(k) + ' · ' + esc(catOf(k).name) + '</option>'; }).join('');
    };
    var groupRows = S.groups.map(function (g) {
      var n = S.shopping.filter(function (i) { return i.group === g.id; }).length;
      return '<div class="srow' + (g.hidden ? ' off' : '') + '">' +
        '<input class="sin" aria-label="Tên nhóm hàng" maxlength="40" value="' + esc(g.name) + '" data-gname="' + esc(g.id) + '">' +
        '<select class="ssel" aria-label="Ghi chi vào danh mục" data-gcat="' + esc(g.id) + '">' + catOpts(g.cat) + '</select>' +
        (g.builtin ? '' : n ? '<button class="mini-btn" data-action="group-hide" data-id="' + esc(g.id) + '">' + (g.hidden ? 'Hiện' : 'Ẩn') + '</button>'
          : '<button class="mini-btn" data-action="group-del" data-id="' + esc(g.id) + '">Xoá</button>') +
        '</div>';
    }).join('');
    var r = ui.restore;
    var confirm = r
      ? '<div class="confirm" role="alert"><span>Thay ' + (r.kind === 'v1' ? 'toàn bộ khoản thu chi và ngân sách hiện tại (' + S.expenses.length + ' khoản) bằng ' + r.data.expenses.length + ' khoản'
          : 'toàn bộ dữ liệu hiện tại bằng bản sao lưu (' + r.data.expenses.length + ' khoản thu chi, ' + r.data.tasks.length + ' việc' + (r.fileCount ? ', ' + r.fileCount + ' tệp' : '') + ')') +
        ' trong file ngày ' + esc(r.savedAt && validDate(r.savedAt) ? r.savedAt.split('-').reverse().join('/') : '?') + '? Không thể hoàn tác.</span>' +
        '<div class="row"><button type="button" class="btn ghost" data-action="cancel-restore">Huỷ</button>' +
        '<button type="button" class="btn danger solid" data-action="do-restore">Khôi phục</button></div></div>'
      : '';
    var mode = Store.mode === 'idb' ? 'Bộ nhớ của trình duyệt trên máy này (IndexedDB).' : Store.mode === 'local' ? 'Bộ nhớ nhỏ của trình duyệt (tệp đính kèm chưa lưu lâu dài).' : 'Chỉ giữ trong phiên này.';

    return topBar('Cài đặt', '<button class="icon-btn" data-action="go" data-view="today" aria-label="Quay lại">' + I.prev + '</button>') +
      (ui.setMsg ? '<p class="note ok" role="status">' + esc(ui.setMsg) + '</p>' : '') +
      '<section class="card"><h2>Danh mục thu chi</h2>' +
        '<p class="lead">Danh mục đã có dữ liệu chỉ ẩn được, không xoá, để sổ và báo cáo cũ không bị lệch. Mã mới cũng dùng ở cột KH khi xuất Excel.</p>' +
        catRows +
        '<form class="add-row" id="catadd"><input id="c-code" maxlength="1" autocomplete="off" aria-label="Mã danh mục mới (1 chữ cái)" placeholder="Mã" class="code-in">' +
          '<input id="c-name" maxlength="30" autocomplete="off" aria-label="Tên danh mục mới" placeholder="Tên, ví dụ: Y tế">' +
          '<button class="btn" type="submit">Thêm</button></form>' +
        '<label class="check-line"><input type="checkbox" id="c-reimb"> Danh mục này được công ty hoàn tiền</label>' +
      '</section>' +
      '<section class="card"><h2>Nhóm hàng mua sắm</h2>' +
        '<p class="lead">Mỗi nhóm có một danh mục chi mặc định. Ví dụ Thực phẩm ghi vào G.</p>' +
        groupRows +
        '<form class="add-row" id="groupadd"><input id="g-name" maxlength="40" autocomplete="off" aria-label="Tên nhóm mới" placeholder="Nhóm mới, ví dụ: Thuốc men">' +
          '<select id="g-cat" class="ssel" aria-label="Danh mục chi của nhóm mới">' + catOpts('K') + '</select>' +
          '<button class="btn" type="submit">Thêm</button></form>' +
      '</section>' +
      '<section class="card"><h2>Tổng kết ngày</h2>' +
        '<label class="srow"><span class="grow">Hiện thẻ Tổng kết ngày từ</span><input type="time" class="sin time" value="' + esc(S.settings.summaryTime || '22:00') + '" data-setting="summaryTime"></label>' +
      '</section>' +
      '<section class="card"><h2>Dữ liệu của anh</h2>' +
        '<p class="lead">Nơi lưu: ' + mode + (ui.storageInfo ? ' ' + esc(ui.storageInfo) : '') + '</p>' +
        '<p class="lead">Dữ liệu chỉ nằm trên máy này. Nên sao lưu mỗi tháng một lần và cất file ở Google Drive hoặc Zalo “Cloud của tôi”.' +
          (S.settings.lastBackup ? ' Lần gần nhất: ' + esc(S.settings.lastBackup.split('-').reverse().join('/')) + '.' : '') + '</p>' +
        '<button class="btn ghost with-icon" data-action="backup" data-files="0">' + I.download + 'Sao lưu dữ liệu (nhẹ, không kèm tệp)</button>' +
        (S.attachments.length ? '<button class="btn ghost with-icon" data-action="backup" data-files="1">' + I.download + 'Sao lưu kèm ' + S.attachments.length + ' tệp đính kèm</button>' : '') +
        '<label class="btn ghost with-icon file-btn" for="restore-file">Khôi phục từ file sao lưu' +
          '<input id="restore-file" type="file" accept=".json,application/json" data-restore="1"></label>' +
        confirm +
        (ui.dataMsg ? '<p class="lead" role="status">' + esc(ui.dataMsg) + '</p>' : '') +
      '</section>' +
      '<p class="lead center">MyLife Hub · giai đoạn 1</p>';
  }

  // ---------- Lớp phủ: bảng hỏi, xem ảnh ----------
  var overlay = document.createElement('div');
  overlay.id = 'overlay';
  document.body.appendChild(overlay);

  function renderOverlay() {
    var v = ui.viewer, sh = ui.sheet;
    if (v) {
      overlay.innerHTML = '<div class="viewer" role="dialog" aria-modal="true" aria-label="' + esc(v.name) + '">' +
        '<img src="' + v.url + '" alt="' + esc(v.name) + '">' +
        '<div class="vbar"><span>' + esc(v.name) + '</span><button class="btn ghost" data-action="viewer-close">Đóng</button></div></div>';
    } else if (sh) {
      overlay.innerHTML = '<div class="scrim" data-action="sheet-close"></div>' +
        '<div class="sheet" role="dialog" aria-modal="true" aria-labelledby="sheet-t">' +
        '<h2 id="sheet-t">' + esc(sh.title) + '</h2>' + (sh.text ? '<p>' + sh.text + '</p>' : '') +
        '<div class="sheet-btns">' + sh.buttons.map(function (b) {
          return '<button class="btn' + (b.cls ? ' ' + b.cls : '') + '" data-action="' + b.action + '"' + (b.id ? ' data-id="' + esc(b.id) + '"' : '') + '>' + esc(b.label) + '</button>';
        }).join('') + '</div></div>';
      var f = overlay.querySelector('.sheet .btn');
      if (f) f.focus();
    } else overlay.innerHTML = '';
    document.body.classList.toggle('modal-open', !!(v || sh));
  }
  function openSheet(sh) { ui.sheet = sh; renderOverlay(); }
  function closeSheet() { ui.sheet = null; renderOverlay(); }

  // ---------- Vẽ lại ----------
  var main = document.getElementById('app'), tabBar = document.getElementById('tabs'), toastEl = document.getElementById('toast');
  var toastTimer = null;
  var VIEWS = { today: viewToday, calendar: viewCalendar, tasks: viewTasks, money: viewMoney, form: viewForm, task: viewTaskForm, shopitem: viewShopItem, settings: viewSettings };
  var FORMS = { form: 1, task: 1, shopitem: 1 };

  function tabs() {
    var cur = FORMS[ui.view] ? ui.back : ui.view === 'settings' ? 'today' : ui.view;
    function t(v, label, icon) {
      return '<button class="tab" data-action="go" data-view="' + v + '"' + (cur === v ? ' aria-current="page"' : '') + '>' + icon + label + '</button>';
    }
    return '<div class="tabs-in">' + t('today', 'Hôm nay', I.today) + t('calendar', 'Lịch', I.cal) +
      '<button class="tab add" data-action="add" aria-label="Ghi nhanh: khoản thu chi, việc, món cần mua">' + I.plus + '</button>' +
      t('tasks', 'Việc', I.tasks) + t('money', 'Tiền', I.wallet) + '</div>';
  }

  var thumbUrls = {};
  function hydrateThumbs() {
    Array.prototype.forEach.call(main.querySelectorAll('img[data-att]'), function (img) {
      var id = img.getAttribute('data-att');
      if (thumbUrls[id]) { img.src = thumbUrls[id]; return; }
      Store.getFile(id + ':t').then(function (b) { return b || Store.getFile(id); }).then(function (b) {
        if (!b) { img.replaceWith(Object.assign(document.createElement('span'), { className: 'nm', textContent: 'Thiếu tệp' })); return; }
        thumbUrls[id] = URL.createObjectURL(b);
        img.src = thumbUrls[id];
      });
    });
  }

  function render(focusId) {
    main.innerHTML = (VIEWS[ui.view] || viewToday)();
    tabBar.innerHTML = tabs();
    tabBar.hidden = !!FORMS[ui.view];
    hydrateThumbs();
    renderOverlay();
    if (focusId) {
      var el = document.getElementById(focusId);
      if (el) { el.focus(); if (el.setSelectionRange && /^(text|search|)$/.test(el.type || '') ) { try { var n = el.value.length; el.setSelectionRange(n, n); } catch (e) {} } }
    }
  }

  function showToast(msg, undoId) {
    toastEl.innerHTML = '<span>' + esc(msg) + '</span>' + (undoId ? '<button class="link-btn" data-action="undo" data-id="' + esc(undoId) + '">Hoàn tác</button>' : '');
    toastEl.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.hidden = true; }, undoId ? 5000 : 2600);
  }

  function go(view) {
    if (view === 'settings' && ui.view !== 'settings') { ui.dataMsg = ''; ui.restore = null; ui.setMsg = ''; loadStorageInfo(); }
    if (view === 'money' && ui.view === 'money' && ui.moneyTab !== 'overview') ui.moneyTab = 'overview';
    ui.view = view; ui.confirmDelete = false; ui.errors = {};
    render(); window.scrollTo(0, 0);
  }
  function goMoney(tab) {
    if (tab === 'report' && ui.moneyTab !== 'report') { ui.year = ui.month.slice(0, 4); ui.exportMsg = ''; }
    if (tab !== 'list' && ui.filter === 'owed') ui.filter = 'all';
    ui.moneyTab = tab; ui.view = 'money'; ui.confirmDelete = false; ui.errors = {};
    render(); window.scrollTo(0, 0);
  }
  function backView() { return FORMS[ui.view] ? ui.back : ui.view; }

  function loadStorageInfo() {
    ui.storageInfo = '';
    Store.estimate().then(function (e) {
      if (!e || !e.quota) return;
      ui.storageInfo = 'Đang dùng ' + fmtSize(e.usage || 0) + ', máy còn cho phép khoảng ' + fmtSize(Math.max(0, e.quota - (e.usage || 0))) + '.';
      if (ui.view === 'settings') render();
    });
  }

  // ---------- Ghi nhanh khoản thu chi ----------
  function defaultCat() {
    var since = addDays(todayISO(), -30), count = {};
    S.expenses.forEach(function (t) { if (t.type === 'chi' && t.date >= since && SPEND_CATS.indexOf(t.cat) >= 0) count[t.cat] = (count[t.cat] || 0) + 1; });
    var best = '', n = 0;
    Object.keys(count).forEach(function (k) { if (count[k] > n) { best = k; n = count[k]; } });
    return best;
  }

  function openExpense(item, preset) {
    ui.back = backView();
    if (item) {
      ui.draft = { id: item.id, isNew: false, cat: item.cat, amount: String(item.amount), amountText: '', date: item.date, content: item.content, note: item.note || '',
        reimbursed: item.reimburse === 'da_hoan', source: item.source || 'manual', sourceIds: item.sourceIds || [] };
    } else {
      var today = todayISO();
      ui.draft = Object.assign({ id: uid(), isNew: true, cat: defaultCat(), amount: '', amountText: '', date: today, content: '', note: '', reimbursed: false, source: 'manual', sourceIds: [] }, preset || {});
    }
    ui.errors = {}; ui.confirmDelete = false; ui.view = 'form';
    render(item ? null : 'f-amount'); window.scrollTo(0, 0);
  }

  function closeExpense() {
    var d = ui.draft;
    if (d && d.isNew) removeAttachments('expense', d.id);
    ui.draft = null;
    go(ui.back);
  }

  function applyAmountOp(op) {
    var d = ui.draft, v;
    if (op === '000') { if (!d.amountText) d.amount = String(Math.min(MAX_AMOUNT, +(d.amount || '0') * 1000 || 0) || ''); }
    else {
      var mul = op === 'k' ? 1e3 : 1e6;
      v = d.amountText ? parseFloat(d.amountText.replace(',', '.')) : +d.amount;
      if (v) { d.amount = String(Math.round(v * mul)); d.amountText = ''; }
    }
    if (+d.amount > MAX_AMOUNT) d.amount = String(MAX_AMOUNT);
    delete ui.errors.amount;
    render();
  }

  function saveExpense() {
    var d = ui.draft, e = {}, amt = +d.amount;
    if (d.amountText) e.amount = 'Bấm × nghìn hoặc × triệu để đổi số có dấu phẩy thành số tiền.';
    else if (!amt) e.amount = 'Nhập số tiền lớn hơn 0.';
    else if (amt > MAX_AMOUNT) e.amount = 'Số tiền quá lớn, kiểm tra lại số chữ số.';
    if (!CATS[d.cat]) e.cat = 'Chọn một danh mục.';
    if (!validDate(d.date)) e.date = 'Chọn ngày hợp lệ.';
    ui.errors = e;
    if (Object.keys(e).length) { render(e.amount ? 'f-amount' : e.date ? 'f-date' : null); return; }

    var old = d.isNew ? null : byId(S.expenses, d.id), c = CATS[d.cat];
    var rec = Object.assign({}, old || { id: d.id, createdAt: Date.now(), source: d.source, sourceIds: d.sourceIds.slice() }, {
      date: d.date, type: d.cat === '+' ? 'thu' : 'chi', cat: d.cat, amount: amt,
      content: (d.content || '').trim() || c.name, note: (d.note || '').trim(),
      reimburse: c.reimbursable ? (d.reimbursed ? 'da_hoan' : 'cho') : null
    });
    rec.reimbursedOn = rec.reimburse === 'da_hoan' ? (old && old.reimbursedOn || todayISO()) : null;
    var ops = [upsert('expenses', rec)];
    if (d.isNew) ops = ops.concat(linkSources(rec));
    persist(ops);
    ui.month = monthOf(rec.date);
    var msg = d.isNew ? 'Đã ghi ' + money(rec.amount) + ' đ · ' + rec.cat : 'Đã lưu thay đổi';
    if (rec.type === 'chi') {
      var w = budgetLine(rec.cat, totals(itemsOfMonth(ui.month))[rec.cat], budgetOf(ui.month)[rec.cat]);
      if (w) msg += '. Chú ý: ' + w.charAt(0).toLowerCase() + w.slice(1);
    }
    ui.draft = null;
    go(ui.back);
    showToast(msg, d.isNew ? rec.id : null);
  }

  // Khoản chi vừa tạo từ Mua sắm / Việc: gắn hai chiều để không ghi trùng
  function linkSources(rec) {
    var ops = [];
    if (rec.source === 'shopping') {
      var priceChanged = false;
      rec.sourceIds.forEach(function (id) {
        var it = byId(S.shopping, id);
        if (!it) return;
        var n = Object.assign({}, it, { expenseId: rec.id });
        if (rec.sourceIds.length === 1) {
          n.price = Math.round(rec.amount / (it.qty || 1));
          S.prices[fold(it.name)] = n.price; priceChanged = true;
        }
        ops.push(upsert('shopping', n));
      });
      if (priceChanged) ops.push(metaOp('prices'));
    } else if (rec.source === 'task') {
      var t = byId(S.tasks, rec.sourceIds[0]);
      if (t) ops.push(upsert('tasks', Object.assign({}, t, { expenseId: rec.id, prompt: 'logged' })));
    }
    return ops;
  }

  function deleteExpense(id) {
    var rec = byId(S.expenses, id);
    if (!rec) return;
    var ops = [drop('expenses', id)];
    S.shopping.forEach(function (it) { if (it.expenseId === id) ops.push(upsert('shopping', Object.assign({}, it, { expenseId: null }))); });
    S.tasks.forEach(function (t) { if (t.expenseId === id) ops.push(upsert('tasks', Object.assign({}, t, { expenseId: null, prompt: 'pending' }))); });
    persist(ops);
    removeAttachments('expense', id);
  }

  // ---------- Việc ----------
  function openTask(t, preset) {
    ui.back = backView();
    if (t) ui.tdraft = { id: t.id, isNew: false, title: t.title, tag: t.tag, priority: t.priority, due: t.due || '', time: t.time || '', repeat: t.repeat || '', notes: t.notes || '', cost: t.cost ? String(t.cost) : '', costCat: t.costCat || TAG_CAT[t.tag] };
    else ui.tdraft = Object.assign({ id: uid(), isNew: true, title: '', tag: 'work', priority: 2, due: todayISO(), time: '', repeat: '', notes: '', cost: '', costCat: 'T' }, preset || {});
    ui.errors = {}; ui.confirmDelete = false; ui.view = 'task';
    render(t ? null : 't-title'); window.scrollTo(0, 0);
  }
  function closeTask() {
    var d = ui.tdraft;
    if (d && d.isNew) removeAttachments('task', d.id);
    ui.tdraft = null;
    go(ui.back);
  }
  function saveTask() {
    var d = ui.tdraft, e = {};
    if (!d.title.trim()) e.title = 'Ghi tên việc.';
    if (d.due && !validDate(d.due)) e.due = 'Chọn ngày hợp lệ.';
    if (d.repeat && !d.due) e.due = 'Việc lặp lại cần có ngày hạn đầu tiên.';
    if (+d.cost > MAX_AMOUNT) e.cost = 'Số tiền quá lớn.';
    ui.errors = e;
    if (Object.keys(e).length) { render(e.title ? 't-title' : 't-due'); return; }
    var old = d.isNew ? null : byId(S.tasks, d.id);
    var due = d.due;
    if (d.repeat === 'monthEnd' && due) { var dd = parseISO(due); due = isoOf(new Date(dd.getFullYear(), dd.getMonth() + 1, 0)); }
    var rec = Object.assign({}, old || { id: d.id, status: 'open', doneAt: null, expenseId: null, prompt: null, createdAt: Date.now() }, {
      title: d.title.trim(), tag: d.tag, priority: +d.priority, due: due || '', time: due ? d.time || '' : '',
      repeat: d.repeat, repeatDay: d.repeat === 'month' && due ? +due.slice(8, 10) : null,
      notes: d.notes.trim(), cost: +d.cost || 0, costCat: +d.cost ? d.costCat : null
    });
    persist([upsert('tasks', rec)]);
    ui.tdraft = null;
    go(ui.back);
    showToast(d.isNew ? 'Đã thêm việc' : 'Đã lưu thay đổi');
  }

  function toggleTask(id) {
    var t = byId(S.tasks, id);
    if (!t) return;
    if (!isOpen(t)) {
      persist([upsert('tasks', Object.assign({}, t, { status: 'open', doneAt: null }))]);
      render(); return;
    }
    var now = Date.now(), done, ops = [];
    if (t.repeat && t.due) {
      // Việc lặp lại: lưu một bản "đã xong" cho lần này, việc gốc chuyển sang lần tới (giữ tệp đính kèm)
      done = Object.assign({}, t, { id: uid(), status: 'done', doneAt: now, repeat: '', repeatDay: null, parentId: t.id, prompt: t.cost ? 'pending' : null, expenseId: null });
      ops.push(upsert('tasks', done));
      ops.push(upsert('tasks', Object.assign({}, t, { due: nextDue(t) })));
    } else {
      done = Object.assign({}, t, { status: 'done', doneAt: now, prompt: t.cost && !t.expenseId ? 'pending' : t.prompt });
      ops.push(upsert('tasks', done));
    }
    persist(ops);
    if (ui.view === 'task') { ui.tdraft = null; ui.view = ui.back; }
    render();
    if (done.cost && !done.expenseId) {
      openSheet({ title: 'Ghi chi phí?', text: 'Việc “' + esc(done.title) + '” có chi phí dự kiến <b>' + money(done.cost) + ' đ</b>. Ghi thành khoản chi?',
        buttons: [{ label: 'Ghi chi', action: 'task-log', id: done.id }, { label: 'Không ghi', action: 'task-skip', id: done.id, cls: 'ghost' }] });
    } else showToast('Đã xong: ' + cut(done.title, 40) + (t.repeat ? '. Lần tới: ' + dueText({ due: nextDue(t) }) : ''));
  }

  function deleteTask(id) {
    persist([drop('tasks', id)]);
    removeAttachments('task', id);
  }

  // ---------- Mua sắm ----------
  function addShopping(text) {
    var parts = text.split(/[,;\n]+/).map(function (s) { return s.trim(); }).filter(Boolean);
    var ops = [], added = 0, dup = 0;
    parts.forEach(function (p) {
      var m = p.match(/^(\d{1,3})\s+(.+)$/), qty = 1, name = p;
      if (m) { qty = +m[1] || 1; name = m[2]; }
      name = name.charAt(0).toUpperCase() + name.slice(1);
      if (S.shopping.some(function (i) { return i.status === 'need' && fold(i.name) === fold(name); })) { dup++; return; }
      ops.push(upsert('shopping', { id: uid(), name: cut(name, 80), qty: qty, price: S.prices[fold(name)] || 0, group: groupFor(name), status: 'need', boughtOn: null, choice: null, expenseId: null, createdAt: Date.now() }));
      added++;
    });
    if (ops.length) persist(ops);
    render('shop-new');
    if (added || dup) showToast((added ? 'Đã thêm ' + added + ' món' : '') + (dup ? (added ? '; ' : '') + dup + ' món đã có trong danh sách' : ''));
  }

  function shopToggle(id) {
    var it = byId(S.shopping, id);
    if (!it) return;
    if (it.status === 'need') {
      var n = Object.assign({}, it, { status: 'bought', boughtOn: todayISO(), choice: null });
      persist([upsert('shopping', n)]);
      render();
      var all = { single: { label: 'Ghi riêng món này', action: 'shop-single', id: id },
        batch: { label: 'Gộp lần đi chợ', action: 'shop-batch', id: id },
        skip: { label: 'Không ghi', action: 'shop-skip', id: id } };
      var order = [S.settings.shopChoice || 'single'].concat(['single', 'batch', 'skip'].filter(function (k) { return k !== S.settings.shopChoice; }));
      openSheet({ title: 'Ghi chi tiêu?', text: '“' + esc(it.name) + '”' + (estOf(it) ? ' · ước tính ' + money(estOf(it)) + ' đ' : ''),
        buttons: order.map(function (k, i) { var b = all[k]; b.cls = i ? 'ghost' : ''; return b; }) });
      return;
    }
    var ex = it.expenseId ? byId(S.expenses, it.expenseId) : null;
    if (ex && ex.sourceIds.length <= 1) {
      openSheet({ title: 'Bỏ đánh dấu đã mua', text: 'Món này đã ghi thành khoản chi <b>' + money(ex.amount) + ' đ</b>. Xoá luôn khoản chi đó?',
        buttons: [{ label: 'Xoá khoản chi', action: 'shop-untick-del', id: id, cls: 'danger' }, { label: 'Giữ khoản chi', action: 'shop-untick-keep', id: id, cls: 'ghost' }] });
      return;
    }
    untickShop(it, false);
  }
  function untickShop(it, delExpense) {
    var ops = [], ex = it.expenseId ? byId(S.expenses, it.expenseId) : null;
    if (ex) {
      if (delExpense) { deleteExpense(ex.id); it = byId(S.shopping, it.id); }
      else if (ex.sourceIds.length > 1) ops.push(upsert('expenses', Object.assign({}, ex, { sourceIds: ex.sourceIds.filter(function (x) { return x !== it.id; }) })));
    }
    ops.push(upsert('shopping', Object.assign({}, it, { status: 'need', boughtOn: null, choice: null, expenseId: null })));
    persist(ops);
    render();
  }
  function setShopChoice(id, choice) {
    var it = byId(S.shopping, id);
    S.settings.shopChoice = choice;
    var ops = [metaOp('settings')];
    if (it) ops.push(upsert('shopping', Object.assign({}, it, { choice: choice })));
    persist(ops);
    return it;
  }
  function logShopping(ids) {
    var items = ids.map(function (id) { return byId(S.shopping, id); }).filter(Boolean);
    if (!items.length) return;
    var est = items.reduce(function (a, it) { return a + estOf(it); }, 0);
    // Danh mục: nhóm chiếm nhiều tiền nhất (chưa có giá thì nhóm nhiều món nhất)
    var weight = {};
    items.forEach(function (it) { var c = groupOf(it.group).cat; weight[c] = (weight[c] || 0) + (estOf(it) || 0) + 0.001; });
    var cat = Object.keys(weight).sort(function (a, b) { return weight[b] - weight[a]; })[0];
    if (!CATS[cat] || CATS[cat].hidden) cat = 'K';
    var date = items.map(function (it) { return it.boughtOn || todayISO(); }).sort().pop();
    var content = items.length === 1 ? items[0].name : 'Đi chợ: ' + items.map(function (it) { return it.name.toLowerCase(); }).join(', ');
    ui.back = backView();
    openExpense(null, { cat: cat, amount: est ? String(est) : '', date: date, content: cut(content, 120), source: 'shopping', sourceIds: items.map(function (it) { return it.id; }) });
  }

  function openShopItem(it) {
    ui.back = backView();
    ui.sdraft = { id: it.id, name: it.name, qty: String(it.qty || 1), price: it.price ? String(it.price) : '', group: it.group, newName: '', newCat: 'K' };
    ui.errors = {}; ui.confirmDelete = false; ui.view = 'shopitem';
    render(); window.scrollTo(0, 0);
  }
  function saveShopItem() {
    var d = ui.sdraft, e = {};
    if (!d.name.trim()) e.name = 'Ghi tên món.';
    if (d.group === '__new' && !d.newName.trim()) e.newName = 'Ghi tên nhóm mới.';
    ui.errors = e;
    if (Object.keys(e).length) { render(e.name ? 's-name' : 's-gname'); return; }
    var it = byId(S.shopping, d.id), ops = [], gid = d.group;
    if (gid === '__new') {
      var g = { id: uid(), name: d.newName.trim(), cat: d.newCat, keywords: [], sort: S.groups.length, hidden: false, builtin: false };
      ops.push(upsert('groups', g)); gid = g.id;
    }
    var name = d.name.trim(), f = fold(name);
    if (gid !== it.group) {
      // Học cách xếp nhóm: lần sau món cùng tên vào nhóm này
      S.groups.forEach(function (g) {
        var has = (g.keywords || []).indexOf(f) >= 0;
        if (g.id === gid && !has) ops.push(upsert('groups', Object.assign({}, g, { keywords: (g.keywords || []).concat(f) })));
        if (g.id !== gid && has) ops.push(upsert('groups', Object.assign({}, g, { keywords: g.keywords.filter(function (k) { return k !== f; }) })));
      });
    }
    var price = +d.price || 0;
    if (price) { S.prices[f] = price; ops.push(metaOp('prices')); }
    ops.push(upsert('shopping', Object.assign({}, it, { name: name, qty: Math.max(1, Math.min(999, +d.qty || 1)), price: price, group: gid })));
    persist(ops);
    ui.sdraft = null;
    go(ui.back);
    showToast('Đã lưu');
  }

  // ---------- Tệp đính kèm ----------
  function compressImage(file, maxSide, quality) {
    return new Promise(function (ok) {
      var url = URL.createObjectURL(file), img = new Image();
      img.onload = function () {
        var w = img.naturalWidth, h = img.naturalHeight, r = Math.min(1, maxSide / Math.max(w, h));
        var c = document.createElement('canvas');
        c.width = Math.max(1, Math.round(w * r)); c.height = Math.max(1, Math.round(h * r));
        c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        if (!c.toBlob) { ok(null); return; }
        c.toBlob(function (b) { ok(b); }, 'image/jpeg', quality);
      };
      img.onerror = function () { URL.revokeObjectURL(url); ok(null); };
      img.src = url;
    });
  }

  var pendingFiles = null;
  function handleAttach(input) {
    var files = Array.prototype.slice.call(input.files || []);
    var type = input.getAttribute('data-owner-type'), owner = input.getAttribute('data-owner-id');
    input.value = '';
    if (!files.length) return;
    var big = files.filter(function (f) { return f.size > BIG_FILE; });
    if (big.length) {
      pendingFiles = { files: files, type: type, owner: owner };
      openSheet({ title: 'Tệp lớn', text: big.map(function (f) { return '“' + esc(f.name) + '” ' + fmtSize(f.size); }).join(', ') + '. Tệp lớn chiếm nhiều bộ nhớ máy và làm file sao lưu nặng. Vẫn lưu?',
        buttons: [{ label: 'Vẫn lưu', action: 'big-yes' }, { label: 'Bỏ tệp lớn', action: 'big-no', cls: 'ghost' }] });
      return;
    }
    saveFiles(files, type, owner);
  }

  function saveFiles(files, type, owner) {
    ui.attBusy += files.length; render();
    files.reduce(function (p, f) {
      return p.then(function () {
        var isImg = /^image\//.test(f.type) && !/gif|svg/.test(f.type);
        var main0 = isImg ? compressImage(f, 1600, 0.82) : Promise.resolve(null);
        return main0.then(function (small) {
          var body = small && small.size < f.size ? small : f;
          var name = body === f ? f.name : f.name.replace(/\.[^.]+$/, '') + '.jpg';
          var thumb = isImg ? compressImage(body, 320, 0.7) : Promise.resolve(null);
          return thumb.then(function (t) {
            var a = { id: uid(), ownerType: type, ownerId: owner, name: name || 'tep', mime: body.type || f.type || 'application/octet-stream', size: body.size, createdAt: Date.now() };
            var map = {}; map[a.id] = body; if (t) map[a.id + ':t'] = t;
            return Store.putFiles(map).then(function () { S.attachments.push(a); return persist([{ store: 'attachments', put: a }]); });
          });
        }).catch(function () { showToast('Không lưu được “' + f.name + '”. Máy có thể đã đầy bộ nhớ.'); })
          .then(function () { ui.attBusy--; });
      });
    }, Promise.resolve()).then(function () { ui.attBusy = 0; render(); });
  }

  function removeAttachments(type, owner) {
    var list = attsOf(type, owner);
    if (!list.length) return;
    var ops = [];
    list.forEach(function (a) {
      ops.push(drop('attachments', a.id), { store: 'files', del: a.id }, { store: 'files', del: a.id + ':t' });
      if (thumbUrls[a.id]) { URL.revokeObjectURL(thumbUrls[a.id]); delete thumbUrls[a.id]; }
      delete Store.memFiles[a.id]; delete Store.memFiles[a.id + ':t'];
    });
    persist(ops);
  }

  function openAttachment(id) {
    var a = byId(S.attachments, id);
    if (!a) return;
    Store.getFile(id).then(function (b) {
      if (!b) { showToast('Không tìm thấy nội dung tệp này trên máy.'); return; }
      var url = URL.createObjectURL(b);
      if (/^image\//.test(a.mime)) { ui.viewer = { url: url, name: a.name }; renderOverlay(); return; }
      var link = document.createElement('a');
      link.href = url; link.target = '_blank'; link.rel = 'noopener';
      if (!/pdf/.test(a.mime)) link.download = a.name;
      document.body.appendChild(link); link.click(); link.remove();
      setTimeout(function () { URL.revokeObjectURL(url); }, 60000);
    });
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
    var items = itemsOfYear(yr), wb = X.utils.book_new(), ys = totals(items);
    var cols = ALL_SPEND.filter(function (k) { return !CATS[k].hidden || ys[k]; });
    var money0 = '#,##0';
    var catCols = cols.map(function (k) { return k + ' - ' + catOf(k).name; });
    var rep = [['BÁO CÁO THU CHI NĂM ' + yr], [], ['Tháng', 'Tổng thu', 'Tổng chi', 'Còn lại'].concat(catCols)];
    function rowOf(label, m) { return [label, m.thu, m.chi, m.con].concat(cols.map(function (k) { return m[k] || 0; })); }
    for (var i = 1; i <= 12; i++) {
      rep.push(rowOf('Tháng ' + i, totals(items.filter(function (t) { return +t.date.slice(5, 7) === i; }))));
    }
    rep.push(rowOf('Cả năm', ys));
    var ws = X.utils.aoa_to_sheet(rep);
    var numCols = [];
    for (i = 1; i < 4 + cols.length; i++) numCols.push(X.utils.encode_col(i));
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
      if (ui.view === 'money' && ui.moneyTab === 'report') render();
    });
  }

  // ---------- Sao lưu / khôi phục ----------
  function blobToDataURL(b) {
    return new Promise(function (ok, fail) { var fr = new FileReader(); fr.onload = function () { ok(fr.result); }; fr.onerror = fail; fr.readAsDataURL(b); });
  }
  function backup(withFiles) {
    var name = 'MyLifeHub_saoluu_' + todayISO() + (withFiles ? '_kemtep' : '') + '.json';
    var data = { expenses: S.expenses, budgets: S.budgets, categories: S.categories, tasks: S.tasks, shopping: S.shopping, groups: S.groups, attachments: S.attachments, settings: S.settings, prices: S.prices };
    ui.dataMsg = withFiles ? 'Đang gom tệp đính kèm…' : ''; render();
    var files = {};
    var collect = withFiles ? S.attachments.reduce(function (p, a) {
      return p.then(function () {
        return Promise.all([Store.getFile(a.id), Store.getFile(a.id + ':t')]).then(function (bs) {
          var jobs = [];
          if (bs[0]) jobs.push(blobToDataURL(bs[0]).then(function (u) { files[a.id] = u; }));
          if (bs[1]) jobs.push(blobToDataURL(bs[1]).then(function (u) { files[a.id + ':t'] = u; }));
          return Promise.all(jobs);
        });
      });
    }, Promise.resolve()) : Promise.resolve();
    collect.then(function () {
      var payload = JSON.stringify({ app: 'mylife-hub', version: 2, savedAt: todayISO(), data: data, files: withFiles ? files : null });
      return saveFile(name, payload, 'application/json');
    }).then(function () {
      S.settings.lastBackup = todayISO(); persist([metaOp('settings')]);
      ui.dataMsg = 'Đã tạo file ' + name + ' (' + S.expenses.length + ' khoản thu chi, ' + S.tasks.length + ' việc' + (withFiles ? ', ' + Object.keys(files).filter(function (k) { return !/:t$/.test(k); }).length + ' tệp' : '') + ').';
    }, function (err) {
      ui.dataMsg = err && err.code === 'declined' ? 'Anh đã huỷ lưu file.' : 'Không lưu được file ở chế độ xem này.';
    }).then(function () { if (ui.view === 'settings') render(); });
  }

  function validItem(t, cats) {
    return t && typeof t.id === 'string' && validDate(t.date) && (t.type === 'thu' || t.type === 'chi') &&
      cats[t.cat] && (t.type === 'thu') === (t.cat === '+') && typeof t.content === 'string' &&
      typeof t.amount === 'number' && t.amount > 0 && t.amount <= MAX_AMOUNT;
  }

  function readRestoreFile(file) {
    ui.restore = null; ui.dataMsg = '';
    if (!file) return;
    var fr = new FileReader();
    fr.onload = function () {
      try {
        var raw = JSON.parse(fr.result), cats = {};
        if (raw && raw.app === 'so-chi-tieu' && Array.isArray(raw.items)) {
          Store.defaults.categories.forEach(function (c) { cats[c.code] = 1; });
          if (raw.items.some(function (t) { return !validItem(t, cats); })) throw new Error('items');
          ui.restore = { kind: 'v1', data: { expenses: raw.items.map(Store.fromLegacyItem), budgets: raw.budgets && typeof raw.budgets === 'object' ? raw.budgets : {} }, savedAt: raw.savedAt };
        } else if (raw && raw.app === 'mylife-hub' && raw.version === 2 && raw.data && Array.isArray(raw.data.expenses)) {
          var d = raw.data;
          ['categories', 'tasks', 'shopping', 'groups', 'attachments'].forEach(function (k) { if (!Array.isArray(d[k])) d[k] = []; });
          (d.categories.length ? d.categories : Store.defaults.categories).forEach(function (c) { cats[c.code] = 1; });
          if (d.expenses.some(function (t) { return !validItem(t, cats); })) throw new Error('items');
          ui.restore = { kind: 'v2', data: d, files: raw.files || null, savedAt: raw.savedAt, fileCount: raw.files ? Object.keys(raw.files).filter(function (k) { return !/:t$/.test(k); }).length : 0 };
        } else throw new Error('format');
      } catch (e) {
        ui.dataMsg = 'File này không phải file sao lưu của app, hoặc đã bị sửa hỏng. Chọn lại file .json do app tạo ra.';
      }
      render();
    };
    fr.onerror = function () { ui.dataMsg = 'Không đọc được file.'; render(); };
    fr.readAsText(file);
  }

  function doRestore() {
    var r = ui.restore, next;
    ui.restore = null;
    if (r.kind === 'v1') {
      next = Object.assign({}, S, { expenses: r.data.expenses, budgets: r.data.budgets });
      finishRestore(next, null, true);
      return;
    }
    var d = r.data;
    next = { expenses: d.expenses, budgets: d.budgets || {}, categories: d.categories, tasks: d.tasks, shopping: d.shopping, groups: d.groups, attachments: d.attachments, settings: d.settings || {}, prices: d.prices || {} };
    if (!r.files) { finishRestore(next, null, true); return; }
    var files = {};
    Promise.all(Object.keys(r.files).map(function (k) {
      return fetch(r.files[k]).then(function (res) { return res.blob(); }).then(function (b) { files[k] = b; });
    })).then(function () { finishRestore(next, files, false); }, function () { ui.dataMsg = 'Không đọc được tệp đính kèm trong file sao lưu.'; render(); });
  }
  function finishRestore(next, files, keepFiles) {
    var keep = keepFiles && Store.idb ? Store.getAllFiles() : Promise.resolve(Store.idb ? null : Store.memFiles);
    keep.then(function (old) {
      return Store.replaceAll(next, files || old || null);
    }).then(function (s) {
      S = s; refreshCats();
      ui.month = latestMonth(); ui.year = ui.month.slice(0, 4);
      ui.dataMsg = 'Đã khôi phục ' + S.expenses.length + ' khoản thu chi, ' + S.tasks.length + ' việc.';
      render();
    }, function () { ui.dataMsg = 'Khôi phục không thành công, dữ liệu cũ vẫn giữ nguyên.'; render(); });
  }

  // ---------- Danh mục, nhóm hàng ----------
  function addCategory() {
    var code = (document.getElementById('c-code').value || '').trim().toUpperCase();
    var name = (document.getElementById('c-name').value || '').trim();
    var reimb = document.getElementById('c-reimb').checked;
    if (!/^[A-Z]$/.test(code)) { ui.setMsg = 'Mã danh mục là 1 chữ cái (A đến Z).'; render('c-code'); return; }
    if (CATS[code]) { ui.setMsg = 'Mã ' + code + ' đã dùng cho “' + CATS[code].name + '”. Chọn chữ khác.'; render('c-code'); return; }
    if (!name) { ui.setMsg = 'Ghi tên danh mục.'; render('c-name'); return; }
    var extra = S.categories.filter(function (c) { return !c.builtin; }).length;
    var c = { code: code, name: name, short: cut(name, 10), kind: 'chi', color: EXTRA_COLORS[extra % EXTRA_COLORS.length], sort: S.categories.length, hidden: false, builtin: false, reimbursable: reimb };
    S.categories.push(c);
    persist([{ store: 'categories', put: c }]);
    refreshCats();
    ui.setMsg = 'Đã thêm danh mục ' + code + ' · ' + name + '.';
    render();
  }
  function addGroup() {
    var name = (document.getElementById('g-name').value || '').trim(), cat = document.getElementById('g-cat').value;
    if (!name) { ui.setMsg = 'Ghi tên nhóm hàng.'; render('g-name'); return; }
    persist([upsert('groups', { id: uid(), name: name, cat: cat, keywords: [], sort: S.groups.length, hidden: false, builtin: false })]);
    ui.setMsg = 'Đã thêm nhóm ' + name + '.';
    render();
  }

  // ---------- Sự kiện ----------
  document.addEventListener('click', function (ev) {
    var b = ev.target.closest('[data-action]');
    if (!b || !S) return;
    var a = b.getAttribute('data-action'), id = b.getAttribute('data-id');
    var it, t;
    switch (a) {
      case 'month': ui.month = shiftMonth(ui.month, +b.getAttribute('data-delta')); render(); break;
      case 'cal-month':
        ui.calMonth = shiftMonth(ui.calMonth, +b.getAttribute('data-delta'));
        ui.calDay = monthOf(todayISO()) === ui.calMonth ? todayISO() : ui.calMonth + '-01';
        render(); break;
      case 'cal-day': ui.calDay = b.getAttribute('data-d'); render(); break;
      case 'go': go(b.getAttribute('data-view')); break;
      case 'money-tab': goMoney(b.getAttribute('data-tab')); break;
      case 'show-owed': ui.filter = 'owed'; ui.query = ''; goMoney('list'); break;
      case 'year': ui.year = String(+ui.year + +b.getAttribute('data-delta')); ui.exportMsg = ''; render(); break;
      case 'open-month': ui.month = b.getAttribute('data-m'); goMoney('overview'); break;
      case 'export': if (!ui.exporting) exportYear(); break;
      case 'backup': backup(b.getAttribute('data-files') === '1'); break;
      case 'cancel-restore': ui.restore = null; render(); break;
      case 'do-restore': doRestore(); break;
      case 'copy-budget':
        S.budgets[ui.month] = Object.assign({}, budgetOf(shiftMonth(ui.month, -1)));
        persist([{ store: 'budgets', month: ui.month }]); render(); showToast('Đã chép hạn mức tháng trước'); break;

      case 'add': {
        var date = b.getAttribute('data-date');
        if (ui.view === 'form' || ui.view === 'task') break;
        openExpense(null, date ? { date: date } : null); break;
      }
      case 'add-kind': {
        var k = b.getAttribute('data-k'), cur = ui.view === 'task' ? 'task' : 'expense';
        if (k === cur) break;
        var keepBack = ui.back;
        if (ui.view === 'form') { if (ui.draft && ui.draft.isNew) removeAttachments('expense', ui.draft.id); ui.draft = null; }
        if (ui.view === 'task') { if (ui.tdraft && ui.tdraft.isNew) removeAttachments('task', ui.tdraft.id); ui.tdraft = null; }
        ui.view = keepBack;
        if (k === 'task') openTask(null);
        else if (k === 'expense') openExpense(null);
        else { ui.taskFilter = 'shop'; go('tasks'); var inp = document.getElementById('shop-new'); if (inp) inp.focus(); }
        ui.back = keepBack;
        break;
      }
      case 'edit': it = byId(S.expenses, id); if (it) openExpense(it); break;
      case 'close-form':
        if (ui.view === 'form') closeExpense();
        else if (ui.view === 'task') closeTask();
        else { ui.sdraft = null; go(ui.back); }
        break;
      case 'cat': ui.draft.cat = b.getAttribute('data-cat'); delete ui.errors.cat; render(); break;
      case 'amt': applyAmountOp(b.getAttribute('data-op')); break;
      case 'f-date': ui.draft.date = b.getAttribute('data-v'); delete ui.errors.date; render(); break;
      case 'pick-sugg': {
        var s = suggList[+b.getAttribute('data-i')];
        if (!s) break;
        ui.draft.content = s.content;
        if (!ui.draft.amount && !ui.draft.amountText) ui.draft.amount = String(s.amount);
        if (CATS[s.cat] && !CATS[s.cat].hidden) ui.draft.cat = s.cat;
        render(); break;
      }
      case 'filter': ui.filter = b.getAttribute('data-f'); render(); break;
      case 'clear-filter': ui.filter = 'all'; ui.query = ''; render(); break;
      case 'ask-delete': ui.confirmDelete = true; render(); break;
      case 'cancel-delete': ui.confirmDelete = false; render(); break;
      case 'do-delete': {
        var name = ui.draft.content;
        deleteExpense(ui.draft.id); ui.draft = null;
        go(ui.back); showToast('Đã xoá “' + cut(name, 40) + '”'); break;
      }
      case 'undo': {
        var ex = byId(S.expenses, id);
        if (ex) { deleteExpense(id); render(); }
        toastEl.hidden = true; showToast('Đã hoàn tác'); break;
      }
      case 'ask-reimburse-all': {
        var owed = owedList();
        openSheet({ title: 'Đã được hoàn tiền?', text: 'Đánh dấu ' + owed.length + ' khoản, tổng <b>' + money(owed.reduce(function (x, t) { return x + t.amount; }, 0)) + ' đ</b>, là công ty đã hoàn tiền?',
          buttons: [{ label: 'Đánh dấu đã hoàn', action: 'reimburse-all' }, { label: 'Huỷ', action: 'sheet-close', cls: 'ghost' }] });
        break;
      }
      case 'reimburse-all': {
        var today = todayISO();
        persist(owedList().map(function (t) { return upsert('expenses', Object.assign({}, t, { reimburse: 'da_hoan', reimbursedOn: today })); }));
        closeSheet(); render(); showToast('Đã đánh dấu công ty hoàn tiền'); break;
      }

      // Việc
      case 'task-new': openTask(null, b.getAttribute('data-due') ? { due: b.getAttribute('data-due') } : null); break;
      case 'task-edit': t = byId(S.tasks, id); if (t) openTask(t); break;
      case 'task-toggle': toggleTask(id); break;
      case 'task-later': {
        t = byId(S.tasks, id);
        if (t) { persist([upsert('tasks', Object.assign({}, t, { due: addDays(todayISO(), 1) }))]); render(); showToast('Đã dời “' + cut(t.title, 30) + '” sang ngày mai'); }
        break;
      }
      case 'task-filter': ui.taskFilter = b.getAttribute('data-f'); render(); break;
      case 'toggle-done': ui.showDone = !ui.showDone; render(); break;
      case 't-set': {
        var f = b.getAttribute('data-f'), v = b.getAttribute('data-v');
        ui.tdraft[f] = f === 'priority' ? +v : v;
        if (f === 'tag' && !ui.tdraft.cost) ui.tdraft.costCat = TAG_CAT[v];
        if (f === 'due') delete ui.errors.due;
        render(); break;
      }
      case 'task-delete': {
        var title = ui.tdraft.title;
        deleteTask(ui.tdraft.id); ui.tdraft = null;
        go(ui.back); showToast('Đã xoá việc “' + cut(title, 30) + '”'); break;
      }
      case 'task-log': {
        t = byId(S.tasks, id); closeSheet();
        if (t) openExpense(null, { cat: t.costCat && CATS[t.costCat] ? t.costCat : TAG_CAT[t.tag], amount: String(t.cost || ''), content: t.title, date: t.doneAt ? isoOf(new Date(t.doneAt)) : todayISO(), source: 'task', sourceIds: [t.id] });
        break;
      }
      case 'task-skip': {
        t = byId(S.tasks, id); closeSheet();
        if (t) { persist([upsert('tasks', Object.assign({}, t, { prompt: 'skipped' }))]); render(); }
        break;
      }

      // Mua sắm
      case 'shop-focus': { var el = document.getElementById('shop-new'); if (el) { el.focus(); el.scrollIntoView({ block: 'center' }); } break; }
      case 'shop-toggle': shopToggle(id); break;
      case 'shop-edit': it = byId(S.shopping, id); if (it) openShopItem(it); break;
      case 'shop-single': closeSheet(); it = setShopChoice(id, 'single'); if (it) logShopping([id]); break;
      case 'shop-batch': closeSheet(); setShopChoice(id, 'batch'); render(); showToast('Đã gộp vào lần đi chợ. Mua xong bấm Ghi chi.'); break;
      case 'shop-skip': closeSheet(); setShopChoice(id, 'skip'); render(); break;
      case 'shop-log': logShopping(pendingShop().map(function (x) { return x.id; })); break;
      case 'shop-untick-del': closeSheet(); it = byId(S.shopping, id); if (it) untickShop(it, true); break;
      case 'shop-untick-keep': closeSheet(); it = byId(S.shopping, id); if (it) untickShop(Object.assign({}, it), false); break;
      case 'shop-again': {
        it = byId(S.shopping, id);
        if (it) addShopping((it.qty > 1 ? it.qty + ' ' : '') + it.name);
        break;
      }
      case 'shop-delete': persist([drop('shopping', ui.sdraft.id)]); ui.sdraft = null; go(ui.back); showToast('Đã xoá món'); break;

      // Tệp
      case 'att-open': openAttachment(id); break;
      case 'att-del': {
        it = byId(S.attachments, id);
        if (it) openSheet({ title: 'Xoá tệp?', text: '“' + esc(it.name) + '” sẽ bị xoá khỏi máy.', buttons: [{ label: 'Xoá tệp', action: 'att-del-yes', id: id, cls: 'danger' }, { label: 'Giữ lại', action: 'sheet-close', cls: 'ghost' }] });
        break;
      }
      case 'att-del-yes': {
        it = byId(S.attachments, id); closeSheet();
        if (it) {
          persist([drop('attachments', id), { store: 'files', del: id }, { store: 'files', del: id + ':t' }]);
          if (thumbUrls[id]) { URL.revokeObjectURL(thumbUrls[id]); delete thumbUrls[id]; }
          render();
        }
        break;
      }
      case 'big-yes': closeSheet(); if (pendingFiles) saveFiles(pendingFiles.files, pendingFiles.type, pendingFiles.owner); pendingFiles = null; break;
      case 'big-no': {
        closeSheet();
        if (pendingFiles) {
          var small = pendingFiles.files.filter(function (f) { return f.size <= BIG_FILE; });
          if (small.length) saveFiles(small, pendingFiles.type, pendingFiles.owner);
        }
        pendingFiles = null; break;
      }
      case 'viewer-close': URL.revokeObjectURL(ui.viewer.url); ui.viewer = null; renderOverlay(); break;
      case 'sheet-close': closeSheet(); break;

      // Cài đặt
      case 'cat-hide': {
        var c = CATS[b.getAttribute('data-code')];
        if (c) { c.hidden = !c.hidden; persist([{ store: 'categories', put: c }]); refreshCats(); render(); }
        break;
      }
      case 'cat-del': {
        var code = b.getAttribute('data-code');
        S.categories = S.categories.filter(function (x) { return x.code !== code; });
        persist([{ store: 'categories', del: code }]);
        Object.keys(S.budgets).forEach(function (m) { if (S.budgets[m][code] != null) { delete S.budgets[m][code]; persist([{ store: 'budgets', month: m }]); } });
        refreshCats(); render(); break;
      }
      case 'group-hide': {
        var g = byId(S.groups, id);
        if (g) { persist([upsert('groups', Object.assign({}, g, { hidden: !g.hidden }))]); render(); }
        break;
      }
      case 'group-del': persist([drop('groups', id)]); render(); break;
    }
  });

  // Hạn mức: lưu ngay khi gõ, vẽ lại khi rời ô (để không ngắt việc gõ)
  function setBudget(k, raw) {
    var b = S.budgets[ui.month] || (S.budgets[ui.month] = {});
    if (raw) b[k] = +raw; else delete b[k];
    if (!Object.keys(b).length) delete S.budgets[ui.month];
    persist([{ store: 'budgets', month: ui.month }]);
  }

  document.addEventListener('change', function (ev) {
    var el = ev.target;
    if (!el.getAttribute || !S) return;
    if (el.getAttribute('data-restore')) { readRestoreFile(el.files[0]); el.value = ''; return; }
    if (el.getAttribute('data-attach')) { handleAttach(el); return; }
    var k = el.getAttribute('data-budget');
    if (k) { render(); showToast('Đã lưu hạn mức ' + catOf(k).name); return; }
    var tk = el.getAttribute('data-tinput');
    if (tk === 'repeat' || tk === 'costCat' || tk === 'due') { ui.tdraft[tk] = el.value; if (tk !== 'costCat') render(); return; }
    var sk = el.getAttribute('data-sinput');
    if (sk === 'group' || sk === 'newCat') { ui.sdraft[sk] = el.value; render(sk === 'group' && el.value === '__new' ? 's-gname' : null); return; }
    if (el.getAttribute('data-input') === 'date' && ui.view === 'form') { ui.draft.date = el.value; render(); return; }
    var cn = el.getAttribute('data-catname');
    if (cn && CATS[cn]) {
      var nm = el.value.trim();
      if (nm) { CATS[cn].name = nm; if (!CATS[cn].builtin) CATS[cn].short = cut(nm, 10); persist([{ store: 'categories', put: CATS[cn] }]); showToast('Đã đổi tên danh mục ' + cn); }
      else el.value = CATS[cn].name;
      return;
    }
    var gn = el.getAttribute('data-gname'), gc = el.getAttribute('data-gcat');
    if (gn || gc) {
      var g = byId(S.groups, gn || gc);
      if (!g) return;
      if (gn) { if (!el.value.trim()) { el.value = g.name; return; } g = Object.assign({}, g, { name: el.value.trim() }); }
      else g = Object.assign({}, g, { cat: el.value });
      persist([upsert('groups', g)]); showToast('Đã lưu nhóm ' + g.name);
      return;
    }
    var st = el.getAttribute('data-setting');
    if (st === 'summaryTime' && /^\d{2}:\d{2}$/.test(el.value)) { S.settings.summaryTime = el.value; persist([metaOp('settings')]); showToast('Thẻ Tổng kết ngày hiện từ ' + el.value); }
  });

  function moneyInput(el) {
    var raw = digits(el.value).replace(/^0+/, '').slice(0, 13);
    el.value = raw ? money(+raw) : '';
    return raw;
  }

  document.addEventListener('input', function (ev) {
    var el = ev.target;
    if (!el.getAttribute || !S) return;
    var bk = el.getAttribute('data-budget');
    if (bk) { setBudget(bk, moneyInput(el)); return; }
    var tk = el.getAttribute('data-tinput');
    if (tk) {
      if (tk === 'cost') { ui.tdraft.cost = moneyInput(el); delete ui.errors.cost; }
      else if (tk !== 'repeat' && tk !== 'costCat') ui.tdraft[tk] = el.value;
      return;
    }
    var sk = el.getAttribute('data-sinput');
    if (sk) {
      if (sk === 'price') ui.sdraft.price = moneyInput(el);
      else if (sk === 'qty') { el.value = digits(el.value).slice(0, 3); ui.sdraft.qty = el.value; }
      else if (sk !== 'group' && sk !== 'newCat') ui.sdraft[sk] = el.value;
      return;
    }
    var k = el.getAttribute('data-input');
    if (!k) return;
    if (k === 'query') {
      // chỉ vẽ lại danh sách, không đụng ô tìm kiếm để gõ tiếng Việt (Telex/VNI) không bị ngắt
      ui.query = el.value;
      var lb = document.getElementById('list-body');
      if (lb) lb.innerHTML = listBody();
      return;
    }
    var d = ui.draft;
    if (k === 'amount') {
      // Cho gõ số có dấu phẩy (1,5) rồi bấm × triệu; số thường thì tự tách hàng nghìn
      var v = el.value.replace(/\./g, '');
      if (/,/.test(v)) {
        var m = v.replace(/[^\d,]/g, '').match(/^(\d{0,6}),?(\d{0,3})/);
        d.amountText = m ? m[1] + ',' + m[2] : '';
        el.value = d.amountText;
      } else { d.amountText = ''; d.amount = moneyInput(el); }
      delete ui.errors.amount;
      return;
    }
    if (k === 'reimbursed') { d.reimbursed = el.checked; return; }
    if (k === 'content') {
      d.content = el.value;
      var box = document.getElementById('sugg');
      if (box) box.innerHTML = suggHTML();
      return;
    }
    if (k === 'date') return; // xử lý khi đổi xong (change)
    d[k] = el.value;
  });

  document.addEventListener('submit', function (ev) {
    var id = ev.target.id;
    if (!S) return;
    ev.preventDefault();
    if (id === 'txform') saveExpense();
    else if (id === 'taskform') saveTask();
    else if (id === 'shopform') saveShopItem();
    else if (id === 'shopadd') { var v = document.getElementById('shop-new').value; if (v.trim()) addShopping(v); }
    else if (id === 'catadd') addCategory();
    else if (id === 'groupadd') addGroup();
  });

  document.addEventListener('keydown', function (ev) {
    if (ev.key !== 'Escape') return;
    if (ui.viewer) { URL.revokeObjectURL(ui.viewer.url); ui.viewer = null; renderOverlay(); }
    else if (ui.sheet) closeSheet();
  });

  // Mở lại app sau nửa đêm hoặc sau 22:00: vẽ lại để thẻ Hôm nay / Tổng kết đúng giờ
  document.addEventListener('visibilitychange', function () {
    if (!document.hidden && S && (ui.view === 'today' || ui.view === 'calendar')) render();
  });

  // ---------- Khởi động ----------
  main.innerHTML = '<p class="lead center" style="padding-top:40vh">Đang mở dữ liệu…</p>';
  Store.init().then(function (s) {
    S = s;
    refreshCats();
    ui.month = latestMonth();
    ui.year = ui.month.slice(0, 4);
    ui.calMonth = monthOf(todayISO());
    ui.calDay = todayISO();
    var q = /[?&]add=(\w+)/.exec(location.search);
    render();
    if (q && q[1] === 'expense') openExpense(null);
    else if (q && q[1] === 'task') openTask(null);
    else if (q && q[1] === 'shop') { ui.taskFilter = 'shop'; go('tasks'); }
    if (Store.migrated) showToast('Đã chuyển ' + S.expenses.length + ' khoản thu chi sang MyLife Hub, không mất khoản nào.');
  });
})();
