/* MyLife Hub: lưu dữ liệu ngay trên máy.
   Dữ liệu chính nằm trong IndexedDB (chứa được ảnh, tài liệu). Toàn bộ bản ghi (trừ nội dung tệp) được nạp vào bộ nhớ
   khi mở app, mỗi lần sửa thì ghi xuống ngay. Lần đầu mở bản mới, dữ liệu của Sổ chi tiêu cũ (localStorage
   'sochitieu.v1') được chuyển sang, đối chiếu số khoản và tổng tiền từng tháng; bản cũ được giữ nguyên để phòng hờ. */
(function () {
  'use strict';

  var DB_NAME = 'mylifehub', DB_VER = 2; // v2: thêm lịch hẹn, hoá đơn định kỳ, phiên tập trung
  var LEGACY_KEY = 'sochitieu.v1', LOCAL_KEY = 'mylife.v2';
  var SCHEMA = 2;
  // Bảng có khoá nằm trong bản ghi; 'meta' (cài đặt, giá đã nhớ) và 'files' (nội dung tệp) dùng khoá riêng
  var STORES = { expenses: 'id', budgets: 'month', categories: 'code', tasks: 'id', shopping: 'id', groups: 'id', attachments: 'id', events: 'id', bills: 'id', focus: 'id' };
  var LISTS = ['expenses', 'categories', 'tasks', 'shopping', 'groups', 'attachments', 'events', 'bills', 'focus'];

  var DEFAULT_CATEGORIES = [
    { code: '+', name: 'Thu nhập', short: 'Thu nhập', kind: 'thu', color: 'var(--cat-plus)' },
    { code: 'A', name: 'Ăn chơi', short: 'Ăn chơi', kind: 'chi', color: 'var(--cat-A)' },
    { code: 'C', name: 'Thiết yếu cá nhân', short: 'Cá nhân', kind: 'chi', color: 'var(--cat-C)' },
    { code: 'G', name: 'Gia đình', short: 'Gia đình', kind: 'chi', color: 'var(--cat-G)' },
    { code: 'K', name: 'Chi khác', short: 'Khác', kind: 'chi', color: 'var(--cat-K)' },
    { code: 'T', name: 'Công ty', short: 'Công ty', kind: 'chi', color: 'var(--cat-T)', reimbursable: true },
    { code: 'B', name: 'Cho con', short: 'Cho con', kind: 'chi', color: 'var(--cat-B)' }
  ].map(function (c, i) { c.sort = i; c.hidden = false; c.builtin = true; c.reimbursable = !!c.reimbursable; return c; });

  var DEFAULT_GROUPS = [
    { id: 'g-food', name: 'Thực phẩm', cat: 'G', keywords: ['thit', 'ca', 'ca hoi', 'tom', 'cua', 'muc', 'rau', 'cu', 'qua', 'trai cay', 'trung', 'gao', 'sua tuoi', 'sua chua', 'mi', 'bun', 'pho', 'banh mi', 'dau an', 'nuoc mam', 'muoi', 'duong', 'bot ngot', 'hat nem', 'nuoc tuong', 'hat tieu', 'tieu xay', 'hanh', 'toi', 'ot', 'chanh', 'dau hu', 'gio', 'cha', 'xuc xich', 'nuoc ngot', 'bia', 'tra', 'ca phe', 'banh'] },
    { id: 'g-home', name: 'Đồ gia dụng', cat: 'G', keywords: ['bot giat', 'nuoc giat', 'nuoc xa', 'nuoc rua chen', 'giay ve sinh', 'khan giay', 'bong den', 'pin', 'tui rac', 'nuoc lau', 'chen', 'dia', 'noi', 'chao', 'choi', 'cay lau', 'mang boc', 'giay bac', 'xa phong', 'nuoc tay'] },
    { id: 'g-self', name: 'Cá nhân', cat: 'C', keywords: ['dau goi', 'dau xa', 'sua tam', 'kem danh rang', 'ban chai', 'dao cao', 'lan khu mui', 'kem chong nang', 'khau trang', 'thuoc', 'vitamin'] },
    { id: 'g-kid', name: 'Cho con', cat: 'B', keywords: ['ta', 'bim', 'sua bot', 'vo', 'but', 'thuoc ke', 'cuc tay', 'do choi', 'sach', 'cap sach', 'dong phuc', 'khan uot'] },
    { id: 'g-other', name: 'Khác', cat: 'K', keywords: [] }
  ].map(function (g, i) { g.sort = i; g.hidden = false; g.builtin = g.id === 'g-other'; return g; });

  var DEFAULT_SETTINGS = { summaryTime: '22:00', digestTime: '07:30', budgetWarn: 0.8, focusWork: 60, focusBreak: 5, focusRun: null,
    notify: false, notified: {}, shopChoice: 'single', migratedFrom: null };

  function req(r) { return new Promise(function (ok, fail) { r.onsuccess = function () { ok(r.result); }; r.onerror = function () { fail(r.error); }; }); }
  function txDone(tx) {
    return new Promise(function (ok, fail) {
      tx.oncomplete = function () { ok(); };
      tx.onerror = function () { fail(tx.error); };
      tx.onabort = function () { fail(tx.error || new Error('abort')); };
    });
  }

  function openIDB() {
    return new Promise(function (ok, fail) {
      if (!window.indexedDB) { fail(new Error('no idb')); return; }
      var r;
      try { r = indexedDB.open(DB_NAME, DB_VER); } catch (e) { fail(e); return; }
      r.onupgradeneeded = function () {
        var d = r.result;
        Object.keys(STORES).forEach(function (s) { if (!d.objectStoreNames.contains(s)) d.createObjectStore(s, { keyPath: STORES[s] }); });
        ['meta', 'files'].forEach(function (s) { if (!d.objectStoreNames.contains(s)) d.createObjectStore(s); });
      };
      r.onsuccess = function () {
        var d = r.result;
        // Một bản app mới hơn cần nâng cấp bộ nhớ: đóng kết nối này để bản kia mở được
        d.onversionchange = function () { d.close(); };
        ok(d);
      };
      r.onerror = function () { fail(r.error); };
      // Thẻ app cũ còn mở: chờ nó đóng (không chuyển sang bộ nhớ tạm, tránh hiện dữ liệu cũ)
      r.onblocked = function () {};
    });
  }

  function emptyState() {
    return { expenses: [], budgets: {}, categories: [], tasks: [], shopping: [], groups: [], attachments: [], events: [], bills: [], focus: [], settings: {}, prices: {} };
  }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  // 5 khoản tháng 5 lấy từ file Excel "quanlychitieuhangngay2025" (giống bản Sổ chi tiêu khi mở lần đầu)
  function seedItems() {
    return [
      { type: 'chi', cat: 'C', content: 'Cafe', amount: 18000 },
      { type: 'thu', cat: '+', content: 'Lương tháng 5', amount: 9334000 },
      { type: 'chi', cat: 'K', content: 'Hụi', amount: 1000000 },
      { type: 'chi', cat: 'K', content: 'Trả tiền Thiện', amount: 1000000 },
      { type: 'chi', cat: 'K', content: 'Trả tiền A.Tuấn', amount: 1000000 }
    ].map(function (t, i) {
      return { id: 'seed' + i, date: '2026-05-29', type: t.type, cat: t.cat, content: t.content, amount: t.amount, note: 'Nhập từ file Excel', createdAt: i };
    });
  }

  // Một khoản của bản cũ -> khoản thu chi của bản mới. Khoản Công ty (T) cũ coi như đã được hoàn tiền,
  // để app không báo "công ty còn nợ" sai cho các tháng trước.
  function fromLegacyItem(t) {
    return {
      id: t.id, date: t.date, type: t.type, cat: t.cat, content: t.content, amount: t.amount, note: t.note || '',
      createdAt: t.createdAt || 0, source: 'manual', sourceIds: [], reimburse: t.cat === 'T' ? 'da_hoan' : null
    };
  }

  function legacyState(v1) {
    var s = emptyState();
    var items = v1 && Array.isArray(v1.items) ? v1.items : seedItems();
    s.expenses = items.map(fromLegacyItem);
    s.budgets = v1 && v1.budgets && typeof v1.budgets === 'object' ? clone(v1.budgets) : {};
    s.categories = clone(DEFAULT_CATEGORIES);
    s.groups = clone(DEFAULT_GROUPS);
    s.settings = Object.assign({}, DEFAULT_SETTINGS, { migratedFrom: v1 ? LEGACY_KEY : null, lastBackup: v1 && v1.lastBackup || null });
    return s;
  }

  // Số khoản + tổng tiền theo từng tháng, để đối chiếu sau khi chuyển dữ liệu
  function fingerprint(list) {
    var by = {};
    list.forEach(function (t) {
      var k = t.date.slice(0, 7) + t.type;
      by[k] = (by[k] || 0) + t.amount;
    });
    return list.length + '|' + Object.keys(by).sort().map(function (k) { return k + ':' + by[k]; }).join(',');
  }

  function readLegacy() {
    try { var raw = localStorage.getItem(LEGACY_KEY); return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
  }

  function readLocal() {
    try { var raw = localStorage.getItem(LOCAL_KEY); return raw ? JSON.parse(raw) : null; } catch (e) { return null; }
  }

  function fillDefaults(s) {
    LISTS.forEach(function (k) { if (!Array.isArray(s[k])) s[k] = []; });
    if (!s.categories.length) s.categories = clone(DEFAULT_CATEGORIES);
    if (!s.groups.length) s.groups = clone(DEFAULT_GROUPS);
    s.settings = Object.assign({}, DEFAULT_SETTINGS, s.settings || {});
    s.prices = s.prices || {};
    return s;
  }

  var Store = {
    mode: 'idb',        // idb | local (localStorage, không lưu được tệp) | memory (chỉ giữ trong phiên)
    warning: '',
    migrated: false,
    idb: null,
    S: null,
    memFiles: {},

    init: function () {
      var self = this;
      return openIDB().then(function (d) {
        self.idb = d;
        return self._readAll();
      }).then(function (s) {
        if (s.settings && s.settings.schema === SCHEMA) return fillDefaults(s);
        return self._migrate();
      }).catch(function () {
        return self._initLocal();
      }).then(function (s) {
        self.S = s;
        try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) {}
        return s;
      });
    },

    _readAll: function () {
      var tx = this.idb.transaction(Object.keys(STORES).concat('meta'), 'readonly');
      var s = emptyState(), jobs = [];
      LISTS.forEach(function (k) { jobs.push(req(tx.objectStore(k).getAll()).then(function (v) { s[k] = v; })); });
      jobs.push(req(tx.objectStore('budgets').getAll()).then(function (rows) { rows.forEach(function (r) { s.budgets[r.month] = r.limits; }); }));
      jobs.push(req(tx.objectStore('meta').get('settings')).then(function (v) { s.settings = v || {}; }));
      jobs.push(req(tx.objectStore('meta').get('prices')).then(function (v) { s.prices = v || {}; }));
      return Promise.all(jobs).then(function () { return s; });
    },

    _writeAll: function (s, files) {
      var names = Object.keys(STORES).concat('meta', 'files');
      var tx = this.idb.transaction(names, 'readwrite');
      names.forEach(function (n) { tx.objectStore(n).clear(); });
      LISTS.forEach(function (k) { s[k].forEach(function (r) { tx.objectStore(k).put(r); }); });
      Object.keys(s.budgets).forEach(function (m) { tx.objectStore('budgets').put({ month: m, limits: s.budgets[m] }); });
      tx.objectStore('meta').put(s.prices, 'prices');
      tx.objectStore('meta').put(Object.assign({}, s.settings, { schema: SCHEMA }), 'settings');
      if (files) Object.keys(files).forEach(function (id) { tx.objectStore('files').put(files[id], id); });
      return txDone(tx);
    },

    _migrate: function () {
      // Ưu tiên bản đã lưu tạm ở localStorage (nếu lần trước trình duyệt không cho dùng IndexedDB), sau đó mới tới Sổ chi tiêu cũ
      var self = this, v1 = readLegacy(), local = readLocal();
      var s = local ? fillDefaults(Object.assign(emptyState(), local)) : legacyState(v1);
      var want = fingerprint(s.expenses);
      return this._writeAll(s).then(function () { return self._readAll(); }).then(function (back) {
        if (fingerprint(back.expenses) !== want) throw new Error('verify');
        self.migrated = !!(v1 || local);
        try { if (local) localStorage.removeItem(LOCAL_KEY); } catch (e) {}
        return fillDefaults(back);
      }).catch(function (err) {
        // Đối chiếu không khớp: dừng lại, dùng dữ liệu cũ (vẫn còn nguyên trong localStorage)
        self.idb = null;
        self.warning = 'Chưa chuyển được dữ liệu sang bộ nhớ mới (' + (err && err.message || 'lỗi') + '). App đang dùng dữ liệu cũ; tệp đính kèm tạm thời chưa lưu được.';
        return self._initLocal(true);
      });
    },

    _initLocal: function (keepWarning) {
      var s = null;
      try {
        var raw = localStorage.getItem(LOCAL_KEY);
        if (raw) s = fillDefaults(Object.assign(emptyState(), JSON.parse(raw)));
        this.mode = 'local';
      } catch (e) { this.mode = 'memory'; }
      if (!s) s = fillDefaults(legacyState(readLegacy()));
      this.idb = null;
      if (!keepWarning) {
        this.warning = this.mode === 'memory'
          ? 'Trình duyệt đang chặn bộ nhớ, dữ liệu sẽ mất khi đóng trang.'
          : 'Trình duyệt không cho dùng bộ nhớ lớn: dữ liệu vẫn lưu, nhưng ảnh và tài liệu đính kèm chỉ giữ đến khi đóng app.';
      }
      if (this.mode === 'local') this._saveLocal(s);
      return s;
    },

    _saveLocal: function (s) {
      if (this.mode !== 'local') return Promise.resolve();
      try { localStorage.setItem(LOCAL_KEY, JSON.stringify(s)); return Promise.resolve(); }
      catch (e) { return Promise.reject(e); }
    },

    // ops: [{ store, put: record }] hoặc [{ store, del: key }]. store = 'budgets' dùng { store, month } (ghi theo S.budgets);
    // store = 'meta' dùng { store, key: 'settings' | 'prices' }. Tất cả ghi trong một lần: lỗi giữa chừng thì không lưu gì.
    commit: function (ops) {
      var S = this.S;
      if (!this.idb) return this._saveLocal(S);
      var names = {};
      ops.forEach(function (o) { names[o.store] = 1; });
      var tx = this.idb.transaction(Object.keys(names), 'readwrite');
      ops.forEach(function (o) {
        var st = tx.objectStore(o.store);
        if (o.store === 'budgets') {
          if (S.budgets[o.month]) st.put({ month: o.month, limits: S.budgets[o.month] }); else st.delete(o.month);
        } else if (o.store === 'meta') {
          st.put(o.key === 'settings' ? Object.assign({}, S.settings, { schema: SCHEMA }) : S[o.key], o.key);
        } else if (o.store === 'files') {
          if (o.del) st.delete(o.del); else st.put(o.put, o.key);
        } else if (o.del) st.delete(o.del);
        else st.put(o.put);
      });
      return txDone(tx);
    },

    getFile: function (id) {
      if (!this.idb) return Promise.resolve(this.memFiles[id] || null);
      return req(this.idb.transaction('files').objectStore('files').get(id));
    },
    getAllFiles: function () {
      if (!this.idb) return Promise.resolve(Object.assign({}, this.memFiles));
      var out = {}, st = this.idb.transaction('files').objectStore('files');
      return new Promise(function (ok, fail) {
        var c = st.openCursor();
        c.onsuccess = function () { var cur = c.result; if (cur) { out[cur.key] = cur.value; cur.continue(); } else ok(out); };
        c.onerror = function () { fail(c.error); };
      });
    },
    putFiles: function (map) {
      var self = this;
      if (!this.idb) { Object.keys(map).forEach(function (k) { self.memFiles[k] = map[k]; }); return Promise.resolve(); }
      return this.commit(Object.keys(map).map(function (k) { return { store: 'files', key: k, put: map[k] }; }));
    },

    // Khôi phục: thay toàn bộ dữ liệu
    replaceAll: function (s, files) {
      var self = this;
      s = fillDefaults(s);
      if (!this.idb) {
        this.S = s; this.memFiles = files || {};
        return this._saveLocal(s).then(function () { return s; });
      }
      return this._writeAll(s, files).then(function () { self.S = s; return s; });
    },

    estimate: function () {
      try { if (navigator.storage && navigator.storage.estimate) return navigator.storage.estimate(); } catch (e) {}
      return Promise.resolve(null);
    },

    legacyState: legacyState,
    fromLegacyItem: fromLegacyItem,
    lists: LISTS,
    defaults: { categories: DEFAULT_CATEGORIES, groups: DEFAULT_GROUPS, settings: DEFAULT_SETTINGS }
  };

  window.MyLifeStore = Store;
})();
