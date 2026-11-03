/* ===== XE ĐẠP TƯỜNG HIỀN – MÁY CHỦ LƯU TRỮ TRÊN GOOGLE DRIVE (bản 2) =====
   Dán toàn bộ file này vào Google Apps Script (script.google.com).
   Tài khoản/mật khẩu bên dưới chỉ mình bạn nhìn thấy, khách không xem được. */
const USER = '';
const PASS = '';
const FOLDER_ID = '1nrdQ5sL5FojUedsypbnFxsGN4bA44GKT'; // thư mục Drive của bạn

const EMPTY = { updated: 0, products: [], shop: { phone: '', address: '' } };

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function folder_() { return DriveApp.getFolderById(FOLDER_ID); }
function sub_(name) {
  const f = folder_(), it = f.getFoldersByName(name);
  return it.hasNext() ? it.next() : f.createFolder(name);
}
function file_(name) {
  const it = folder_().getFilesByName(name);
  return it.hasNext() ? it.next() : null;
}
function read_(name, def) {
  const f = file_(name);
  if (!f) return def;
  try { return JSON.parse(f.getBlob().getDataAsString()); } catch (e) { return def; }
}
function write_(name, obj) {
  const s = JSON.stringify(obj), f = file_(name);
  if (f) f.setContent(s); else folder_().createFile(name, s, 'application/json');
}
function idOf_(url) {
  const m = String(url).match(/[?&]id=([^&]+)/);
  return m ? m[1] : null;
}

// Khách mở trang -> lấy dữ liệu sản phẩm (ai cũng đọc được, không sửa được)
function doGet(e) {
  return json_({ ok: true, state: read_('data.json', EMPTY) });
}

// Chủ shop lưu thay đổi (bắt buộc đúng tài khoản + mật khẩu)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.u !== USER || b.p !== PASS) return json_({ ok: false, error: 'Sai tài khoản hoặc mật khẩu' });
    if (b.action === 'login') return json_({ ok: true });
    if (b.action === 'log') return json_({ ok: true, log: read_('log.json', []) });
    if (b.action !== 'save') return json_({ ok: false, error: 'Yêu cầu không hợp lệ' });

    const props = PropertiesService.getScriptProperties();
    let inline = props.getProperty('inlineImages') === '1'; // true = Google chặn chia sẻ ảnh -> lưu ảnh ngay trong data.json
    const old = read_('data.json', EMPTY), st = b.state, keep = {};

    st.products = (st.products || []).map(function (p) {
      p.imgs = (p.imgs || []).map(function (s) {
        s = String(s);
        if (s.indexOf('data:image') !== 0) {            // ảnh đã có sẵn trên Drive
          const id = idOf_(s);
          if (id) keep[id] = 1;
          return s;
        }
        if (inline) return s;                           // giữ ảnh dạng dữ liệu, không cần chia sẻ
        const m = s.match(/^data:(image\/[\w.+-]+);base64,(.+)$/);
        if (!m) return null;
        try {                                           // thử lưu ảnh thành file Drive công khai
          const f = sub_('anh').createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], p.id + '-' + Date.now() + '.jpg'));
          try {
            f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (e2) { f.setTrashed(true); throw e2; }
          keep[f.getId()] = 1;
          return 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1000';
        } catch (e3) {                                  // bị chặn -> chuyển sang lưu ảnh trong data.json
          inline = true;
          props.setProperty('inlineImages', '1');
          return s;
        }
      }).filter(Boolean);
      return p;
    });

    // dọn ảnh của sản phẩm/ảnh đã xóa
    (old.products || []).forEach(function (p) {
      (p.imgs || []).forEach(function (u) {
        const id = idOf_(u);
        if (id && !keep[id]) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e4) {} }
      });
    });

    st.updated = Date.now();
    const out = { updated: st.updated, products: st.products, shop: st.shop || EMPTY.shop };
    write_('data.json', out);
    const lg = read_('log.json', []);
    lg.unshift({ t: st.updated, m: b.msg || 'Cập nhật' });
    write_('log.json', lg.slice(0, 500));
    return json_({ ok: true, state: out });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Chạy MỘT LẦN để cấp quyền truy cập Google Drive cho đoạn mã này
function authorize() {
  folder_().getName();
  DriveApp.getRootFolder();
  Logger.log('Đã cấp quyền xong.');
}

// (Tùy chọn) Nếu sau này Google cho phép chia sẻ ảnh, chạy hàm này để quay lại cách lưu ảnh thành file Drive
function resetImageMode() {
  PropertiesService.getScriptProperties().deleteProperty('inlineImages');
}
