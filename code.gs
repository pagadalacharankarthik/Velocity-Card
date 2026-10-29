// ============================================================
// VELOCITY CARD ENGINE — Google Apps Script Backend
// Deploy as: Web App → Anyone (for public form)
// Sheet tabs auto-created on first run
// ============================================================

const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
const ADMIN_PASSWORD = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
const CLOUDINARY_UPLOAD_PRESET = 'YOUR_PRESET'; // ← Optional: for image hosting
const CLOUDINARY_CLOUD_NAME = 'YOUR_CLOUD';

// Plan definitions
const PLANS = {
  basic:    { price: 499,  name: 'Basic',    templates: ['temple','floral','minimal'], photos: 3, music: false, memory: false, validity_days: 90  },
  standard: { price: 999,  name: 'Standard', templates: ['all'],                      photos: 8, music: true,  memory: false, validity_days: 180 },
  premium:  { price: 1499, name: 'Premium',  templates: ['all'],                      photos: 20, music: true, memory: true,  validity_days: 365 }
};

// Sheet column definitions
const ORDERS_COLS = [
  'ORDER_ID','CREATED_AT','STATUS','PLAN',
  'SHOP_NAME','CONTACT_NAME','PHONE','WHATSAPP','EMAIL',
  'BRIDE_NAME','BRIDE_NAME_REGIONAL','GROOM_NAME','GROOM_NAME_REGIONAL',
  'LANGUAGE','RELIGION','TEMPLATE',
  'WEDDING_DATE','WEDDING_TIME','WEDDING_VENUE','WEDDING_ADDRESS','MAPS_LINK',
  'RECEPTION_DATE','RECEPTION_TIME','RECEPTION_VENUE',
  'RSVP_CONTACT','CUSTOM_MESSAGE',
  'BRIDE_PHOTO_URL','GROOM_PHOTO_URL','COUPLE_PHOTO_URL',
  'EXTRA_PHOTOS','MUSIC_URL','VIDEO_URL',
  'AMOUNT','PAYMENT_STATUS','PAYMENT_REF',
  'PUBLISHED_URL','QR_PREVIEW_URL','PUBLISHED_AT',
  'VIEWS','RSVP_COUNT','ADMIN_NOTES',
  'BRIDE_PHOTO_SIZE','GROOM_PHOTO_SIZE','COUPLE_PHOTO_SIZE',
  'INVITATION_DATA'
];

const RSVP_COLS = ['RSVP_ID','ORDER_ID','GUEST_NAME','ATTENDANCE','MESSAGE','PHONE','CREATED_AT'];
const WISHES_COLS = ['WISH_ID','ORDER_ID','GUEST_NAME','WISH_TEXT','PHOTO_URL','CREATED_AT'];
const ANALYTICS_COLS = ['LOG_ID','ORDER_ID','EVENT','IP','DEVICE','SOURCE','CREATED_AT'];
const SHOPS_COLS = ['SHOP_ID','SHOP_NAME','EMAIL','PHONE','PLAN','ORDERS_USED','ORDERS_LIMIT','JOINED_AT','STATUS'];

// ============================================================
// MAIN ROUTER
// ============================================================
function doPost(e) {
  try {
    // Debug log incoming request (helps diagnose missing payloads)
    try {
      var dbg = {
        parameter: e && e.parameter ? e.parameter : {},
        postType: e && e.postData ? e.postData.type : null,
        postContentsPreview: e && e.postData && e.postData.contents ? String(e.postData.contents).slice(0,1000) : ''
      };
      Logger.log('doPost incoming: %s', JSON.stringify(dbg));
    } catch (li) {}
    // Parse incoming data (supports application/json and x-www-form-urlencoded)
    const data = parseFormData(e);
    const action = data.action || '';

    switch (action) {
      case 'submit_order':     return res(submitOrder(data), e);
      case 'upload_photo':     return res(handlePhotoUpload(data), e);
      case 'admin_login':      return res(adminLogin(data), e);
      case 'get_admin_dashboard': return res(getAdminDashboard(data), e);
      case 'get_orders':       return res(getOrders(data), e);
      case 'update_order':     return res(updateOrder(data), e);
      case 'publish_order':    return res(publishOrder(data), e);
      case 'submit_rsvp':      return res(submitRsvp(data), e);
      case 'submit_wish':      return res(submitWish(data), e);
      case 'get_invitation':   return res(getInvitation(data), e);
      case 'log_view':         return res(logAnalytics(data), e);
      case 'get_analytics':    return res(getAnalytics(data), e);
      case 'get_earnings':     return res(getEarnings(data), e);
      case 'verify_payment':   return res(verifyPayment(data), e);
      case 'generate_qr':      return res(generateQRPreview(data), e);
      case 'register_shop':    return res(registerShop(data), e);
      case 'shop_login':       return res(shopLogin(data), e);
      case 'get_shop_orders':  return res(getShopOrders(data), e);
      default: return res({ success: false, error: 'Unknown action' }, e);
    }
  } catch (err) {
    return res({ success: false, error: err.toString() }, e);
  }
}

// Robust parser for Apps Script POST payloads
function parseFormData(e) {
  const params = {};
  try {
    if (e.postData && e.postData.type && e.postData.type.indexOf('application/x-www-form-urlencoded') !== -1) {
      const pairs = (e.postData.contents || '').split('&').filter(Boolean);
      pairs.forEach(function(p) {
        var idx = p.indexOf('=');
        var k = decodeURIComponent(idx >= 0 ? p.slice(0, idx) : p);
        var v = decodeURIComponent(idx >= 0 ? p.slice(idx + 1) : '');
        try { params[k] = JSON.parse(v); } catch(err) { params[k] = v; }
      });
    } else if (e.postData && e.postData.contents) {
      try { Object.assign(params, JSON.parse(e.postData.contents)); } catch(err) { /* ignore */ }
    }
  } catch (ex) {}
  // merge query params if present
  if (e.parameter) for (var k in e.parameter) if (!(k in params)) params[k] = e.parameter[k];
  return params;
}

function doGet(e) {
  const order_id = e.parameter.order_id;
  const preview = e.parameter.preview;
  if (order_id) return res(getInvitation({ order_id, preview }));
  return ContentService.createTextOutput('Velocity Card Engine Active');
}

function res(data) {
  // If caller provided a JSONP callback (e.g., ?callback=fn) return JS wrapper
  var callback = null;
  try { if (typeof arguments[1] === 'object' && arguments[1].parameter && arguments[1].parameter.callback) callback = arguments[1].parameter.callback; } catch(e) {}
  if (callback) {
    var body = callback + '(' + JSON.stringify(data) + ');';
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SHEET SETUP
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  ensureSheet(ss, 'ORDERS', ORDERS_COLS);
  ensureSheet(ss, 'RSVP', RSVP_COLS);
  ensureSheet(ss, 'WISHES', WISHES_COLS);
  ensureSheet(ss, 'ANALYTICS', ANALYTICS_COLS);
  ensureSheet(ss, 'SHOPS', SHOPS_COLS);
  
  // Config sheet for plan/pricing data
  let cfg = ss.getSheetByName('CONFIG');
  if (!cfg) {
    cfg = ss.insertSheet('CONFIG');
    cfg.getRange('A1:B10').setValues([
      ['KEY','VALUE'],
      ['BASIC_PRICE','499'],
      ['STANDARD_PRICE','999'],
      ['PREMIUM_PRICE','1499'],
      ['UPI_ID','yourname@upi'],
      ['WHATSAPP_NUM','919999999999'],
      ['PLATFORM_NAME','Charan Labs Solution'],
      ['ADMIN_EMAIL','charankarthik366@gmail.com'],
      ['CURRENCY','₹'],
      ['ACTIVE','true']
    ]);
  }
}

function ensureSheet(ss, name, cols) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
    sheet.getRange(1, 1, 1, cols.length).setFontWeight('bold').setBackground('#1a1208').setFontColor('#C9A84C');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// ORDER SUBMISSION
// ============================================================
function submitOrder(data) {
  // Defensive: ensure we received a payload
  if (!data || typeof data !== 'object') return { success: false, error: 'No payload received (data is undefined). Check request body and Content-Type.' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ORDERS') || setupSheets();
  
  // Validate required fields
  const required = ['shop_name','contact_name','phone','email','bride_name','groom_name','wedding_date','plan'];
  for (const field of required) {
    if (!data[field]) return { success: false, error: `Missing field: ${field}` };
  }
  
  // Validate plan
  if (!PLANS[data.plan]) return { success: false, error: 'Invalid plan' };
  
  const order_id = 'VCE' + Date.now().toString().slice(-8);
  const now = new Date().toISOString();
  const plan = PLANS[data.plan];
  
  const row = [
    order_id, now, 'SUBMITTED', data.plan.toUpperCase(),
    data.shop_name, data.contact_name, data.phone, data.whatsapp || data.phone, data.email,
    data.bride_name, data.bride_name_regional || '', data.groom_name, data.groom_name_regional || '',
    data.language || 'English', data.religion || 'Hindu', data.template || 'auto',
    data.wedding_date, data.wedding_time || '', data.wedding_venue || '', data.wedding_address || '', data.maps_link || '',
    data.reception_date || '', data.reception_time || '', data.reception_venue || '',
    data.rsvp_contact || data.phone, data.custom_message || '',
    '', '', '',  // photo URLs (filled on upload)
    '', '', '',  // extra photos, music, video
    plan.price, 'PENDING', '',
    '', '', '',  // published URL, QR URL, published at
    0, 0, '',   // views, rsvp_count, admin_notes
    data.bride_photo_size || '', data.groom_photo_size || '', data.couple_photo_size || '',
    '{}'         // invitation_data JSON
  ];
  
  const ordersSheet = ss.getSheetByName('ORDERS');
  ordersSheet.appendRow(row);
  
  // Send confirmation email
  try {
    MailApp.sendEmail({
      to: data.email,
      subject: `✨ Order Received — ${order_id} | Charan Labs Solution`,
      htmlBody: buildConfirmationEmail(data, order_id, plan)
    });
  } catch(e) {}
  
  // WhatsApp notification to admin (via log)
  logAdminAlert(order_id, data.shop_name, plan.price);
  
  return { success: true, order_id, amount: plan.price, plan: plan.name };
}

function buildConfirmationEmail(data, order_id, plan) {
  return `
  <div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;background:#1a1208;color:#FAF7F0;padding:40px;border:1px solid #C9A84C33;">
    <div style="text-align:center;margin-bottom:30px;">
      <h1 style="font-size:28px;color:#C9A84C;font-weight:300;">✦ Charan Labs Solution</h1>
      <p style="color:#C9A84C99;letter-spacing:3px;font-size:11px;text-transform:uppercase;">Premium Digital Wedding Invitations</p>
    </div>
    <h2 style="font-size:20px;font-weight:400;margin-bottom:20px;">Order Confirmed, ${data.contact_name}</h2>
    <div style="background:rgba(201,168,76,0.08);border:1px solid rgba(201,168,76,0.2);padding:20px;margin-bottom:20px;">
      <p><strong style="color:#C9A84C;">Order ID:</strong> ${order_id}</p>
      <p><strong style="color:#C9A84C;">Plan:</strong> ${plan.name} — ₹${plan.price}</p>
      <p><strong style="color:#C9A84C;">Couple:</strong> ${data.bride_name} & ${data.groom_name}</p>
      <p><strong style="color:#C9A84C;">Wedding:</strong> ${data.wedding_date}</p>
    </div>
    <p style="color:#FAF7F099;line-height:1.8;">
      Please complete payment of <strong style="color:#C9A84C;">₹${plan.price}</strong> to our UPI ID.<br>
      Our team will begin crafting your invitation within 24 hours of payment confirmation.
    </p>
    <div style="text-align:center;margin-top:30px;padding:20px;border-top:1px solid rgba(201,168,76,0.15);">
      <p style="color:#C9A84C99;font-size:12px;font-style:italic;">Every love story deserves a masterpiece</p>
    </div>
  </div>`;
}

function logAdminAlert(order_id, shop_name, amount) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ANALYTICS');
  if (sheet) sheet.appendRow([
    'ADM' + Date.now(), order_id, 'NEW_ORDER_ALERT', shop_name, '', `₹${amount}`, new Date().toISOString()
  ]);
}

// ============================================================
// PHOTO UPLOAD (Base64 → Google Drive)
// ============================================================
function handlePhotoUpload(data) {
  const { order_id, photo_type, base64_data, file_name, file_size_kb, mime_type } = data;
  
  // Validate file size requirements
  const SIZE_REQUIREMENTS = {
    bride:   { min_kb: 100, max_kb: 5120, recommended: '1MB-3MB, min 800×1000px portrait' },
    groom:   { min_kb: 100, max_kb: 5120, recommended: '1MB-3MB, min 800×1000px portrait' },
    couple:  { min_kb: 100, max_kb: 5120, recommended: '2MB-5MB, min 1200×800px landscape' },
    extra:   { min_kb: 50,  max_kb: 8192, recommended: 'Any size, min 600px on shortest side' }
  };
  
  const req = SIZE_REQUIREMENTS[photo_type] || SIZE_REQUIREMENTS.extra;
  if (file_size_kb < req.min_kb) return { success: false, error: `Photo too small. Minimum ${req.min_kb}KB.` };
  if (file_size_kb > req.max_kb) return { success: false, error: `Photo too large. Maximum ${req.max_kb / 1024}MB.` };
  
  try {
    // Save to Google Drive
    const folder = getOrCreateFolder(order_id, photo_type);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64_data.split(',').pop()),
      mime_type || 'image/jpeg',
      file_name || `${photo_type}_${Date.now()}.jpg`
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const file_url = `https://drive.google.com/uc?id=${file.getId()}`;
    
    // Update order row
    updateOrderField(order_id, photo_type.toUpperCase() + '_PHOTO_URL', file_url);
    updateOrderField(order_id, photo_type.toUpperCase() + '_PHOTO_SIZE', `${Math.round(file_size_kb)}KB`);
    
    return { success: true, url: file_url, file_id: file.getId() };
  } catch(err) {
    return { success: false, error: 'Upload failed: ' + err.toString() };
  }
}

function getOrCreateFolder(order_id, subfolder) {
  const root_name = 'VelocityCardEngine_Photos';
  let root;
  const roots = DriveApp.getFoldersByName(root_name);
  root = roots.hasNext() ? roots.next() : DriveApp.createFolder(root_name);
  
  let order_folder;
  const of = root.getFoldersByName(order_id);
  order_folder = of.hasNext() ? of.next() : root.createFolder(order_id);
  
  let sub;
  const sf = order_folder.getFoldersByName(subfolder);
  sub = sf.hasNext() ? sf.next() : order_folder.createFolder(subfolder);
  return sub;
}

// ============================================================
// ADMIN FUNCTIONS
// ============================================================
function adminLogin(data) {
  if (data.password !== ADMIN_PASSWORD) return { success: false, error: 'Invalid password' };
  const token = Utilities.base64Encode(data.password + '_' + Date.now());
  return { success: true, token, role: 'admin' };
}

function getOrders(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ORDERS');
  if (!sheet) return { success: true, orders: [] };
  
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const orders = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  
  // Filter by status or shop if provided
  let filtered = orders;
  if (data.status && data.status !== 'ALL') filtered = filtered.filter(o => o.STATUS === data.status);
  if (data.shop_name) filtered = filtered.filter(o => o.SHOP_NAME === data.shop_name);
  
  return { success: true, orders: filtered.reverse() }; // newest first
}

function updateOrder(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const { order_id, fields } = data; // fields = {STATUS:'...', ADMIN_NOTES:'...', ...}
  
  for (const [key, value] of Object.entries(fields)) {
    updateOrderField(order_id, key, value);
  }
  return { success: true };
}

function publishOrder(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const { order_id, invitation_data } = data;
  
  const script_url = ScriptApp.getService().getUrl();
  const published_url = `${script_url}?order_id=${order_id}`;
  const qr_url = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(published_url)}`;
  
  updateOrderField(order_id, 'STATUS', 'PUBLISHED');
  updateOrderField(order_id, 'PUBLISHED_URL', published_url);
  updateOrderField(order_id, 'QR_PREVIEW_URL', qr_url);
  updateOrderField(order_id, 'PUBLISHED_AT', new Date().toISOString());
  if (invitation_data) updateOrderField(order_id, 'INVITATION_DATA', JSON.stringify(invitation_data));
  
  // Email customer with published link
  const order = getOrderById(order_id);
  if (order && order.EMAIL) {
    try {
      MailApp.sendEmail({
        to: order.EMAIL,
        subject: `🎊 Your Wedding Invitation is LIVE! — ${order_id}`,
        htmlBody: buildPublishedEmail(order, published_url, qr_url)
      });
    } catch(e) {}
  }
  
  return { success: true, published_url, qr_url };
}

function generateQRPreview(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const script_url = ScriptApp.getService().getUrl();
  const preview_url = `${script_url}?order_id=${data.order_id}&preview=true`;
  const qr_url = `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encodeURIComponent(preview_url)}`;
  updateOrderField(data.order_id, 'QR_PREVIEW_URL', qr_url);
  updateOrderField(data.order_id, 'STATUS', 'QR_PREVIEW');
  return { success: true, preview_url, qr_url };
}

function verifyPayment(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  updateOrderField(data.order_id, 'PAYMENT_STATUS', 'PAID');
  updateOrderField(data.order_id, 'PAYMENT_REF', data.payment_ref || '');
  updateOrderField(data.order_id, 'STATUS', 'PAYMENT_VERIFIED');
  return { success: true };
}

function getAnalytics(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const orders = getOrdersRaw(ss);
  const today = new Date();
  const thisMonth = today.getMonth();
  const thisYear = today.getFullYear();
  
  const summary = {
    total_orders: orders.length,
    pending: orders.filter(o => ['SUBMITTED','PENDING_REVIEW'].includes(o.STATUS)).length,
    published: orders.filter(o => o.STATUS === 'PUBLISHED').length,
    total_revenue: orders.filter(o => o.PAYMENT_STATUS === 'PAID').reduce((s, o) => s + (parseInt(o.AMOUNT) || 0), 0),
    monthly_revenue: orders.filter(o => {
      if (o.PAYMENT_STATUS !== 'PAID') return false;
      const d = new Date(o.CREATED_AT);
      return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
    }).reduce((s, o) => s + (parseInt(o.AMOUNT) || 0), 0),
    by_plan: {
      basic: orders.filter(o => o.PLAN === 'BASIC').length,
      standard: orders.filter(o => o.PLAN === 'STANDARD').length,
      premium: orders.filter(o => o.PLAN === 'PREMIUM').length
    },
    by_status: {}
  };
  
  ['SUBMITTED','PENDING_REVIEW','PHOTO_REVIEW','EDITOR_ADJUSTMENT','QR_PREVIEW','APPROVED','PAYMENT_VERIFIED','PUBLISHED','ARCHIVED'].forEach(s => {
    summary.by_status[s] = orders.filter(o => o.STATUS === s).length;
  });
  
  return { success: true, analytics: summary };
}

// OPTIMIZED: Load dashboard analytics and orders in ONE sheet read (instead of two)
function getAdminDashboard(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const orders = getOrdersRaw(ss); // Read sheet ONCE
  
  const today = new Date();
  const thisMonth = today.getMonth();
  const thisYear = today.getFullYear();
  
  const analytics = {
    total_orders: 0,
    pending: 0,
    published: 0,
    total_revenue: 0,
    monthly_revenue: 0,
    by_plan: { basic: 0, standard: 0, premium: 0 },
    by_status: {
      SUBMITTED: 0,
      PENDING_REVIEW: 0,
      PHOTO_REVIEW: 0,
      EDITOR_ADJUSTMENT: 0,
      QR_PREVIEW: 0,
      APPROVED: 0,
      PAYMENT_VERIFIED: 0,
      PUBLISHED: 0,
      ARCHIVED: 0
    }
  };
  
  orders.forEach(o => {
    analytics.total_orders += 1;
    const status = o.STATUS || '';
    const plan = (o.PLAN || '').toLowerCase();
    const amount = parseInt(o.AMOUNT) || 0;
    
    if (['SUBMITTED', 'PENDING_REVIEW'].includes(status)) analytics.pending += 1;
    if (status === 'PUBLISHED') analytics.published += 1;
    if (o.PAYMENT_STATUS === 'PAID') {
      analytics.total_revenue += amount;
      const d = new Date(o.CREATED_AT);
      if (d.getFullYear() === thisYear && d.getMonth() === thisMonth) {
        analytics.monthly_revenue += amount;
      }
    }
    if (plan in analytics.by_plan) analytics.by_plan[plan] += 1;
    if (status in analytics.by_status) analytics.by_status[status] += 1;
  });
  
  return { success: true, analytics, orders: orders.reverse() }; // Return both!
}

function getEarnings(data) {
  if (!validateAdmin(data.token)) return { success: false, error: 'Unauthorized' };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const orders = getOrdersRaw(ss).filter(o => o.PAYMENT_STATUS === 'PAID');
  
  // Group by month
  const monthly = {};
  orders.forEach(o => {
    const d = new Date(o.CREATED_AT);
    const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
    monthly[key] = (monthly[key] || 0) + (parseInt(o.AMOUNT) || 0);
  });
  
  return { success: true, earnings: { monthly, total: orders.reduce((s,o) => s + (parseInt(o.AMOUNT)||0), 0) } };
}

// ============================================================
// PUBLIC: INVITATION DATA
// ============================================================
function getInvitation(data) {
  const order = getOrderById(data.order_id);
  if (!order) return { success: false, error: 'Order not found' };
  if (!data.preview && order.STATUS !== 'PUBLISHED') return { success: false, error: 'Invitation not yet published' };
  
  // Log view
  logAnalytics({ order_id: data.order_id, event: 'VIEW', source: 'direct' });
  
  // Increment view count
  const count = parseInt(order.VIEWS || '0') + 1;
  updateOrderField(data.order_id, 'VIEWS', count);
  
  return { 
    success: true, 
    invitation: {
      order_id: data.order_id,
      bride_name: order.BRIDE_NAME,
      bride_name_regional: order.BRIDE_NAME_REGIONAL,
      groom_name: order.GROOM_NAME,
      groom_name_regional: order.GROOM_NAME_REGIONAL,
      language: order.LANGUAGE,
      religion: order.RELIGION,
      template: order.TEMPLATE,
      wedding_date: order.WEDDING_DATE,
      wedding_time: order.WEDDING_TIME,
      wedding_venue: order.WEDDING_VENUE,
      wedding_address: order.WEDDING_ADDRESS,
      maps_link: order.MAPS_LINK,
      reception_date: order.RECEPTION_DATE,
      reception_time: order.RECEPTION_TIME,
      reception_venue: order.RECEPTION_VENUE,
      rsvp_contact: order.RSVP_CONTACT,
      custom_message: order.CUSTOM_MESSAGE,
      bride_photo: order.BRIDE_PHOTO_URL,
      groom_photo: order.GROOM_PHOTO_URL,
      couple_photo: order.COUPLE_PHOTO_URL,
      extra_photos: order.EXTRA_PHOTOS ? order.EXTRA_PHOTOS.split('|') : [],
      music_url: order.MUSIC_URL,
      invitation_data: order.INVITATION_DATA ? JSON.parse(order.INVITATION_DATA) : {},
      plan: order.PLAN,
      is_preview: !!data.preview
    }
  };
}

// ============================================================
// RSVP & WISHES
// ============================================================
function submitRsvp(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('RSVP');
  const row = [
    'RSVP' + Date.now(), data.order_id,
    data.guest_name, data.attendance, data.message || '', data.phone || '',
    new Date().toISOString()
  ];
  sheet.appendRow(row);
  
  const count = parseInt(getOrderById(data.order_id)?.RSVP_COUNT || '0') + 1;
  updateOrderField(data.order_id, 'RSVP_COUNT', count);
  return { success: true };
}

function submitWish(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('WISHES');
  sheet.appendRow([
    'WISH' + Date.now(), data.order_id,
    data.guest_name, data.wish_text, data.photo_url || '',
    new Date().toISOString()
  ]);
  return { success: true };
}

// ============================================================
// ANALYTICS
// ============================================================
function logAnalytics(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('ANALYTICS');
    sheet.appendRow([
      'LOG' + Date.now(), data.order_id,
      data.event, data.ip || '', data.device || '', data.source || '',
      new Date().toISOString()
    ]);
  } catch(e) {}
  return { success: true };
}

// ============================================================
// SHOP OWNER REGISTRATION & LOGIN
// ============================================================
function registerShop(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SHOPS');
  const shop_id = 'SHOP' + Date.now().toString().slice(-6);
  sheet.appendRow([
    shop_id, data.shop_name, data.email, data.phone,
    data.plan || 'basic', 0, PLANS[data.plan || 'basic'] ? 10 : 5,
    new Date().toISOString(), 'ACTIVE'
  ]);
  return { success: true, shop_id };
}

function shopLogin(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('SHOPS');
  const rows = sheet.getDataRange().getValues();
  const shop = rows.slice(1).find(r => r[2] === data.email);
  if (!shop) return { success: false, error: 'Shop not found' };
  return { success: true, shop_id: shop[0], shop_name: shop[1], plan: shop[4] };
}

function getShopOrders(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const orders = getOrdersRaw(ss).filter(o => o.EMAIL === data.email || o.SHOP_NAME === data.shop_name);
  return { success: true, orders: orders.map(o => ({
    order_id: o.ORDER_ID, status: o.STATUS, couple: `${o.BRIDE_NAME} & ${o.GROOM_NAME}`,
    wedding_date: o.WEDDING_DATE, published_url: o.PUBLISHED_URL,
    qr_url: o.QR_PREVIEW_URL, views: o.VIEWS, rsvp_count: o.RSVP_COUNT,
    plan: o.PLAN, payment_status: o.PAYMENT_STATUS, amount: o.AMOUNT
  })) };
}

// ============================================================
// HELPERS
// ============================================================
function validateAdmin(token) {
  if (!token) return false;
  const decoded = Utilities.base64Decode(token);
  return decoded && decoded.toString().startsWith(ADMIN_PASSWORD);
}

function getOrderById(order_id) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ORDERS');
  if (!sheet) return null;
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const row = rows.find(r => r[0] === order_id);
  if (!row) return null;
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

function getOrdersRaw(ss) {
  const sheet = ss.getSheetByName('ORDERS');
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function updateOrderField(order_id, field, value) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('ORDERS');
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const col_idx = headers.indexOf(field);
  if (col_idx === -1) return false;
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === order_id) {
      sheet.getRange(i + 1, col_idx + 1).setValue(value);
      return true;
    }
  }
  return false;
}

function buildPublishedEmail(order, url, qr_url) {
  return `
  <div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;background:#1a1208;color:#FAF7F0;padding:40px;">
    <h1 style="color:#C9A84C;font-weight:300;font-size:24px;">🎊 Your Invitation is Live!</h1>
    <p>Dear ${order.CONTACT_NAME},</p>
    <p style="color:#FAF7F099;">The wedding invitation for <strong style="color:#C9A84C;">${order.BRIDE_NAME} & ${order.GROOM_NAME}</strong> is now published and ready to share!</p>
    <div style="text-align:center;margin:30px 0;">
      <a href="${url}" style="background:#C9A84C;color:#1a1208;padding:15px 30px;text-decoration:none;font-family:Arial,sans-serif;font-size:14px;letter-spacing:2px;">VIEW INVITATION →</a>
    </div>
    <p style="color:#FAF7F099;font-size:13px;">Share the link or scan the QR code below with guests via WhatsApp.</p>
    <div style="text-align:center;"><img src="${qr_url}" alt="QR Code" style="width:200px;border:4px solid #C9A84C33;"/></div>
  </div>`;
}

// Run this once to set up all sheets
function onOpen() { setupSheets(); }

// -------------------------
// Test helpers (run from Apps Script editor)
// -------------------------
function testSubmitOrder() {
  const fake = {
    action: 'submit_order',
    shop_name: 'Test Shop', contact_name: 'Tester', phone: '9999999999', email: 'test@example.com',
    bride_name: 'Priya', groom_name: 'Arjun', wedding_date: '2025-12-12', plan: 'basic'
  };
  const out = submitOrder(fake);
  Logger.log('testSubmitOrder -> %s', JSON.stringify(out));
}

function testDoPostForm() {
  const e = {
    postData: { type: 'application/x-www-form-urlencoded', contents: 'action=submit_order&shop_name=FormShop&contact_name=FormUser&phone=9999999999&email=form@example.com&bride_name=Priya&groom_name=Arjun&wedding_date=2025-12-12&plan=basic' },
    parameter: {}
  };
  const out = doPost(e);
  Logger.log('testDoPostForm -> %s', out.getContent ? out.getContent() : JSON.stringify(out));
}

function testDoPostJson() {
  const payload = { action: 'submit_order', shop_name: 'JsonShop', contact_name: 'JsonUser', phone: '9999999999', email: 'json@example.com', bride_name: 'Priya', groom_name: 'Arjun', wedding_date: '2025-12-12', plan: 'basic' };
  const e = { postData: { type: 'application/json', contents: JSON.stringify(payload) }, parameter: {} };
  const out = doPost(e);
  Logger.log('testDoPostJson -> %s', out.getContent ? out.getContent() : JSON.stringify(out));
}

function testHandlePhotoUpload() {
  const fake = { order_id: 'TEST123', photo_type: 'bride', base64_data: 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD', file_name: 'bride.jpg', file_size_kb: 150, mime_type: 'image/jpeg' };
  const out = handlePhotoUpload(fake);
  Logger.log('testHandlePhotoUpload -> %s', JSON.stringify(out));
}

function testAdminLogin() {
  const out = adminLogin({ password: ADMIN_PASSWORD });
  Logger.log('testAdminLogin -> %s', JSON.stringify(out));
}