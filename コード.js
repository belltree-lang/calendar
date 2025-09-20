/***** Googleカレンダー連携（CalendarApp + Advanced Service 併用） *****
 * デプロイ:
 *  - デプロイ > 新しいデプロイ > 種類: ウェブアプリ
 *  - 実行するユーザー：自分
 *  - アクセスできるユーザー：全員（匿名）
 *
 * 前提:
 *  - 高度なサービス「Calendar」を有効化（スクリプトエディタ > サービス > 追加）
 *
 * 機能（POST /exec）:
 *  - 予定作成（終日/時間指定/自動衝突回避/翌営業日ロール/土日祝は許可制）
 *      { title, description?, date?, startTime?, endTime?, durationHours?, allDay?, calendarId?, useAdvanced?,
 *        autoAvoidConflict?, allowWeekendHoliday?, minGapMinutes?, businessWindows? }
 *  - FreeBusy取得: { action:"freebusy", timeMin, timeMax, calendarId? }
 *  - 一覧取得   : { action:"list",    timeMin, timeMax, calendarId?, pageToken? }
 *
 * 仕様:
 *  - 営業時間（既定）：["04:30-06:30", "08:00-19:00"]
 *  - 土日祝は原則NG（allowWeekendHoliday=true の明示が無ければ翌営業日にロール）
 *  - 同日に空きが無ければ、最大14営業日先まで自動ロール
 *******************************************************************/

const TZ = 'Asia/Tokyo';
const JP_HOLIDAY_CAL_ID = 'ja.japanese#holiday@group.v.calendar.google.com';

// ===================== 共通ユーティリティ =====================
function fmtDateJst_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}
function rfc3339Jst_(d) {
  const s = Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm:ssZ"); // +0900
  return s.replace(/([+-]\d{2})(\d{2})$/, '$1:$2'); // +0900 → +09:00
}
function isValidDate_(d) {
  return d instanceof Date && !isNaN(d.getTime());
}
function jsonOk_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonErr_(code, message, extra) {
  const body = extra ? { error: { code, message, ...extra } } : { error: { code, message } };
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===================== カレンダー取得/URL補助 =====================
function getCalendarByIdOrDefault_(calendarId) {
  if (calendarId) {
    const cal = CalendarApp.getCalendarById(calendarId);
    if (!cal) throw new Error('calendarId が不正、または権限がありません: ' + calendarId);
    return cal;
  }
  return CalendarApp.getDefaultCalendar();
}
function tryGetHtmlLinkByIcalUid_(calendarId, icalUid) {
  try {
    const calId = calendarId || 'primary';
    const res = Calendar.Events.list(calId, {
      iCalUID: icalUid,
      maxResults: 1,
      singleEvents: true
    });
    if (res && res.items && res.items.length > 0) {
      return res.items[0].htmlLink || '';
    }
    return '';
  } catch (e) {
    console.warn('htmlLink fetch by iCalUID failed:', e && e.message);
    return '';
  }
}

// ===================== JP 祝日 / 営業日判定 =====================
function isWeekend_(d) {
  const day = d.getDay(); // 0:日, 6:土
  return day === 0 || day === 6;
}
function isHolidayJP_(d) {
  try {
    const holCal = CalendarApp.getCalendarById(JP_HOLIDAY_CAL_ID);
    if (!holCal) return false;
    const events = holCal.getEventsForDay(d);
    return events && events.length > 0;
  } catch (e) {
    console.warn('holiday check failed:', e && e.message);
    return false; // 取得失敗時は祝日扱いしない
  }
}
function isBusinessDay_(d, allowWeekendHoliday) {
  if (allowWeekendHoliday === true) return true; // 明示許可で週末祝日もOK
  if (isWeekend_(d)) return false;
  if (isHolidayJP_(d)) return false;
  return true;
}
function nextBusinessDay_(d, allowWeekendHoliday) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate()); // 日付のみ
  for (let i = 0; i < 30; i++) {
    x.setDate(x.getDate() + 1);
    if (isBusinessDay_(x, allowWeekendHoliday)) return x;
  }
  return x; // フォールバック
}

// ===================== FreeBusy（Advanced Service） =====================
function getFreeBusy_(timeMinStr, timeMaxStr, calendarId) {
  const calId = calendarId || 'primary';
  const parse = (s, endOfDay) => {
    if (!s) return null;
    if (s.length === 10) {
      return new Date(`${s}T${endOfDay ? '23:59:59' : '00:00:00'}+09:00`);
    }
    return new Date(s);
  };
  const start = parse(timeMinStr, false);
  const end   = parse(timeMaxStr, true);
  if (!isValidDate_(start) || !isValidDate_(end)) {
    throw new Error('invalid timeMin/timeMax');
  }
  const req = { timeMin: rfc3339Jst_(start), timeMax: rfc3339Jst_(end), items: [{ id: calId }] };
  const res = Calendar.Freebusy.query(req);
  const busy = (((res.calendars || {})[calId] || {}).busy) || [];
  return { timeZone: TZ, busy: busy.map(b => ({ start: b.start, end: b.end })) };
}

// ===================== 一覧取得（Advanced Service） =====================
/** ---- Events.list のitemを軽量JSONへ整形 ---- */
function normalizeEventItem_(item) {
  const out = {
    id: item.iCalUID || item.id || '',
    eventId: item.id || '',                 // Advancedの内部ID
    htmlLink: item.htmlLink || '',
    status: item.status || '',
    summary: item.summary || '',
    description: item.description || '',
    location: item.location || '',
    start: item.start || {},                // {date|dateTime,timeZone}
    end: item.end || {},
  };
  if (item.attendees && item.attendees.length) {
    out.attendees = item.attendees.map(a => ({ email: a.email, responseStatus: a.responseStatus }));
  }
  return out;
}
/** ---- 期間内のイベント一覧 ----
 * 入力: { action:"list", timeMin:"YYYY-MM-DD|RFC3339", timeMax:"YYYY-MM-DD|RFC3339", calendarId?, pageToken? }
 * 出力: { items:[{id,eventId,summary,htmlLink,start,end,description?,location?,attendees?}], nextPageToken? }
 */
function listEvents_(timeMinStr, timeMaxStr, calendarId, pageToken) {
  const calId = calendarId || 'primary';
  const parse = (s, endOfDay) => {
    if (s.length === 10) return new Date(`${s}T${endOfDay ? '23:59:59' : '00:00:00'}+09:00`);
    return new Date(s);
  };
  const timeMin = parse(timeMinStr, false);
  const timeMax = parse(timeMaxStr, true);

  const params = {
    timeMin: rfc3339Jst_(timeMin),
    timeMax: rfc3339Jst_(timeMax),
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 250
  };
  if (pageToken) params.pageToken = pageToken;

  const res = Calendar.Events.list(calId, params);
  const items = (res.items || []).map(normalizeEventItem_);
  const out = { items };
  if (res.nextPageToken) out.nextPageToken = res.nextPageToken;
  return out;
}

// ===================== 空き枠探索（1日 / 複数ウィンドウ対応） =====================
/**
 * @param {string|null} calendarId
 * @param {string} dateStr YYYY-MM-DD
 * @param {string} prefer 'HH:mm' 推奨開始（例 '10:00'）
 * @param {number} durationH 必要時間（h）
 * @param {Array<string>} windows 例 ['04:30-06:30','08:00-19:00']
 * @param {number} minGapMin 予定間の最小間隔（分）
 * @returns {{start: Date, end: Date}|null}
 */
function findNextFreeSlotForDay_(calendarId, dateStr, prefer, durationH, windows, minGapMin) {
  const gapMs = (typeof minGapMin === 'number' ? minGapMin : 15) * 60 * 1000;
  const durMs = (durationH > 0 ? durationH : 1) * 60 * 60 * 1000;
  const toDate = hhmm => new Date(`${dateStr}T${hhmm}:00+09:00`);
  const preferDate = prefer ? toDate(prefer) : toDate('10:00');

  const fb = getFreeBusy_(dateStr, dateStr, calendarId || null);
  const busy = fb.busy.map(b => ({ s: new Date(b.start).getTime(), e: new Date(b.end).getTime() }))
                      .sort((a,b) => a.s - b.s);

  function scanWindow(winStart, winEnd, startCursor) {
    const windowStart = Math.max(winStart, startCursor);
    const candidates = busy.concat([{ s: winEnd, e: winEnd }]); // 終端番兵
    let cursor = windowStart;

    for (let i = 0; i < candidates.length; i++) {
      const blockS = Math.max(candidates[i].s, winStart);
      const blockE = Math.min(candidates[i].e, winEnd);
      if (cursor + durMs <= blockS - gapMs) {
        return { start: new Date(cursor), end: new Date(cursor + durMs) };
      }
      if (cursor < blockE + gapMs) cursor = blockE + gapMs;
      if (cursor >= winEnd) break;
    }
    return null;
  }

  for (let idx = 0; idx < windows.length; idx++) {
    const [s, e] = windows[idx].split('-');
    const winS = toDate(s).getTime();
    const winE = toDate(e).getTime();
    const startCursor = (preferDate.getTime() >= winS && preferDate.getTime() < winE) ? preferDate.getTime() : winS;
    const slot = scanWindow(winS, winE, startCursor);
    if (slot) return slot;
  }
  return null;
}

// ===================== 空き枠探索（翌営業日へ自動ロール） =====================
/**
 * @returns {{start: Date, end: Date, dateStr: string}}
 */
function findNextFreeSlotAcrossDays_(calendarId, startDateStr, prefer, durationH, windows, minGapMin, allowWeekendHoliday) {
  let d = new Date(`${startDateStr}T00:00:00+09:00`);
  if (!isBusinessDay_(d, allowWeekendHoliday)) d = nextBusinessDay_(d, allowWeekendHoliday);

  for (let i = 0; i < 14; i++) { // 最大2週間先まで探索
    const dateStr = fmtDateJst_(d);
    if (isBusinessDay_(d, allowWeekendHoliday)) {
      const slot = findNextFreeSlotForDay_(calendarId, dateStr, prefer, durationH, windows, minGapMin);
      if (slot) return { ...slot, dateStr };
    }
    d = nextBusinessDay_(d, allowWeekendHoliday);
  }
  throw new Error('no free slot within 14 business days');
}

// ===================== イベント作成（CalendarApp / Advanced） =====================
function createEvent_CalendarApp_Timed_(title, description, start, end, calendarId) {
  const cal = getCalendarByIdOrDefault_(calendarId);
  const ev  = cal.createEvent(title, start, end, { description: description || '' });
  const icalUid  = ev.getId();
  const htmlLink = tryGetHtmlLinkByIcalUid_(calendarId, icalUid);
  return {
    id: icalUid,
    htmlLink: htmlLink,
    summary: ev.getTitle(),
    description: ev.getDescription() || '',
    start: { dateTime: rfc3339Jst_(ev.getStartTime()), timeZone: TZ },
    end:   { dateTime: rfc3339Jst_(ev.getEndTime()),   timeZone: TZ },
    status: 'confirmed',
    creator: { email: Session.getActiveUser().getEmail() || '' }
  };
}
function createEvent_CalendarApp_AllDay_(title, description, dateStr, calendarId) {
  const cal = getCalendarByIdOrDefault_(calendarId);
  let d = new Date(`${dateStr}T00:00:00+09:00`);
  const allowWk = false; // allDayは原則ビジネスデー扱い
  if (!isBusinessDay_(d, allowWk)) d = nextBusinessDay_(d, allowWk);

  const ev  = cal.createAllDayEvent(title, d, { description: description || '' });
  const icalUid  = ev.getId();
  const htmlLink = tryGetHtmlLinkByIcalUid_(calendarId, icalUid);
  return {
    id: icalUid,
    htmlLink: htmlLink,
    summary: ev.getTitle(),
    description: ev.getDescription() || '',
    start: { date: fmtDateJst_(d) },
    end:   { date: fmtDateJst_(d) },
    status: 'confirmed',
    creator: { email: Session.getActiveUser().getEmail() || '' }
  };
}
function createEvent_Advanced_Timed_(title, description, start, end, calendarId) {
  const calId = calendarId || 'primary';
  const event = {
    summary: title,
    description: description || '',
    start: { dateTime: rfc3339Jst_(start), timeZone: TZ },
    end:   { dateTime: rfc3339Jst_(end),   timeZone: TZ }
  };
  return Calendar.Events.insert(event, calId);
}
function createEvent_Advanced_AllDay_(title, description, dateStr, calendarId) {
  const calId = calendarId || 'primary';
  let d = new Date(`${dateStr}T00:00:00+09:00`);
  const allowWk = false;
  if (!isBusinessDay_(d, allowWk)) d = nextBusinessDay_(d, allowWk);

  const next = new Date(d.getTime() + 24 * 60 * 60 * 1000);
  const event = {
    summary: title,
    description: description || '',
    start: { date: fmtDateJst_(d) },
    end:   { date: fmtDateJst_(next) }
  };
  return Calendar.Events.insert(event, calId);
}

// ===================== Web エンドポイント =====================
function doGet(e) {
  return jsonOk_({ status: 'ok', message: 'ウェブアプリは動作中です' });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      console.warn('doPost: empty payload or not called via HTTP POST');
      return jsonErr_(400, 'empty payload');
    }
    const data = JSON.parse(e.postData.contents);

    // ---- 一覧取得 ----
    if (data.action === 'list') {
      const timeMin = (data.timeMin || '').trim();
      const timeMax = (data.timeMax || '').trim();
      if (!timeMin || !timeMax) return jsonErr_(400, 'timeMin/timeMax required');
      const result = listEvents_(timeMin, timeMax, data.calendarId || null, data.pageToken || null);
      return jsonOk_(result);
    }

    // ---- Freebusy ----
    if (data.action === 'freebusy') {
      const timeMin = (data.timeMin || '').trim();
      const timeMax = (data.timeMax || '').trim();
      if (!timeMin || !timeMax) return jsonErr_(400, 'timeMin/timeMax required');
      const fb = getFreeBusy_(timeMin, timeMax, data.calendarId || null);
      return jsonOk_(fb);
    }

    // ---- 予定作成 ----
    const title = (data.title || '').trim();
    if (!title) return jsonErr_(400, 'title is required');

    const description = data.description || '';
    const calendarId  = data.calendarId || null;
    const useAdvanced = data.useAdvanced === true;

    // 営業時間（上書き可）
    const windows = Array.isArray(data.businessWindows) && data.businessWindows.length
      ? data.businessWindows
      : ['04:30-06:30', '08:00-19:00'];
    const minGapMinutes = (typeof data.minGapMinutes === 'number') ? data.minGapMinutes : 15;

    const allowWeekendHoliday = data.allowWeekendHoliday === true; // 明示許可のみ True

    // 終日イベント
    if (data.allDay === true) {
      const dateStr0 = data.date;
      if (!dateStr0) return jsonErr_(400, 'date is required for allDay');

      let d = new Date(`${dateStr0}T00:00:00+09:00`);
      if (!isBusinessDay_(d, allowWeekendHoliday)) {
        if (!allowWeekendHoliday) d = nextBusinessDay_(d, false);
      }
      const dateStr = fmtDateJst_(d);
      const created = useAdvanced
        ? createEvent_Advanced_AllDay_(title, description, dateStr, calendarId)
        : createEvent_CalendarApp_AllDay_(title, description, dateStr, calendarId);
      return jsonOk_(created);
    }

    // 時間指定イベント
    const duration = (typeof data.durationHours === 'number' && data.durationHours > 0) ? data.durationHours : 1;
    let start, end, dateStrForSlot;

    if (data.date && data.startTime && data.endTime) {
      start = new Date(`${data.date}T${data.startTime}:00+09:00`);
      end   = new Date(`${data.date}T${data.endTime}:00+09:00`);
      dateStrForSlot = data.date;
    } else if (data.date && data.startTime) {
      start = new Date(`${data.date}T${data.startTime}:00+09:00`);
      end   = new Date(start.getTime() + duration * 60 * 60 * 1000);
      dateStrForSlot = data.date;
    } else if (data.date) {
      dateStrForSlot = data.date;
    } else {
      const now = new Date();
      start = new Date(now.getTime() + 5 * 60 * 1000);
      end   = new Date(start.getTime() + duration * 60 * 60 * 1000);
      dateStrForSlot = fmtDateJst_(start);
    }

    const autoAvoid = (data.autoAvoidConflict !== false); // 省略時 true
    if (autoAvoid) {
      const prefer = data.startTime || '10:00';
      const slot = findNextFreeSlotAcrossDays_(
        calendarId,
        dateStrForSlot,
        prefer,
        duration,
        windows,
        minGapMinutes,
        allowWeekendHoliday
      );
      start = slot.start;
      end   = slot.end;
      // slot.dateStr は start から再計算できるため省略
    } else {
      // ガード
      if (!isValidDate_(start) || !isValidDate_(end)) return jsonErr_(400, 'invalid start/end');
      if (end.getTime() <= start.getTime()) end = new Date(start.getTime() + 60 * 60 * 1000);

      // 土日祝チェック（許可なければ翌営業日に日付ごとシフト）
      if (!allowWeekendHoliday) {
        const d0 = new Date(start.getFullYear(), start.getMonth(), start.getDate());
        if (!isBusinessDay_(d0, false)) {
          const nd = nextBusinessDay_(d0, false);
          const delta = nd.getTime() - d0.getTime();
          start = new Date(start.getTime() + delta);
          end   = new Date(end.getTime() + delta);
        }
      }
    }

    const created = useAdvanced
      ? createEvent_Advanced_Timed_(title, description, start, end, calendarId)
      : createEvent_CalendarApp_Timed_(title, description, start, end, calendarId);

    return jsonOk_(created);

  } catch (err) {
    console.warn('doPost error:', (err && err.stack) || err);
    return jsonErr_(500, (err && err.message) ? err.message : 'internal error');
  }
}
