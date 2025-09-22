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
const META_START = '---META---';
const META_END = '---ENDMETA---';

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
function clamp_(value, min, max) {
  return Math.min(max, Math.max(min, value));
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

function toJstDate_(date) {
  if (!isValidDate_(date)) return null;
  const iso = Utilities.formatDate(date, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
  return new Date(iso);
}

// ===================== メタ情報処理 / 優先度スコア =====================
function sanitizeMetaValue_(value, depth) {
  if (depth > 3) return undefined;
  if (value === null) return null;
  const type = typeof value;
  if (type === 'string') return value;
  if (type === 'number') return isFinite(value) ? value : undefined;
  if (type === 'boolean') return value;
  if (Array.isArray(value)) {
    const arr = [];
    for (let i = 0; i < value.length; i++) {
      const sanitized = sanitizeMetaValue_(value[i], depth + 1);
      if (sanitized !== undefined) arr.push(sanitized);
    }
    return arr;
  }
  if (type === 'object') {
    const obj = {};
    const keys = Object.keys(value);
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      const sanitized = sanitizeMetaValue_(value[key], depth + 1);
      if (sanitized !== undefined) obj[key] = sanitized;
    }
    return obj;
  }
  return undefined;
}
function sanitizeMetaRoot_(meta) {
  if (!meta || typeof meta !== 'object' || Array.isArray(meta)) return null;
  const sanitized = sanitizeMetaValue_(meta, 0);
  if (!sanitized || Array.isArray(sanitized)) return null;
  const keys = Object.keys(sanitized);
  return keys.length ? sanitized : null;
}
function parseDateFlexible_(value) {
  if (!value) return null;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
      return new Date(`${trimmed}T00:00:00+09:00`);
    }
    const d = new Date(trimmed);
    return isValidDate_(d) ? d : null;
  }
  return null;
}
function toNumberOrDefault_(value, defaultValue) {
  if (typeof value === 'number' && isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) {
    const num = Number(value);
    if (!isNaN(num)) return num;
  }
  return defaultValue;
}
function computePriorityScore_(meta, explicitScore) {
  if (typeof explicitScore === 'string' && explicitScore.trim()) {
    const parsed = Number(explicitScore);
    if (!isNaN(parsed)) {
      return clamp_(Math.round(parsed * 100) / 100, 0, 100);
    }
  }
  if (typeof explicitScore === 'number' && isFinite(explicitScore)) {
    return clamp_(Math.round(explicitScore * 100) / 100, 0, 100);
  }
  if (!meta || typeof meta !== 'object') return null;

  const metaWithDefaults = meta;
  let deadlineScore = 0;
  if (metaWithDefaults.deadline) {
    const parsed = parseDateFlexible_(metaWithDefaults.deadline);
    if (isValidDate_(parsed)) {
      const dayStr = Utilities.formatDate(parsed, TZ, 'yyyy-MM-dd');
      const deadlineStart = new Date(`${dayStr}T00:00:00+09:00`);
      const todayStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
      const todayStart = new Date(`${todayStr}T00:00:00+09:00`);
      const diffDays = Math.floor((deadlineStart.getTime() - todayStart.getTime()) / (24 * 60 * 60 * 1000));
      deadlineScore = clamp_((30 - diffDays) * 2, 0, 60);
    }
  }

  const impact = clamp_(toNumberOrDefault_(metaWithDefaults.impact, 3), 1, 5);
  const effort = clamp_(toNumberOrDefault_(metaWithDefaults.effort, 3), 1, 5);
  const must = (typeof metaWithDefaults.must === 'string')
    ? metaWithDefaults.must.toLowerCase() === 'true'
    : metaWithDefaults.must === true;

  const impactScore = impact * 6;
  const effortPenalty = effort * 3;
  const mustBonus = must ? 10 : 0;
  const rawScore = deadlineScore + impactScore - effortPenalty + mustBonus;
  return clamp_(Math.round(rawScore * 100) / 100, 0, 100);
}
function mergePriorityIntoMeta_(meta, priorityScore) {
  const hasMeta = meta && typeof meta === 'object' && !Array.isArray(meta);
  const base = hasMeta ? JSON.parse(JSON.stringify(meta)) : {};
  if (typeof priorityScore === 'number' && isFinite(priorityScore)) {
    base.priorityScore = priorityScore;
  } else if (hasMeta && typeof meta.priorityScore === 'number' && isFinite(meta.priorityScore)) {
    base.priorityScore = clamp_(meta.priorityScore, 0, 100);
  } else if (hasMeta && typeof meta.priorityScore === 'string' && meta.priorityScore.trim()) {
    const parsed = Number(meta.priorityScore);
    if (!isNaN(parsed)) base.priorityScore = clamp_(parsed, 0, 100);
  }
  const keys = Object.keys(base);
  return keys.length ? base : null;
}
function splitDescriptionAndMeta_(description) {
  const text = typeof description === 'string' ? description : '';
  const regex = new RegExp(`${META_START}[\s\S]*?${META_END}`, 'm');
  const match = text.match(regex);
  if (!match) return { body: text, meta: null };

  const metaBlock = match[0];
  const metaJson = metaBlock
    .replace(META_START, '')
    .replace(META_END, '')
    .trim();
  let meta = null;
  if (metaJson) {
    try {
      const parsed = JSON.parse(metaJson);
      meta = sanitizeMetaRoot_(parsed);
    } catch (err) {
      console.warn('meta parse failed:', err && err.message);
    }
  }
  const cleaned = text.replace(metaBlock, '').replace(/\s+$/, '');
  return { body: cleaned.replace(/\n{3,}/g, '\n\n'), meta };
}
function buildDescriptionWithMeta_(description, metaObj) {
  const base = splitDescriptionAndMeta_(typeof description === 'string' ? description : '');
  const trimmed = base.body.replace(/\s+$/, '');
  if (!metaObj) return trimmed;
  const metaJson = JSON.stringify(metaObj);
  const prefix = trimmed ? `${trimmed}\n\n` : '';
  return `${prefix}${META_START}\n${metaJson}\n${META_END}`;
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

function fetchEventById_(calendarId, eventIdOrIcalUid) {
  if (!eventIdOrIcalUid) return null;
  const calId = calendarId || 'primary';
  try {
    return Calendar.Events.get(calId, eventIdOrIcalUid);
  } catch (err) {
    console.warn('fetchEventById_: direct get failed', err && err.message);
  }
  try {
    const res = Calendar.Events.list(calId, {
      iCalUID: eventIdOrIcalUid,
      maxResults: 1,
      singleEvents: true
    });
    if (res && res.items && res.items.length > 0) {
      return res.items[0];
    }
  } catch (err2) {
    console.warn('fetchEventById_: lookup by iCalUID failed', err2 && err2.message);
  }
  return null;
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
  const metaInfo = splitDescriptionAndMeta_(out.description || '');
  if (metaInfo.meta) {
    const priority = computePriorityScore_(metaInfo.meta, metaInfo.meta.priorityScore);
    const mergedMeta = mergePriorityIntoMeta_(metaInfo.meta, priority);
    if (mergedMeta) {
      out.meta = mergedMeta;
      if (typeof mergedMeta.priorityScore === 'number') out.priorityScore = mergedMeta.priorityScore;
    }
  }
  out.descriptionPlain = metaInfo.body || '';
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

// ===================== 補助機能 (Nudge / Reschedule) =====================
function getEventStartDate_(event) {
  if (event && event.start) {
    if (event.start.dateTime) return new Date(event.start.dateTime);
    if (event.start.date) return new Date(`${event.start.date}T00:00:00+09:00`);
  }
  return null;
}

function getEventEndDate_(event) {
  if (event && event.end) {
    if (event.end.dateTime) return new Date(event.end.dateTime);
    if (event.end.date) return new Date(`${event.end.date}T00:00:00+09:00`);
  }
  return null;
}

function isAllDayEvent_(event) {
  return !!(event && event.start && event.start.date && !event.start.dateTime);
}

function isTruthyFlag_(value) {
  if (value === true) return true;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    return normalized === 'true' || normalized === 'done' || normalized === 'completed' || normalized === '完了' || normalized === '済';
  }
  return false;
}

function isEventCompleted_(event) {
  if (!event) return false;
  const meta = event.meta || {};
  if (isTruthyFlag_(meta.completed)) return true;
  if (isTruthyFlag_(meta.done)) return true;
  if (isTruthyFlag_(meta.status)) return true;
  if (meta.progress && typeof meta.progress === 'string') {
    if (isTruthyFlag_(meta.progress)) return true;
  }
  if (typeof event.summary === 'string' && /^\s*(✅|✔|\[done\]|\[完了\]|完了|済)/i.test(event.summary)) return true;
  return false;
}

function determinePreferredStartTime_(priorityScore, originalStart) {
  if (typeof priorityScore === 'number' && isFinite(priorityScore)) {
    if (priorityScore >= 85) return '08:30';
    if (priorityScore >= 70) return '09:30';
    if (priorityScore >= 55) return '11:00';
    if (priorityScore >= 40) return '13:30';
  }
  if (isValidDate_(originalStart)) {
    return Utilities.formatDate(originalStart, TZ, 'HH:mm');
  }
  return '15:00';
}

function generateNudgeForEvent_(normalizedEvent) {
  const summary = (normalizedEvent && normalizedEvent.summary) ? normalizedEvent.summary : 'タスク';
  const meta = (normalizedEvent && normalizedEvent.meta) ? normalizedEvent.meta : {};
  const descriptionPlain = (normalizedEvent && typeof normalizedEvent.descriptionPlain === 'string')
    ? normalizedEvent.descriptionPlain
    : '';
  const lines = descriptionPlain
    .split(/\n+/)
    .map(line => line.trim())
    .filter(line => line.length > 0);

  const firstStepCandidates = [];
  if (meta.project) firstStepCandidates.push(`プロジェクト「${meta.project}」のゴールを確認`);
  if (lines.length) firstStepCandidates.push(`メモ「${lines[0]}」を読み返す`);
  const firstStep = firstStepCandidates.length
    ? firstStepCandidates[0]
    : `まずは「${summary}」の目的を整理しましょう`;

  const checklist = [];
  if (meta.deadline) checklist.push(`締切 ${meta.deadline} をカレンダーで確認`);
  if (Array.isArray(meta.tags) && meta.tags.length) {
    checklist.push(`関連タグ: ${meta.tags.slice(0, 3).join(', ')}`);
  }
  for (let i = 0; i < lines.length && checklist.length < 5; i++) {
    checklist.push(`メモ確認: ${lines[i]}`);
  }
  const checklistDefaults = [
    '必要な資料やリンクを開く',
    '完了条件を箇条書きにする',
    '5分でできる最初のアクションを決める'
  ];
  for (let i = 0; i < checklistDefaults.length; i++) {
    if (checklist.indexOf(checklistDefaults[i]) === -1) {
      checklist.push(checklistDefaults[i]);
    }
  }

  const priorityScore = (normalizedEvent && typeof normalizedEvent.priorityScore === 'number')
    ? normalizedEvent.priorityScore
    : null;
  const ifThen = [];
  if (priorityScore !== null && priorityScore >= 80) {
    ifThen.push('もしブロッカーがあれば→すぐに関係者へ相談');
  } else if (priorityScore !== null && priorityScore >= 60) {
    ifThen.push('もし時間が足りなければ→翌日の午前に再配置');
  } else {
    ifThen.push('進捗が止まったら→15分のフォローアップ枠を確保');
  }
  if (meta.deadline) {
    ifThen.push(`締切 ${meta.deadline} に間に合わない場合→優先度とリスケを再検討`);
  }

  const uniqueChecklist = Array.from(new Set(checklist)).slice(0, 5);
  const uniqueIfThen = Array.from(new Set(ifThen));

  return {
    firstStep,
    checklist: uniqueChecklist,
    ifThen: uniqueIfThen,
    context: {
      summary,
      priorityScore: priorityScore !== null ? priorityScore : undefined,
      start: normalizedEvent && normalizedEvent.start ? normalizedEvent.start : undefined
    }
  };
}

function shouldRescheduleEvent_(event, referenceTime) {
  if (!event) return false;
  if (isEventCompleted_(event)) return false;
  const meta = event.meta || {};
  if (meta.autoReschedule === false) return false;
  if (typeof meta.autoReschedule === 'string' && meta.autoReschedule.trim().toLowerCase() === 'false') {
    return false;
  }
  const start = getEventStartDate_(event);
  if (!isValidDate_(start)) return false;
  if (!isValidDate_(referenceTime)) return false;
  return start.getTime() <= referenceTime.getTime();
}

function pickBusinessWindows_(candidate, fallback) {
  if (Array.isArray(candidate) && candidate.length) {
    const filtered = candidate
      .map(v => (typeof v === 'string' ? v.trim() : ''))
      .filter(v => /^\d{2}:\d{2}-\d{2}:\d{2}$/.test(v));
    if (filtered.length) return filtered;
  }
  return fallback;
}

function reschedulePendingEvents_(options) {
  const calendarId = (options && options.calendarId) || null;
  const referenceRaw = options && options.referenceTime ? new Date(options.referenceTime) : new Date();
  const referenceBase = isValidDate_(referenceRaw) ? referenceRaw : new Date();
  const referenceJst = toJstDate_(referenceBase) || new Date();

  const dayStr = Utilities.formatDate(referenceJst, TZ, 'yyyy-MM-dd');
  const hour = Number(Utilities.formatDate(referenceJst, TZ, 'HH'));
  const minute = Number(Utilities.formatDate(referenceJst, TZ, 'mm'));
  const cutoffReached = hour > 20 || (hour === 20 && minute >= 0);
  if (!cutoffReached) {
    return { checkedAt: rfc3339Jst_(referenceJst), rescheduled: [] };
  }

  const defaultWindows = (options && Array.isArray(options.businessWindows) && options.businessWindows.length)
    ? pickBusinessWindows_(options.businessWindows, ['04:30-06:30', '08:00-19:00'])
    : ['04:30-06:30', '08:00-19:00'];
  const minGapMinutes = (options && typeof options.minGapMinutes === 'number') ? options.minGapMinutes : 15;
  const allowWeekendHoliday = options && options.allowWeekendHoliday === true;

  const list = listEvents_(dayStr, dayStr, calendarId, null);
  const events = (list.items || []).filter(ev => ev && ev.status !== 'cancelled');
  const referenceTime = new Date(`${dayStr}T${Utilities.formatDate(referenceJst, TZ, 'HH:mm:ss')}+09:00`);
  const candidates = events.filter(ev => shouldRescheduleEvent_(ev, referenceTime));
  if (!candidates.length) {
    return { checkedAt: rfc3339Jst_(referenceJst), rescheduled: [] };
  }

  const prepared = candidates.map(ev => {
    let baseMeta = sanitizeMetaRoot_(ev.meta);
    if (!baseMeta) baseMeta = {};
    else baseMeta = JSON.parse(JSON.stringify(baseMeta));
    if (baseMeta && typeof baseMeta.priorityScore !== 'undefined') {
      delete baseMeta.priorityScore;
    }
    const computedPriority = computePriorityScore_(baseMeta, null);
    let priorityValue;
    if (typeof computedPriority === 'number' && isFinite(computedPriority)) {
      priorityValue = computedPriority;
    } else if (typeof ev.priorityScore === 'number' && isFinite(ev.priorityScore)) {
      priorityValue = clamp_(ev.priorityScore, 0, 100);
    } else {
      priorityValue = 50;
    }
    priorityValue = clamp_(Math.round(priorityValue * 100) / 100, 0, 100);
    return { raw: ev, metaBase: baseMeta, priority: priorityValue };
  });

  prepared.sort((a, b) => {
    if (b.priority !== a.priority) return b.priority - a.priority;
    const startA = getEventStartDate_(a.raw);
    const startB = getEventStartDate_(b.raw);
    const timeA = isValidDate_(startA) ? startA.getTime() : 0;
    const timeB = isValidDate_(startB) ? startB.getTime() : 0;
    return timeA - timeB;
  });

  const startDay = new Date(`${dayStr}T00:00:00+09:00`);
  const firstTarget = nextBusinessDay_(startDay, allowWeekendHoliday);
  const startDateStr = fmtDateJst_(firstTarget);
  const calId = calendarId || 'primary';
  const rescheduled = [];
  const errors = [];

  for (let i = 0; i < prepared.length; i++) {
    const entry = prepared[i];
    const event = entry.raw;
    const eventId = event.eventId || event.id;
    if (!eventId) {
      errors.push({ id: event.id || '', message: 'eventId missing' });
      continue;
    }
    try {
      const eventWindows = pickBusinessWindows_(entry.metaBase.businessWindows, defaultWindows);
      const eventAllowWeekend = entry.metaBase.allowWeekendHoliday === true ? true : allowWeekendHoliday;
      const eventMinGap = (typeof entry.metaBase.minGapMinutes === 'number') ? entry.metaBase.minGapMinutes : minGapMinutes;
      const metaForStorage = mergePriorityIntoMeta_(entry.metaBase, entry.priority) || {};
      metaForStorage.lastRescheduledAt = rfc3339Jst_(referenceJst);
      const prevCount = (typeof entry.metaBase.rescheduleCount === 'number' && isFinite(entry.metaBase.rescheduleCount))
        ? entry.metaBase.rescheduleCount
        : 0;
      metaForStorage.rescheduleCount = prevCount + 1;
      const originalStart = getEventStartDate_(event);
      if (isValidDate_(originalStart)) {
        metaForStorage.previousStart = rfc3339Jst_(originalStart);
      }

      if (isAllDayEvent_(event)) {
        let targetDate = new Date(`${startDateStr}T00:00:00+09:00`);
        if (!isBusinessDay_(targetDate, eventAllowWeekend)) {
          targetDate = nextBusinessDay_(targetDate, eventAllowWeekend);
        }
        const targetDateStr = fmtDateJst_(targetDate);
        const nextDate = new Date(targetDate.getTime() + 24 * 60 * 60 * 1000);
        const update = {
          start: { date: targetDateStr },
          end: { date: fmtDateJst_(nextDate) },
          description: buildDescriptionWithMeta_(event.description || '', metaForStorage)
        };
        const patched = Calendar.Events.patch(update, calId, eventId);
        const normalized = normalizeEventItem_(patched);
        rescheduled.push({
          id: normalized.id || event.id || eventId,
          eventId: normalized.eventId || eventId,
          newDate: targetDateStr,
          priorityScore: (typeof normalized.priorityScore === 'number') ? normalized.priorityScore : entry.priority
        });
        continue;
      }

      const startDate = getEventStartDate_(event);
      const endDate = getEventEndDate_(event);
      const durationHours = (isValidDate_(startDate) && isValidDate_(endDate))
        ? Math.max((endDate.getTime() - startDate.getTime()) / (60 * 60 * 1000), 0.5)
        : 1;
      const prefer = determinePreferredStartTime_(entry.priority, startDate);
      const slot = findNextFreeSlotAcrossDays_(calendarId, startDateStr, prefer, durationHours, eventWindows, eventMinGap, eventAllowWeekend);
      const update = {
        start: { dateTime: rfc3339Jst_(slot.start), timeZone: TZ },
        end: { dateTime: rfc3339Jst_(slot.end), timeZone: TZ },
        description: buildDescriptionWithMeta_(event.description || '', metaForStorage)
      };
      const patched = Calendar.Events.patch(update, calId, eventId);
      const normalized = normalizeEventItem_(patched);
      rescheduled.push({
        id: normalized.id || event.id || eventId,
        eventId: normalized.eventId || eventId,
        newDate: fmtDateJst_(slot.start),
        newStart: normalized.start || { dateTime: rfc3339Jst_(slot.start), timeZone: TZ },
        priorityScore: (typeof normalized.priorityScore === 'number') ? normalized.priorityScore : entry.priority
      });
    } catch (err) {
      console.warn('reschedulePendingEvents_: failed', eventId, err && err.message);
      errors.push({ id: event.id || eventId, message: err && err.message });
    }
  }

  const response = { checkedAt: rfc3339Jst_(referenceJst), rescheduled };
  if (prepared.length) response.candidates = prepared.length;
  if (errors.length) response.errors = errors;
  return response;
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
  const desc = ev.getDescription() || '';
  const metaInfo = splitDescriptionAndMeta_(desc);
  const priority = metaInfo.meta ? computePriorityScore_(metaInfo.meta, metaInfo.meta.priorityScore) : null;
  const mergedMeta = metaInfo.meta ? mergePriorityIntoMeta_(metaInfo.meta, priority) : null;
  const result = {
    id: icalUid,
    htmlLink: htmlLink,
    summary: ev.getTitle(),
    description: desc,
    descriptionPlain: metaInfo.body || '',
    start: { dateTime: rfc3339Jst_(ev.getStartTime()), timeZone: TZ },
    end:   { dateTime: rfc3339Jst_(ev.getEndTime()),   timeZone: TZ },
    status: 'confirmed',
    creator: { email: Session.getActiveUser().getEmail() || '' }
  };
  if (mergedMeta) {
    result.meta = mergedMeta;
    if (typeof mergedMeta.priorityScore === 'number') result.priorityScore = mergedMeta.priorityScore;
  }
  return result;
}
function createEvent_CalendarApp_AllDay_(title, description, dateStr, calendarId) {
  const cal = getCalendarByIdOrDefault_(calendarId);
  let d = new Date(`${dateStr}T00:00:00+09:00`);
  const allowWk = false; // allDayは原則ビジネスデー扱い
  if (!isBusinessDay_(d, allowWk)) d = nextBusinessDay_(d, allowWk);

  const ev  = cal.createAllDayEvent(title, d, { description: description || '' });
  const icalUid  = ev.getId();
  const htmlLink = tryGetHtmlLinkByIcalUid_(calendarId, icalUid);
  const desc = ev.getDescription() || '';
  const metaInfo = splitDescriptionAndMeta_(desc);
  const priority = metaInfo.meta ? computePriorityScore_(metaInfo.meta, metaInfo.meta.priorityScore) : null;
  const mergedMeta = metaInfo.meta ? mergePriorityIntoMeta_(metaInfo.meta, priority) : null;
  const result = {
    id: icalUid,
    htmlLink: htmlLink,
    summary: ev.getTitle(),
    description: desc,
    descriptionPlain: metaInfo.body || '',
    start: { date: fmtDateJst_(d) },
    end:   { date: fmtDateJst_(d) },
    status: 'confirmed',
    creator: { email: Session.getActiveUser().getEmail() || '' }
  };
  if (mergedMeta) {
    result.meta = mergedMeta;
    if (typeof mergedMeta.priorityScore === 'number') result.priorityScore = mergedMeta.priorityScore;
  }
  return result;
}
function createEvent_Advanced_Timed_(title, description, start, end, calendarId) {
  const calId = calendarId || 'primary';
  const event = {
    summary: title,
    description: description || '',
    start: { dateTime: rfc3339Jst_(start), timeZone: TZ },
    end:   { dateTime: rfc3339Jst_(end),   timeZone: TZ }
  };
  const inserted = Calendar.Events.insert(event, calId);
  const desc = (inserted && inserted.description) || event.description || '';
  const metaInfo = splitDescriptionAndMeta_(desc);
  const priority = metaInfo.meta ? computePriorityScore_(metaInfo.meta, metaInfo.meta.priorityScore) : null;
  const mergedMeta = metaInfo.meta ? mergePriorityIntoMeta_(metaInfo.meta, priority) : null;
  const result = inserted || {};
  result.description = desc;
  result.descriptionPlain = metaInfo.body || '';
  if (mergedMeta) {
    result.meta = mergedMeta;
    if (typeof mergedMeta.priorityScore === 'number') result.priorityScore = mergedMeta.priorityScore;
  }
  return result;
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
  const inserted = Calendar.Events.insert(event, calId);
  const desc = (inserted && inserted.description) || event.description || '';
  const metaInfo = splitDescriptionAndMeta_(desc);
  const priority = metaInfo.meta ? computePriorityScore_(metaInfo.meta, metaInfo.meta.priorityScore) : null;
  const mergedMeta = metaInfo.meta ? mergePriorityIntoMeta_(metaInfo.meta, priority) : null;
  const result = inserted || {};
  result.description = desc;
  result.descriptionPlain = metaInfo.body || '';
  if (mergedMeta) {
    result.meta = mergedMeta;
    if (typeof mergedMeta.priorityScore === 'number') result.priorityScore = mergedMeta.priorityScore;
  }
  return result;
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

    if (data.action === 'nudge') {
      const eventId = (data.eventId || '').trim();
      if (!eventId) return jsonErr_(400, 'eventId required');
      const event = fetchEventById_(data.calendarId || null, eventId);
      if (!event) return jsonErr_(404, 'event not found', { eventId });
      const normalized = normalizeEventItem_(event);
      const suggestion = generateNudgeForEvent_(normalized);
      return jsonOk_({ ...suggestion, event: normalized });
    }

    if (data.action === 'reschedulePending') {
      const result = reschedulePendingEvents_({
        calendarId: data.calendarId || null,
        referenceTime: data.time || null,
        businessWindows: Array.isArray(data.businessWindows) ? data.businessWindows : null,
        minGapMinutes: (typeof data.minGapMinutes === 'number') ? data.minGapMinutes : null,
        allowWeekendHoliday: data.allowWeekendHoliday === true
      });
      return jsonOk_(result);
    }

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

    const rawDescription = typeof data.description === 'string' ? data.description : '';
    const calendarId  = data.calendarId || null;
    const useAdvanced = data.useAdvanced === true;

    const sanitizedMeta = sanitizeMetaRoot_(data.meta);
    const priorityScore = computePriorityScore_(sanitizedMeta, data.priorityScore);
    const metaForStorage = mergePriorityIntoMeta_(sanitizedMeta, priorityScore);
    const description = buildDescriptionWithMeta_(rawDescription, metaForStorage);

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
      if (metaForStorage) {
        created.meta = created.meta || metaForStorage;
        if (typeof priorityScore === 'number') {
          created.priorityScore = typeof created.priorityScore === 'number' ? created.priorityScore : priorityScore;
          created.meta.priorityScore = created.priorityScore;
        }
      }
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

    if (metaForStorage) {
      created.meta = created.meta || metaForStorage;
      if (typeof priorityScore === 'number') {
        created.priorityScore = typeof created.priorityScore === 'number' ? created.priorityScore : priorityScore;
        created.meta.priorityScore = created.priorityScore;
      }
    }

    return jsonOk_(created);

  } catch (err) {
    console.warn('doPost error:', (err && err.stack) || err);
    return jsonErr_(500, (err && err.message) ? err.message : 'internal error');
  }
}
