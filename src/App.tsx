import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { FormEvent } from "react";
import { getCalendarDayMetadata } from "./calendarMetadata";

type Account = {
  id: string;
  provider: string;
  provider_user_id: string;
  email: string;
  display_name: string;
  status: string;
  last_synced_at?: string | null;
};

type CalendarItem = {
  id: string;
  source: string;
  source_label: string;
  read_only: boolean;
  title: string;
  start_utc: string;
  end_utc: string;
  start_timezone?: string | null;
  end_timezone?: string | null;
  is_all_day: boolean;
  location?: string | null;
  sensitivity?: string | null;
  category?: string | null;
  note_color?: string | null;
  note_body?: string | null;
  reminder_at_utc?: string | null;
  completed_at?: string | null;
};

type MonthView = {
  year: number;
  month: number;
  items: CalendarItem[];
  accounts: Account[];
};

type ExternalEventDetail = {
  id: string;
  source: string;
  provider_event_id: string;
  title: string;
  body_content_type?: string | null;
  body_content?: string | null;
  start_utc: string;
  end_utc: string;
  start_timezone?: string | null;
  end_timezone?: string | null;
  is_all_day: boolean;
  location?: string | null;
  attendees_json?: string | null;
  organizer_json?: string | null;
  web_link?: string | null;
  online_meeting_url?: string | null;
  categories_json?: string | null;
  reminder_minutes_before_start?: number | null;
  is_reminder_on: boolean;
  sensitivity?: string | null;
  last_modified_utc?: string | null;
};

type Toast = {
  kind: "info" | "error" | "success";
  message: string;
};

type NoteColor = "paper" | "sage" | "rose" | "blue" | "amber";

type CreateNoteDraft = {
  dateKey: string;
  title: string;
  body: string;
  color: NoteColor;
  reminderEnabled: boolean;
  reminderLocal: string;
};

type ReminderNotice = {
  id: string;
  title: string;
  body?: string | null;
  color: string;
  reminder_at_utc: string;
};

const WEEKDAYS = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"];
const NOTE_COLORS: Array<{ value: NoteColor; label: string }> = [
  { value: "paper", label: "纸黄" },
  { value: "sage", label: "鼠尾草" },
  { value: "rose", label: "玫瑰" },
  { value: "blue", label: "浅蓝" },
  { value: "amber", label: "琥珀" },
];

const OUTLOOK_AUTO_SYNC_STORAGE_KEY = "calendar-notes:auto-sync-classic-outlook";
const OUTLOOK_AUTO_SYNC_INTERVAL_MS = 10 * 60 * 1000;

function App() {
  const floatingNoteId = getFloatingNoteIdFromWindow();
  if (floatingNoteId !== null) {
    return <FloatingNoteWindow initialNoteId={floatingNoteId} />;
  }

  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const [monthView, setMonthView] = useState<MonthView>({ year, month, items: [], accounts: [] });
  const [selectedDate, setSelectedDate] = useState(toDateKey(now));
  const [selectedEventId, setSelectedEventId] = useState<string | null>(null);
  const [eventDetail, setEventDetail] = useState<ExternalEventDetail | null>(null);
  const [noteDraft, setNoteDraft] = useState<CreateNoteDraft | null>(null);
  const [noteError, setNoteError] = useState<string | null>(null);
  const [savingNote, setSavingNote] = useState(false);
  const [outlookRefreshing, setOutlookRefreshing] = useState(false);
  const [autoSyncOutlook, setAutoSyncOutlook] = useState(readAutoSyncOutlookPreference);
  const [toast, setToast] = useState<Toast | null>(null);
  const [detailCollapsed, setDetailCollapsed] = useState(false);
  const [reminderNotice, setReminderNotice] = useState<ReminderNotice | null>(null);
  const noteDialogRef = useRef<HTMLDialogElement>(null);
  const reminderDialogRef = useRef<HTMLDialogElement>(null);
  const noteTitleInputRef = useRef<HTMLInputElement>(null);
  const lastFocusedElementRef = useRef<HTMLElement | null>(null);
  const outlookRefreshInFlightRef = useRef(false);
  const refreshOutlookEventsRef = useRef<((manual?: boolean) => Promise<void>) | null>(null);

  const loadMonth = useCallback(async () => {
    if (!isTauriRuntime()) {
      setMonthView({ year, month, items: [], accounts: [] });
      return;
    }
    try {
      const result = await invoke<MonthView>("get_month_view", { year, month });
      setMonthView(result);
    } catch (error) {
      setToast({ kind: "error", message: stringifyError(error) });
    }
  }, [year, month]);

  useEffect(() => {
    void loadMonth();
  }, [loadMonth]);

  const itemsByDate = useMemo(() => groupItemsByDate(monthView.items), [monthView.items]);
  const selectedItems = itemsByDate.get(selectedDate) ?? [];
  const calendarDays = useMemo(() => buildCalendarDays(year, month), [year, month]);
  const selectedItem = useMemo(
    () => monthView.items.find((item) => item.id === selectedEventId) ?? null,
    [monthView.items, selectedEventId],
  );
  const refreshOutlookEvents = useCallback(async (manual = false) => {
    if (outlookRefreshInFlightRef.current) return;
    outlookRefreshInFlightRef.current = true;
    setOutlookRefreshing(true);
    if (manual) {
      setToast({ kind: "info", message: "正在刷新日程..." });
    }
    try {
      ensureTauriRuntime();
      await invoke("refresh_local_outlook");
      await loadMonth();
      if (manual) {
      setToast({ kind: "success", message: "日程已刷新" });
      }
    } catch (error) {
      setToast({ kind: "error", message: stringifyError(error) });
    } finally {
      outlookRefreshInFlightRef.current = false;
      setOutlookRefreshing(false);
    }
  }, [loadMonth]);

  useEffect(() => {
    refreshOutlookEventsRef.current = refreshOutlookEvents;
  }, [refreshOutlookEvents]);

  useEffect(() => {
    if (!selectedEventId || selectedItem?.source !== "outlook") {
      setEventDetail(null);
      return;
    }
    if (!isTauriRuntime()) {
      setEventDetail(null);
      return;
    }
    invoke<ExternalEventDetail | null>("get_external_event_detail", { eventId: selectedEventId })
      .then(setEventDetail)
      .catch((error) => setToast({ kind: "error", message: stringifyError(error) }));
  }, [selectedEventId, selectedItem?.source]);

  useEffect(() => {
    const dialog = noteDialogRef.current;
    if (!dialog) return;

    if (noteDraft && !dialog.open) {
      dialog.showModal();
      window.requestAnimationFrame(() => noteTitleInputRef.current?.focus());
    }

    if (!noteDraft && dialog.open) {
      dialog.close();
      lastFocusedElementRef.current?.focus();
    }
  }, [noteDraft]);

  useEffect(() => {
    const dialog = reminderDialogRef.current;
    if (!dialog) return;

    if (reminderNotice && !dialog.open) {
      dialog.showModal();
    }

    if (!reminderNotice && dialog.open) {
      dialog.close();
    }
  }, [reminderNotice]);

  useEffect(() => {
    if (!isTauriRuntime()) return;

    let unlisten: (() => void) | undefined;
    void listen<ReminderNotice>("todo-reminder", (event) => {
      setReminderNotice(event.payload);
      setSelectedEventId(event.payload.id);
      void loadMonth();
    })
      .then((listener) => {
        unlisten = listener;
      })
      .catch((error) => setToast({ kind: "error", message: stringifyError(error) }));

    return () => unlisten?.();
  }, [loadMonth]);

  useEffect(() => {
    window.localStorage.setItem(OUTLOOK_AUTO_SYNC_STORAGE_KEY, autoSyncOutlook ? "true" : "false");
    if (!autoSyncOutlook) return;

    void refreshOutlookEventsRef.current?.(false);
    const intervalId = window.setInterval(() => {
      void refreshOutlookEventsRef.current?.(false);
    }, OUTLOOK_AUTO_SYNC_INTERVAL_MS);

    return () => window.clearInterval(intervalId);
  }, [autoSyncOutlook]);

  function jumpMonth(offset: number) {
    const date = new Date(year, month - 1 + offset, 1);
    setYear(date.getFullYear());
    setMonth(date.getMonth() + 1);
    setSelectedDate(toDateKey(date));
    setSelectedEventId(null);
  }

  function jumpToday() {
    const today = new Date();
    setYear(today.getFullYear());
    setMonth(today.getMonth() + 1);
    setSelectedDate(toDateKey(today));
  }

  function openNoteDialog(dateKey: string) {
    lastFocusedElementRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    setSelectedDate(dateKey);
    setNoteError(null);
    setNoteDraft({ dateKey, title: "", body: "", color: "paper", reminderEnabled: false, reminderLocal: defaultReminderLocal() });
  }

  function closeNoteDialog() {
    if (savingNote) return;
    setNoteDraft(null);
    setNoteError(null);
  }

  async function submitNote(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!noteDraft) return;

    const title = noteDraft.title.trim();
    if (!title) {
      setNoteError("请输入待办标题");
      noteTitleInputRef.current?.focus();
      return;
    }

    const reminderAtUtc = noteDraft.reminderEnabled ? parseReminderLocal(noteDraft.reminderLocal) : null;
    if (reminderAtUtc === "invalid") {
      setNoteError("提醒时间格式无效");
      return;
    }

    setSavingNote(true);
    setNoteError(null);
    try {
      ensureTauriRuntime();
      const created = await invoke<CalendarItem>("create_note", {
        request: {
          date_key: noteDraft.dateKey,
          title,
          body: noteDraft.body,
          color: noteDraft.color,
          reminder_at_utc: reminderAtUtc,
        },
      });
      await loadMonth();
      setSelectedDate(noteDraft.dateKey);
      setSelectedEventId(created.id);
      setNoteDraft(null);
      setToast({ kind: "success", message: "待办已创建" });
    } catch (error) {
      setNoteError(stringifyError(error));
    } finally {
      setSavingNote(false);
    }
  }

  async function setNoteCompleted(noteId: string, completed: boolean) {
    try {
      ensureTauriRuntime();
      const updated = await invoke<CalendarItem>("set_note_completed", {
        request: { note_id: noteId, completed },
      });
      await loadMonth();
      setSelectedEventId(updated.id);
      if (completed && reminderNotice?.id === noteId) {
        setReminderNotice(null);
      }
      setToast({ kind: "success", message: completed ? "已标记完成" : "已恢复为未完成" });
    } catch (error) {
      setToast({ kind: "error", message: stringifyError(error) });
    }
  }

  async function openFloatingNote(noteId: string) {
    if (!isTauriRuntime()) return;
    try {
      await invoke("open_floating_note", { noteId });
    } catch (error) {
      setToast({ kind: "error", message: stringifyError(error) });
    }
  }

  return (
    <main className="app-shell">
      <section className="stage" aria-label="Calendar Notes 自动同步月历">
        <header className="topbar">
          <div className="brand">
            <h1>Calendar Notes</h1>
          </div>

          <div className="top-actions">
            <button className="button" type="button" onClick={jumpToday}>今天</button>
            <label className="sync-toggle">
              <input
                type="checkbox"
                checked={autoSyncOutlook}
                onChange={(event) => setAutoSyncOutlook(event.target.checked)}
              />
              <span>自动同步 Outlook</span>
              <em>每 10 分钟</em>
            </label>
            <button className="button primary" type="button" onClick={() => refreshOutlookEvents(true)} disabled={outlookRefreshing}>
              {outlookRefreshing ? "刷新中..." : "刷新 Outlook 日程"}
            </button>
          </div>
        </header>

        {toast && (
          <div className={`toast ${toast.kind}`} role="status">
            <span>{toast.message}</span>
            <button type="button" onClick={() => setToast(null)} aria-label="关闭提示">×</button>
          </div>
        )}

        <div className={`layout ${detailCollapsed ? "detail-collapsed" : ""}`}>
          <section className="calendar-panel" aria-label={`${year} 年 ${month} 月大月历`}>
            <div className="calendar-head">
              <div>
                <h2>{year} 年 {month} 月</h2>
                <p>双击日期格新建待办；格内内容按开始时间升序展示。</p>
              </div>
              <div className="month-actions">
                <button className="button" type="button" onClick={() => jumpMonth(-1)}>上一月</button>
                <button className="button primary" type="button" onClick={jumpToday}>本月</button>
                <button className="button" type="button" onClick={() => jumpMonth(1)}>下一月</button>
              </div>
            </div>

            <div className="month-grid" role="grid">
              {WEEKDAYS.map((weekday) => <span className="weekday" key={weekday}>{weekday}</span>)}
              {calendarDays.map((day) => {
                const dayItems = itemsByDate.get(day.key) ?? [];
                const visibleItems = dayItems.slice(0, 4);
                return (
                  <button
                    className={`day-cell ${day.inMonth ? "" : "outside"} ${day.isToday ? "today" : ""} ${day.key === selectedDate ? "selected" : ""} ${day.metadata.holidayKind === "holiday" ? "holiday" : ""} ${day.metadata.holidayKind === "workday" ? "workday" : ""}`}
                    key={day.key}
                    type="button"
                    onClick={() => {
                      setSelectedDate(day.key);
                      setSelectedEventId(dayItems[0]?.id ?? null);
                    }}
                    onDoubleClick={(event) => {
                      event.preventDefault();
                      openNoteDialog(day.key);
                    }}
                    aria-label={`${day.label}，${dayItems.length} 项内容，双击创建待办`}
                    title="双击创建待办"
                  >
                    <span className="date-row">
                      <strong>{day.date.getDate()}</strong>
                      <span className="day-badges">
                        {day.metadata.holidayKind === "holiday" && <em className="holiday-badge">{"\u4f11"}</em>}
                        {day.metadata.holidayKind === "workday" && <em className="workday-badge">{"\u73ed"}</em>}
                        {dayItems.length > 0 && <em>{dayItems.length}项</em>}
                      </span>
                    </span>
                    <span className="lunar-row">
                      <span>{day.metadata.holidayName || day.metadata.lunarFestival || day.metadata.solarFestival || day.metadata.lunarLabel}</span>
                    </span>
                    <span className="tasks">
                      {visibleItems.map((item) => (
                        <span className={taskPillClassName(item)} key={item.id}>
                          <time>{formatItemTimeLabel(item)}</time>
                          <span>{item.title}</span>
                        </span>
                      ))}
                      {dayItems.length > visibleItems.length && <span className="more">+{dayItems.length - visibleItems.length} 项</span>}
                    </span>
                  </button>
                );
              })}
            </div>
          </section>

          <aside className={`detail-panel ${detailCollapsed ? "collapsed" : ""}`} aria-label="选中日期详情">
            <div className="panel-chrome detail-chrome">
              <span className="panel-label">详情</span>
              <button
                className="panel-toggle"
                type="button"
                aria-expanded={!detailCollapsed}
                aria-controls="right-detail-content"
                onClick={() => setDetailCollapsed((collapsed) => !collapsed)}
              >
                {detailCollapsed ? "展开" : "收起"}
              </button>
            </div>
            {detailCollapsed && <div className="collapsed-rail">详情</div>}
            <div className="panel-content" id="right-detail-content" hidden={detailCollapsed}>
            <div className="detail-head">
              <div>
                <h2>{formatDateLabel(selectedDate)}</h2>
                <p>{selectedItems.length} 项内容</p>
              </div>
              <div className="detail-actions">
                <button className="button primary" type="button" onClick={() => openNoteDialog(selectedDate)}>新建待办</button>
                {selectedItem?.source === "local_note" && (
                  <button className="button" type="button" onClick={() => openFloatingNote(selectedItem.id)}>
                    悬浮便签
                  </button>
                )}
                <button className="button" type="button" onClick={() => setSelectedEventId(selectedItems[0]?.id ?? null)} disabled={selectedItems.length === 0}>查看首项</button>
              </div>
            </div>
            <div className="timeline">
              {selectedItems.length === 0 ? (
                <div className="empty-state">当天没有内容。双击日期格或点击“新建待办”开始记录。</div>
              ) : selectedItems.map((item) => (
                <button
                  className={timelineItemClassName(item, item.id === selectedEventId)}
                  key={item.id}
                  type="button"
                  onClick={() => setSelectedEventId(item.id)}
                >
                  <time>{formatItemTimeLabel(item)}</time>
                  <span><strong>{item.title}</strong><em>{formatItemSubtitle(item)}</em></span>
                </button>
              ))}
            </div>

            {selectedItem?.source === "local_note" && (
              <section className={`event-detail note-detail note-${normalizeNoteColor(selectedItem.note_color)} ${selectedItem.completed_at ? "completed" : ""}`} aria-label="本地待办详情">
                <p className="readonly">待办 · 本地</p>
                <h3>{selectedItem.title}</h3>
                <dl>
                  <div><dt>日期</dt><dd>{formatDateLabel(toDateKey(new Date(selectedItem.start_utc)))}</dd></div>
                  <div><dt>颜色</dt><dd>{getNoteColorLabel(selectedItem.note_color)}</dd></div>
                  {selectedItem.reminder_at_utc && <div><dt>提醒</dt><dd>{formatDateTime(selectedItem.reminder_at_utc)}</dd></div>}
                  <div><dt>状态</dt><dd>{selectedItem.completed_at ? `已完成 · ${formatDateTime(selectedItem.completed_at)}` : "未完成"}</dd></div>
                </dl>
                <div className="detail-actions note-status-actions">
                  <button
                    className={`button ${selectedItem.completed_at ? "" : "primary"}`}
                    type="button"
                    onClick={() => setNoteCompleted(selectedItem.id, !selectedItem.completed_at)}
                  >
                    {selectedItem.completed_at ? "恢复未完成" : "标记完成"}
                  </button>
                </div>
                {selectedItem.note_body ? (
                  <div className="body-preview plain">{selectedItem.note_body}</div>
                ) : (
                  <div className="empty-state compact">这条待办还没有正文。</div>
                )}
              </section>
            )}

            {selectedItem?.source === "outlook" && eventDetail && (
              <section className="event-detail" aria-label="只读日程详情">
                <p className="readonly">外部日程 · 只读</p>
                <h3>{eventDetail.title}</h3>
                <dl>
                  <div><dt>时间</dt><dd>{formatDateTime(eventDetail.start_utc)} - {formatDateTime(eventDetail.end_utc)}</dd></div>
                  {eventDetail.location && <div><dt>地点</dt><dd>{eventDetail.location}</dd></div>}
                  {eventDetail.web_link && <div><dt>链接</dt><dd><a href={eventDetail.web_link} target="_blank" rel="noreferrer">打开原始日程</a></dd></div>}
                  {eventDetail.online_meeting_url && <div><dt>线上会议</dt><dd><a href={eventDetail.online_meeting_url} target="_blank" rel="noreferrer">打开会议链接</a></dd></div>}
                </dl>
                {eventDetail.body_content && (
                  eventDetail.body_content_type === "html" ? (
                    <div className="body-preview" dangerouslySetInnerHTML={{ __html: eventDetail.body_content }} />
                  ) : (
                    <div className="body-preview plain">{eventDetail.body_content}</div>
                  )
                )}
              </section>
            )}
            </div>
          </aside>
        </div>

        <dialog
          className="note-dialog"
          ref={noteDialogRef}
          aria-labelledby="note-dialog-title"
          onCancel={(event) => {
            event.preventDefault();
            closeNoteDialog();
          }}
        >
          {noteDraft && (
            <form className="note-card" onSubmit={submitNote} noValidate>
              <div className="note-dialog-head">
                <p className="readonly">本地待办</p>
                <h3 id="note-dialog-title">新建待办</h3>
                <span>创建到 {formatDateLabel(noteDraft.dateKey)}</span>
              </div>

              <label className="field" htmlFor="note-title">
                <span>标题</span>
                <input
                  id="note-title"
                  ref={noteTitleInputRef}
                  value={noteDraft.title}
                  onChange={(event) => setNoteDraft((draft) => draft ? { ...draft, title: event.target.value } : draft)}
                  aria-invalid={Boolean(noteError)}
                  aria-describedby={noteError ? "note-title-error" : undefined}
                  required
                  placeholder="例如：跟进需求、买咖啡豆"
                />
              </label>

              <label className="field" htmlFor="note-body">
                <span>正文</span>
                <textarea
                  id="note-body"
                  rows={4}
                  value={noteDraft.body}
                  onChange={(event) => setNoteDraft((draft) => draft ? { ...draft, body: event.target.value } : draft)}
                  placeholder="补充说明、灵感或待办细节"
                />
              </label>

              <div className={`reminder-card ${noteDraft.reminderEnabled ? "enabled" : ""}`}>
                <div className="reminder-card-head">
                  <div>
                    <p className="reminder-kicker">提醒设置</p>
                    <strong>到点未完成时提醒</strong>
                    <span>默认不启用；修改时间后自动开启。</span>
                  </div>
                  <label className="switch-field" htmlFor="note-reminder-enabled">
                    <input
                      id="note-reminder-enabled"
                      type="checkbox"
                      checked={noteDraft.reminderEnabled}
                      onChange={(event) => setNoteDraft((draft) => draft ? { ...draft, reminderEnabled: event.target.checked } : draft)}
                    />
                    <span className="switch-track" aria-hidden="true"><span /></span>
                    <span className="switch-text">{noteDraft.reminderEnabled ? "已启用" : "未启用"}</span>
                  </label>
                </div>
                <label className="reminder-time-field" htmlFor="note-reminder">
                  <span>提醒时间</span>
                  <input
                    id="note-reminder"
                    type="datetime-local"
                    value={noteDraft.reminderLocal}
                    onChange={(event) => setNoteDraft((draft) => draft ? { ...draft, reminderEnabled: true, reminderLocal: event.target.value } : draft)}
                    aria-describedby="note-reminder-help"
                  />
                </label>
                <p id="note-reminder-help">默认预设为当前时间后 2 小时；只有启用后才会保存提醒。</p>
              </div>

              <fieldset className="color-field">
                <legend>颜色</legend>
                <div className="color-options">
                  {NOTE_COLORS.map((option) => (
                    <label className={`color-option note-${option.value} ${noteDraft.color === option.value ? "selected" : ""}`} key={option.value}>
                      <input
                        type="radio"
                        name="note-color"
                        value={option.value}
                        checked={noteDraft.color === option.value}
                        onChange={() => setNoteDraft((draft) => draft ? { ...draft, color: option.value } : draft)}
                      />
                      <span className="color-dot" aria-hidden="true" />
                      <span>{option.label}</span>
                    </label>
                  ))}
                </div>
              </fieldset>

              {noteError && <p className="form-error" id="note-title-error" role="alert">{noteError}</p>}

              <div className="dialog-actions">
                <button className="button" type="button" onClick={closeNoteDialog} disabled={savingNote}>取消</button>
                <button className="button primary" type="submit" disabled={savingNote}>{savingNote ? "保存中" : "创建待办"}</button>
              </div>
            </form>
          )}
        </dialog>

        <dialog
          className="note-dialog reminder-dialog"
          ref={reminderDialogRef}
          aria-labelledby="reminder-dialog-title"
          onCancel={(event) => {
            event.preventDefault();
            setReminderNotice(null);
          }}
        >
          {reminderNotice && (
            <section className={`note-card note-${normalizeNoteColor(reminderNotice.color)}`}>
              <div className="note-dialog-head">
                <p className="readonly">待办提醒</p>
                <h3 id="reminder-dialog-title">{reminderNotice.title}</h3>
                <span>提醒时间：{formatDateTime(reminderNotice.reminder_at_utc)}</span>
              </div>
              {reminderNotice.body ? (
                <div className="body-preview plain">{reminderNotice.body}</div>
              ) : (
                <div className="empty-state compact">这条待办还没有正文。</div>
              )}
              <div className="dialog-actions">
                <button className="button" type="button" onClick={() => setReminderNotice(null)}>稍后处理</button>
                <button className="button primary" type="button" onClick={() => setNoteCompleted(reminderNotice.id, true)}>标记完成</button>
              </div>
            </section>
          )}
        </dialog>
      </section>
    </main>
  );
}

function FloatingNoteWindow({ initialNoteId }: { initialNoteId: string | null }) {
  const [highlightNoteId, setHighlightNoteId] = useState<string | null>(initialNoteId);
  const [notes, setNotes] = useState<CalendarItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [alwaysOnTop, setAlwaysOnTop] = useState(true);

  const loadUpcomingNotes = useCallback(async () => {
    if (!isTauriRuntime()) return;
    try {
      setLoading(true);
      setError(null);
      const result = await invoke<CalendarItem[]>("get_upcoming_local_notes");
      setNotes(result);
    } catch (loadError) {
      setError(stringifyError(loadError));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void loadUpcomingNotes();
  }, [loadUpcomingNotes]);

  useEffect(() => {
    if (!isTauriRuntime()) return;

    void invoke<string | null>("get_active_floating_note_id")
      .then((activeNoteId) => {
        if (activeNoteId) setHighlightNoteId(activeNoteId);
      })
      .catch((activeError) => {
        setError(stringifyError(activeError));
        setLoading(false);
      });

    let disposed = false;
    let unlisten: (() => void) | null = null;
    void listen<string>("floating-note-open", (event) => {
      if (!disposed) {
        setHighlightNoteId(event.payload);
        void loadUpcomingNotes();
      }
    }).then((cleanup) => {
      if (disposed) cleanup();
      else unlisten = cleanup;
    });

    return () => {
      disposed = true;
      unlisten?.();
    };
  }, []);

  async function toggleAlwaysOnTop() {
    const nextValue = !alwaysOnTop;
    try {
      await invoke("set_floating_note_always_on_top", { alwaysOnTop: nextValue });
      setAlwaysOnTop(nextValue);
    } catch (toggleError) {
      setError(stringifyError(toggleError));
    }
  }

  async function closeWindow() {
    await getCurrentWindow().close();
  }

  async function completeNote(noteId: string) {
    try {
      await invoke<CalendarItem>("set_note_completed", {
        request: { note_id: noteId, completed: true },
      });
      setNotes((currentNotes) => currentNotes.filter((item) => item.id !== noteId));
    } catch (completeError) {
      setError(stringifyError(completeError));
    }
  }

  const groupedNotes = useMemo(() => groupItemsByDate(notes), [notes]);
  const upcomingGroups = Array.from(groupedNotes.entries());
  return (
    <main className="floating-note-shell note-paper">
      <header className="floating-note-grip" data-tauri-drag-region>
        <div data-tauri-drag-region>
          <span className="floating-note-kicker" data-tauri-drag-region>桌面待办</span>
          <strong data-tauri-drag-region>{notes.length} 条后续待办</strong>
        </div>
        <div className="floating-note-controls">
          <button type="button" onClick={loadUpcomingNotes}>刷新</button>
          <button type="button" onClick={toggleAlwaysOnTop} aria-pressed={alwaysOnTop}>
            {alwaysOnTop ? "取消置顶" : "置顶"}
          </button>
          <button type="button" onClick={closeWindow} aria-label="关闭悬浮便签">×</button>
        </div>
      </header>

      <section className="floating-note-card" aria-label="悬浮便签内容">
        <i className="floating-note-tape" aria-hidden="true" />
        {loading ? (
          <div className="floating-note-empty">正在加载后续待办...</div>
        ) : error ? (
          <div className="floating-note-empty error">{error}</div>
        ) : upcomingGroups.length > 0 ? (
          <>
            <div className="floating-note-title-row">
              <p>今天及之后</p>
              <span>按日期分组</span>
            </div>
            <div className="floating-note-list" aria-label="后续待办列表">
              {upcomingGroups.map(([dateKey, dayNotes]) => (
                <section className="floating-note-day" key={dateKey} aria-label={formatDateLabel(dateKey)}>
                  <header>
                    <strong>{formatFloatingDateLabel(dateKey)}</strong>
                    <span>{dayNotes.length} 项</span>
                  </header>
                  {dayNotes.map((item) => (
                    <article
                      className={`floating-note-item note-${normalizeNoteColor(item.note_color)} ${item.id === highlightNoteId ? "highlight" : ""}`}
                      key={item.id}
                    >
                      <div>
                        <time>{formatItemTimeLabel(item)}</time>
                        <span>{getNoteColorLabel(item.note_color)}</span>
                      </div>
                      <h2>{item.title}</h2>
                      {item.note_body?.trim() && <p>{item.note_body.trim()}</p>}
                      <footer>
                        {item.reminder_at_utc ? <span>提醒 {formatDateTime(item.reminder_at_utc)}</span> : <span>无提醒</span>}
                        <button type="button" onClick={() => completeNote(item.id)}>完成</button>
                      </footer>
                    </article>
                  ))}
                </section>
              ))}
            </div>
          </>
        ) : (
          <div className="floating-note-empty">今天之后没有未完成待办。</div>
        )}
      </section>
    </main>
  );
}

function getFloatingNoteIdFromWindow() {
  const queryNoteId = new URLSearchParams(window.location.search).get("floatingNoteId");
  if (queryNoteId) return queryNoteId;
  if (!isTauriRuntime()) return null;
  const label = getCurrentWindow().label;
  if (label === "floating-note") return "";
  const prefix = "floating-note-";
  return label.startsWith(prefix) ? label.slice(prefix.length) : null;
}

function buildCalendarDays(year: number, month: number) {
  const first = new Date(year, month - 1, 1);
  const mondayOffset = (first.getDay() + 6) % 7;
  const start = new Date(year, month - 1, 1 - mondayOffset);
  return Array.from({ length: 42 }, (_, index) => {
    const date = new Date(start);
    date.setDate(start.getDate() + index);
    const key = toDateKey(date);
    return {
      key,
      date,
      inMonth: date.getMonth() === month - 1,
      isToday: key === toDateKey(new Date()),
      label: `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`,
      metadata: getCalendarDayMetadata(date),
    };
  });
}

function groupItemsByDate(items: CalendarItem[]) {
  const grouped = new Map<string, CalendarItem[]>();
  for (const item of items) {
    const key = toDateKey(new Date(item.start_utc));
    const dayItems = grouped.get(key) ?? [];
    dayItems.push(item);
    grouped.set(key, dayItems);
  }
  for (const dayItems of grouped.values()) {
    dayItems.sort((left, right) => {
      const timeDiff = new Date(left.start_utc).getTime() - new Date(right.start_utc).getTime();
      if (timeDiff !== 0) return timeDiff;
      return left.title.localeCompare(right.title, "zh-CN");
    });
  }
  return grouped;
}

function toDateKey(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatTime(value: string) {
  return new Intl.DateTimeFormat("zh-CN", { hour: "2-digit", minute: "2-digit", hour12: false }).format(new Date(value));
}

function formatDateTime(value: string) {
  return new Intl.DateTimeFormat("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false }).format(new Date(value));
}

function formatDateLabel(key: string) {
  const date = new Date(`${key}T00:00:00`);
  return new Intl.DateTimeFormat("zh-CN", { month: "long", day: "numeric", weekday: "long" }).format(date);
}

function formatFloatingDateLabel(key: string) {
  const todayKey = toDateKey(new Date());
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowKey = toDateKey(tomorrow);
  if (key === todayKey) return "今天";
  if (key === tomorrowKey) return "明天";
  return formatDateLabel(key);
}

function formatItemTimeLabel(item: CalendarItem) {
  if (item.source === "local_note") return item.completed_at ? "完成" : "待办";
  return item.is_all_day ? "全天" : formatTime(item.start_utc);
}

function formatItemSubtitle(item: CalendarItem) {
  if (item.source === "local_note") {
    if (item.completed_at) return `已完成 · ${formatDateTime(item.completed_at)}`;
    if (item.reminder_at_utc) return `提醒 · ${formatDateTime(item.reminder_at_utc)}`;
    return item.note_body?.trim() || `本地待办 · ${getNoteColorLabel(item.note_color)}`;
  }
  return item.location || "只读日程";
}

function taskPillClassName(item: CalendarItem) {
  const classes = ["task-pill"];
  if (item.source === "local_note") classes.push("local-note", `note-${normalizeNoteColor(item.note_color)}`);
  if (item.completed_at) classes.push("completed");
  if (item.reminder_at_utc && !item.completed_at) classes.push("has-reminder");
  if (item.sensitivity === "private") classes.push("private");
  if (item.is_all_day && item.source !== "local_note") classes.push("all-day");
  return classes.join(" ");
}

function timelineItemClassName(item: CalendarItem, active: boolean) {
  const classes = ["timeline-item"];
  if (active) classes.push("active");
  if (item.source === "local_note") classes.push("local-note", `note-${normalizeNoteColor(item.note_color)}`);
  if (item.completed_at) classes.push("completed");
  if (item.reminder_at_utc && !item.completed_at) classes.push("has-reminder");
  return classes.join(" ");
}
function normalizeNoteColor(value?: string | null): NoteColor {
  if (value === "sage" || value === "rose" || value === "blue" || value === "amber") return value;
  return "paper";
}

function getNoteColorLabel(value?: string | null) {
  const color = normalizeNoteColor(value);
  return NOTE_COLORS.find((option) => option.value === color)?.label ?? "纸黄";
}

function defaultReminderLocal() {
  const date = new Date();
  date.setHours(date.getHours() + 2);
  return toDatetimeLocalValue(date);
}

function toDatetimeLocalValue(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hour = String(date.getHours()).padStart(2, "0");
  const minute = String(date.getMinutes()).padStart(2, "0");
  return `${year}-${month}-${day}T${hour}:${minute}`;
}

function parseReminderLocal(value: string) {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "invalid";
  return date.toISOString();
}

function readAutoSyncOutlookPreference() {
  if (typeof window === "undefined") return false;
  return window.localStorage.getItem(OUTLOOK_AUTO_SYNC_STORAGE_KEY) === "true";
}

function stringifyError(error: unknown) {
  if (typeof error === "string") return error;
  if (error instanceof Error) return error.message;
  return JSON.stringify(error);
}

function isTauriRuntime() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

function ensureTauriRuntime() {
  if (!isTauriRuntime()) {
    throw new Error("请在 Tauri 桌面应用中使用此功能");
  }
}

export default App;
