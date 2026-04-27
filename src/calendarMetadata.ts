import { Solar } from "lunar-javascript";

export type HolidayKind = "holiday" | "workday";

export type CalendarDayMetadata = {
  dateKey: string;
  lunarLabel: string;
  lunarFestival?: string;
  solarFestival?: string;
  holidayName?: string;
  holidayKind?: HolidayKind;
};

type HolidayRecord = {
  name: string;
  kind: HolidayKind;
};

const CHINA_2026_HOLIDAYS: Record<string, HolidayRecord> = {
  "2026-01-01": { name: "元旦", kind: "holiday" },
  "2026-01-02": { name: "元旦", kind: "holiday" },
  "2026-01-03": { name: "元旦", kind: "holiday" },
  "2026-01-04": { name: "调休上班", kind: "workday" },

  "2026-02-14": { name: "调休上班", kind: "workday" },
  "2026-02-15": { name: "春节", kind: "holiday" },
  "2026-02-16": { name: "除夕", kind: "holiday" },
  "2026-02-17": { name: "春节", kind: "holiday" },
  "2026-02-18": { name: "春节", kind: "holiday" },
  "2026-02-19": { name: "春节", kind: "holiday" },
  "2026-02-20": { name: "春节", kind: "holiday" },
  "2026-02-21": { name: "春节", kind: "holiday" },
  "2026-02-22": { name: "春节", kind: "holiday" },
  "2026-02-23": { name: "春节", kind: "holiday" },
  "2026-02-28": { name: "调休上班", kind: "workday" },

  "2026-04-04": { name: "清明节", kind: "holiday" },
  "2026-04-05": { name: "清明节", kind: "holiday" },
  "2026-04-06": { name: "清明节", kind: "holiday" },

  "2026-05-01": { name: "劳动节", kind: "holiday" },
  "2026-05-02": { name: "劳动节", kind: "holiday" },
  "2026-05-03": { name: "劳动节", kind: "holiday" },
  "2026-05-04": { name: "劳动节", kind: "holiday" },
  "2026-05-05": { name: "劳动节", kind: "holiday" },
  "2026-05-09": { name: "调休上班", kind: "workday" },

  "2026-06-19": { name: "端午节", kind: "holiday" },
  "2026-06-20": { name: "端午节", kind: "holiday" },
  "2026-06-21": { name: "端午节", kind: "holiday" },

  "2026-09-20": { name: "调休上班", kind: "workday" },
  "2026-09-25": { name: "中秋节", kind: "holiday" },
  "2026-09-26": { name: "中秋节", kind: "holiday" },
  "2026-09-27": { name: "中秋节", kind: "holiday" },

  "2026-10-01": { name: "国庆节", kind: "holiday" },
  "2026-10-02": { name: "国庆节", kind: "holiday" },
  "2026-10-03": { name: "国庆节", kind: "holiday" },
  "2026-10-04": { name: "国庆节", kind: "holiday" },
  "2026-10-05": { name: "国庆节", kind: "holiday" },
  "2026-10-06": { name: "国庆节", kind: "holiday" },
  "2026-10-07": { name: "国庆节", kind: "holiday" },
  "2026-10-10": { name: "调休上班", kind: "workday" },
};

export function getCalendarDayMetadata(date: Date): CalendarDayMetadata {
  const dateKey = toDateKey(date);
  const solar = Solar.fromYmd(date.getFullYear(), date.getMonth() + 1, date.getDate());
  const lunar = solar.getLunar();
  const lunarFestivals = lunar.getFestivals();
  const solarFestivals = solar.getFestivals();
  const holiday = CHINA_2026_HOLIDAYS[dateKey];
  const lunarLabel = lunar.getDay() === 1
    ? `${lunar.getMonthInChinese()}月`
    : lunar.getDayInChinese();

  return {
    dateKey,
    lunarLabel,
    lunarFestival: lunarFestivals[0],
    solarFestival: solarFestivals[0],
    holidayName: holiday?.name,
    holidayKind: holiday?.kind,
  };
}

export function isChinaHoliday(dateKey: string) {
  return CHINA_2026_HOLIDAYS[dateKey]?.kind === "holiday";
}

export function isChinaAdjustedWorkday(dateKey: string) {
  return CHINA_2026_HOLIDAYS[dateKey]?.kind === "workday";
}

function toDateKey(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
