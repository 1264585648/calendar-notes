declare module "lunar-javascript" {
  export class Solar {
    static fromYmd(year: number, month: number, day: number): Solar;
    getFestivals(): string[];
    getLunar(): Lunar;
  }

  export class Lunar {
    getDay(): number;
    getDayInChinese(): string;
    getMonthInChinese(): string;
    getFestivals(): string[];
  }
}
