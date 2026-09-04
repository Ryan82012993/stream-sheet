declare module 'luckyexcel' {
  export default class LuckyExcel {
    static transformExcelToLucky(
      file: File,
      callback: (exportJson: any, luckysheetHtml: string) => void,
      errorHandler?: (err: any) => void
    ): void;
    static transformExcelToLuckyByUrl(
      url: string,
      callback: (exportJson: any, luckysheetHtml: string) => void,
      errorHandler?: (err: any) => void
    ): void;
  }
}
