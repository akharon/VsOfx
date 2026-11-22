// Worker of the spreadsheet family: xlsx (exceljs), ods, csv (papaparse).
import type { Format } from '../../shared/model';
import type { FormatAdapter } from '../adapters/Adapter';
import { CsvAdapter } from '../adapters/csvAdapter';
import { OdsAdapter } from '../adapters/odsAdapter';
import { XlsxAdapter } from '../adapters/xlsxAdapter';
import { runWorker } from './core';

runWorker((format: Format): FormatAdapter => {
  switch (format) {
    case 'xlsx':
      return new XlsxAdapter();
    case 'ods':
      return new OdsAdapter();
    case 'csv':
      return new CsvAdapter();
    default:
      throw new Error(`The sheet worker cannot open ${format} files.`);
  }
});
