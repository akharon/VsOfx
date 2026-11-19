import { describe, expect, it } from 'vitest';
import { formatDate, formatNumber, isDateFormat } from '../src/webview/adapters/numFmt.ts';

describe('number formats', () => {
  it.each([
    [1234.5, undefined, '1234.5'],
    [1234.5, 'General', '1234.5'],
    [0.1 + 0.2, 'General', '0.3'],
    [1234.5, '0', '1235'],
    [1234.5, '0.00', '1234.50'],
    [1234.5, '#,##0.00', '1,234.50'],
    [1234567, '#,##0', '1,234,567'],
    [0.256, '0%', '26%'],
    [0.256, '0.0%', '25.6%'],
    [5, '000', '005'],
    [1.5, '#.##', '1.5'],
    [2, '#.##', '2'],
    [0.5, '#.##', '.5'],
    [-12.5, '0.00', '-12.50'],
    [-12.5, '0.00;(0.00)', '(12.50)'],
    [0, '0.00;-0.00;"zero"', 'zero'],
    [1234.5, '"$"#,##0.00', '$1,234.50'],
    [1234.5, '[$€-407] #,##0.00', '€ 1,234.50'],
    [1234.5, '#,##0.00 "kr"', '1,234.50 kr'],
    [12345.678, '0.00E+00', '1.23E+04'],
    [0.5, '# ?/?', '0.5'],
    [3, '[Red]0.0', '3.0'],
  ])('formats %s with %s', (n, fmt, expected) => {
    expect(formatNumber(n, fmt)).toBe(expected);
  });

  it('recognises date formats', () => {
    expect(isDateFormat('yyyy-mm-dd')).toBe(true);
    expect(isDateFormat('h:mm AM/PM')).toBe(true);
    expect(isDateFormat('0.00')).toBe(false);
    expect(isDateFormat('#,##0 "days"')).toBe(false);
    expect(isDateFormat('General')).toBe(false);
  });

  it('formats dates and times', () => {
    const d = new Date(Date.UTC(2024, 0, 5, 14, 7, 9));
    expect(formatDate(d, 'yyyy-mm-dd')).toBe('2024-01-05');
    expect(formatDate(d, 'dd/mm/yy')).toBe('05/01/24');
    expect(formatDate(d, 'd mmm yyyy')).toBe('5 Jan 2024');
    expect(formatDate(d, 'dddd, mmmm d')).toBe('Friday, January 5');
    expect(formatDate(d, 'hh:mm:ss')).toBe('14:07:09'); // mm after hh = minutes
    expect(formatDate(d, 'h:mm AM/PM')).toBe('2:07 PM');
    expect(formatDate(d, 'yyyy-mm-dd hh:mm')).toBe('2024-01-05 14:07');
    expect(formatDate(d, undefined)).toBe('2024-01-05 14:07:09');
  });
});
