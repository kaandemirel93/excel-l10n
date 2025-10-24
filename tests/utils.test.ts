import { colIndexToLetter, colLetterToIndex, makeTuId } from '../src/utils/index';

test('column conversions', () => {
  expect(colIndexToLetter(1)).toBe('A');
  expect(colIndexToLetter(26)).toBe('Z');
  expect(colIndexToLetter(27)).toBe('AA');
  expect(colLetterToIndex('A')).toBe(1);
  expect(colLetterToIndex('AA')).toBe(27);
});

test('tu id stable', () => {
  expect(makeTuId('Sheet 1', 2, 'A')).toBe('Sheet%201::R2CA');
});
