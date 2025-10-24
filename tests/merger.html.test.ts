import { _composeHtmlFromSkeleton } from '../src/merger/index.js';

describe('composeHtmlFromSkeleton', () => {
  test('reconstructs HTML with inline placeholders', () => {
    const skeleton = '<div>[[htmltxt:1]][[io:1]][[htmltxt:2]][[ic:1]][[htmltxt:3]]</div>';
    const inlineMap = { '1': { open: '<b>', close: '</b>' } };
    const translated = 'Hola, [[io:1]]negrita[[ic:1]] texto.';
    expect(_composeHtmlFromSkeleton(translated, skeleton, inlineMap)).toBe('<div>Hola, <b>negrita</b> texto.</div>');
  });

  test('ignores extra whitespace-only text tokens gracefully', () => {
    const skeleton = '<p>[[htmltxt:1]]<span>[[io:2]][[htmltxt:2]][[ic:2]]</span>[[htmltxt:3]]</p>';
    const inlineMap = { '2': { open: '<i>', close: '</i>' } };
    const translated = 'Primero [[io:2]]segundo[[ic:2]] fin';
    expect(_composeHtmlFromSkeleton(translated, skeleton, inlineMap)).toBe('<p>Primero <span><i>segundo</i></span> fin</p>');
  });
});
