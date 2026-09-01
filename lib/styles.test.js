const { StyleRegistry, normalizeDeclarations, joinClasses } = require('./styles');

describe('StyleRegistry', () => {
  test('deduplicates identical styles within a prefix', () => {
    const styles = new StyleRegistry();
    expect(styles.add('t', 'color:red;')).toBe('t1');
    expect(styles.add('t', 'color:red;')).toBe('t1');
    expect(styles.toCSS()).toBe('.t1{color:red;}');
  });

  test('keeps style namespaces separate', () => {
    const styles = new StyleRegistry();
    expect(styles.add('p', 'color:red')).toBe('p1');
    expect(styles.add('t', 'color:red')).toBe('t1');
  });

  test('ignores empty styles', () => {
    const styles = new StyleRegistry();
    expect(styles.add('p', '')).toBe('');
    expect(styles.toCSS()).toBe('');
  });
});

describe('style helpers', () => {
  test('normalizes strings and objects', () => {
    expect(normalizeDeclarations(';color:red;;')).toBe('color:red;');
    expect(normalizeDeclarations({ color: 'red', empty: '' })).toBe('color:red;');
  });

  test('joins truthy classes', () => {
    expect(joinClasses('title', '', ['p1', null])).toBe('title p1');
  });
});
