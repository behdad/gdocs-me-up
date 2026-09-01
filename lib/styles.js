/**
 * Deduplicates generated CSS declarations and gives them compact, stable names.
 *
 * Google Docs repeats the same computed paragraph and text styles many times. Keeping
 * those declarations in a stylesheet makes the HTML substantially smaller and easier
 * to inspect without sacrificing document-specific fidelity.
 */
class StyleRegistry {
  constructor() {
    this.styles = new Map();
    this.rules = [];
    this.counts = new Map();
  }

  add(prefix, declarations) {
    const css = normalizeDeclarations(declarations);
    if (!css) return '';

    const key = `${prefix}\0${css}`;
    const existing = this.styles.get(key);
    if (existing) return existing;

    const count = (this.counts.get(prefix) || 0) + 1;
    this.counts.set(prefix, count);
    const className = `${prefix}${count}`;
    this.styles.set(key, className);
    this.rules.push(`.${className}{${css}}`);
    return className;
  }

  toCSS() {
    return this.rules.join('\n');
  }
}

function normalizeDeclarations(declarations) {
  if (!declarations) return '';
  if (typeof declarations === 'string') {
    return declarations.trim().replace(/^;+|;+$/g, '') +
      (declarations.trim() ? ';' : '');
  }

  return Object.entries(declarations)
    .filter(([, value]) => value !== undefined && value !== null && value !== '')
    .map(([property, value]) => `${property}:${value};`)
    .join('');
}

function joinClasses(...classes) {
  return classes.flat().filter(Boolean).join(' ');
}

module.exports = {
  StyleRegistry,
  normalizeDeclarations,
  joinClasses
};
