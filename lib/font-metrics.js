'use strict';

const fontkit = require('fontkit');

const metricPromises = new Map();

function collectFontFamilies(value, families = new Set()) {
  if (!value || typeof value !== 'object') return families;
  const family = value.weightedFontFamily?.fontFamily;
  if (family) families.add(family);
  for (const child of Object.values(value)) collectFontFamilies(child, families);
  return families;
}

async function loadGoogleFontMetrics(families) {
  const entries = await Promise.all(
    [...families].map(async family => [family, await loadGoogleFontMetric(family)])
  );
  return new Map(entries.filter(([, metric]) => metric));
}

async function loadGoogleFontMetric(family) {
  if (!metricPromises.has(family)) {
    metricPromises.set(family, fetchGoogleFontMetric(family));
  }
  return metricPromises.get(family);
}

async function fetchGoogleFontMetric(family) {
  try {
    const query = encodeURIComponent(family).replace(/%20/g, '+');
    const cssResponse = await fetch(`https://fonts.googleapis.com/css2?family=${query}:wght@400`, {
      headers: { 'user-agent': 'Mozilla/5.0' },
      signal: AbortSignal.timeout(10000)
    });
    if (!cssResponse.ok) return null;
    const css = await cssResponse.text();
    const source = css.match(/src:\s*url\((https:\/\/[^)]+)\)/)?.[1];
    if (!source) return null;

    const fontResponse = await fetch(source, { signal: AbortSignal.timeout(10000) });
    if (!fontResponse.ok) return null;
    const font = fontkit.create(Buffer.from(await fontResponse.arrayBuffer()));
    if (!font.unitsPerEm) return null;

    return {
      unitsPerEm: font.unitsPerEm,
      ascent: font.ascent,
      descent: font.descent,
      lineGap: font.lineGap || 0
    };
  } catch {
    return null;
  }
}

function naturalLineHeight(metric) {
  if (!metric?.unitsPerEm) return null;
  return (metric.ascent - metric.descent + metric.lineGap) / metric.unitsPerEm;
}

function lineHeightFor(fontFamily, spacingPercent, metrics) {
  const natural = naturalLineHeight(metrics?.get(fontFamily));
  return roundCSSNumber((natural ?? 1) * spacingPercent / 100);
}

function roundCSSNumber(value) {
  return Math.round(value * 1e6) / 1e6;
}

module.exports = {
  collectFontFamilies,
  loadGoogleFontMetrics,
  naturalLineHeight,
  lineHeightFor
};
