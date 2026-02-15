/**
 * List handling functions for Google Docs export
 */

const {
  INDENT_LEVEL_0_MAX,
  INDENT_LEVEL_1_MIN,
  NUMBERED_GLYPH_TYPES,
  BULLET_STYLE_MAP,
  NUMBER_STYLE_MAP
} = require('./constants');

/**
 * Infer nesting level when not explicitly provided by the API.
 * Uses indentation as a signal to determine the correct level.
 *
 * @param {object} bullet - The bullet object from the paragraph
 * @param {object} paragraphStyle - The paragraph style containing indentStart
 * @param {number} prevLevel - The previous item's nesting level
 * @returns {number} The inferred nesting level
 */
function inferNestingLevel(bullet, paragraphStyle, prevLevel) {
  if (bullet?.nestingLevel !== undefined) {
    return bullet.nestingLevel;
  }

  const indentStart = paragraphStyle?.indentStart?.magnitude || 0;

  // Use indentation heuristics
  if (indentStart <= INDENT_LEVEL_0_MAX) {
    return 0;
  } else if (indentStart >= INDENT_LEVEL_1_MIN) {
    return 1;
  }

  // Fallback: continue at previous level if in a list
  return (prevLevel >= 0) ? prevLevel : 0;
}

/**
 * Detect the bullet style for an unordered list.
 *
 * @param {object} glyph - The glyph definition from list properties
 * @returns {string} The bullet style name (disc, circle, square, dash)
 */
function detectBulletStyle(glyph) {
  if (!glyph?.glyphSymbol) return 'disc';
  return BULLET_STYLE_MAP[glyph.glyphSymbol] || 'disc';
}

/**
 * Detect the numbering style for an ordered list.
 *
 * @param {object} glyph - The glyph definition from list properties
 * @returns {string} The numbering style name (decimal, upper-alpha, etc.)
 */
function detectNumberStyle(glyph) {
  if (!glyph?.glyphType) return 'decimal';
  return NUMBER_STYLE_MAP[glyph.glyphType] || 'decimal';
}

/**
 * Determine if a list is numbered based on glyph properties and item counts.
 *
 * @param {object} glyph - The glyph definition from list properties
 * @param {string} listId - The list identifier
 * @param {number} nestingLevel - The nesting level
 * @param {object} listItemCounts - The item count map
 * @returns {boolean} True if the list is numbered
 */
function isNumberedList(glyph, listId, nestingLevel, listItemCounts) {
  // If glyphSymbol is present (●, ○, -, etc.) → bullet list
  const isBullet = glyph?.glyphSymbol !== undefined;
  if (isBullet) return false;

  // If glyphType is explicitly a numbered type → numbered list
  const explicitlyNumbered = NUMBERED_GLYPH_TYPES.includes(glyph?.glyphType);
  if (explicitlyNumbered) return true;

  // If GLYPH_TYPE_UNSPECIFIED: use item count heuristic
  // Single-item lists → numbered (section markers)
  // Multi-item lists → bullets
  if (glyph?.glyphType === 'GLYPH_TYPE_UNSPECIFIED') {
    const key = `${listId}:${nestingLevel}`;
    const itemCount = listItemCounts?.[key] || 1;
    return (itemCount === 1);
  }

  // Default to bullet list
  return false;
}

/**
 * Detect list state changes and return actions to perform.
 *
 * @param {object} paragraph - The paragraph containing bullet information
 * @param {object} doc - The full document object
 * @param {Array} listStack - Stack tracking open lists
 * @param {boolean} isRTL - Whether the list is right-to-left
 * @param {number} prevLevel - Previous nesting level
 * @param {string} prevListId - Previous list ID
 * @returns {string|null} Pipe-separated list of actions or null
 */
function detectListChange(paragraph, doc, listStack, isRTL, prevLevel, prevListId) {
  const bullet = paragraph.bullet;
  if (!bullet) return null;

  const listId = bullet.listId;
  const nestingLevel = inferNestingLevel(bullet, paragraph.paragraphStyle, prevLevel);

  const listDef = doc.lists?.[listId];
  if (!listDef?.listProperties?.nestingLevels) return null;

  const glyph = listDef.listProperties.nestingLevels[nestingLevel];
  if (!glyph) return null;

  // Determine if this is a numbered or bullet list
  const isNumbered = isNumberedList(glyph, listId, nestingLevel, doc.___listItemCounts);

  // Detect the specific style
  const bulletStyle = isNumbered ? null : detectBulletStyle(glyph);
  const numberStyle = isNumbered ? detectNumberStyle(glyph) : null;

  // Build the list type identifier
  const startType = isNumbered ? 'OL' : 'UL';
  const rtlFlag = isRTL ? '_RTL' : '';
  const styleFlag = isNumbered
    ? (numberStyle !== 'decimal' ? `_${numberStyle.toUpperCase().replace(/-/g, '_')}` : '')
    : (bulletStyle !== 'disc' ? `_${bulletStyle.toUpperCase()}` : '');

  // Starting a list for the first time
  if (listStack.length === 0) {
    return `start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`;
  }

  // Check if nesting level changed
  if (nestingLevel > prevLevel) {
    // Going deeper - start nested list
    return `start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`;
  } else if (nestingLevel < prevLevel) {
    // Coming back up - close nested lists
    let actions = [];
    for (let i = prevLevel; i > nestingLevel; i--) {
      actions.push(`endLIST`);
    }

    // After closing nested lists, check if we need to switch lists at current level
    const stackIndexAfterClosing = listStack.length - (prevLevel - nestingLevel);
    if (stackIndexAfterClosing > 0) {
      const parentType = listStack[stackIndexAfterClosing - 1]?.split(':')[0];
      const wantType = startType.toLowerCase() + (isRTL ? '_rtl' : '') + styleFlag.toLowerCase();

      // Check if parent list type changed
      if (parentType !== wantType) {
        actions.push(`end${parentType?.toUpperCase() || 'UL'}`);
        actions.push(`start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`);
      }
    }

    return actions.join('|');
  }

  // Same level - check if list type changed
  const currentType = listStack[listStack.length - 1]?.split(':')[0];
  const wantType = startType.toLowerCase() + (isRTL ? '_rtl' : '') + styleFlag.toLowerCase();

  // Only switch lists if the TYPE changed (OL vs UL)
  if (currentType !== wantType) {
    return `end${currentType?.toUpperCase() || 'UL'}|start${startType}${rtlFlag}${styleFlag}:${nestingLevel}`;
  }

  return null;
}

/**
 * Handle list state transitions by processing list change actions.
 * Actions can be: startUL:0, startOL_RTL:1, startUL_DASH:0, endLIST, endUL, etc.
 *
 * @param {string} listChange - Pipe-separated list of actions
 * @param {Array} listStack - Stack tracking open lists
 * @param {Array} htmlLines - Array of HTML lines being built
 */
function handleListState(listChange, listStack, htmlLines) {
  if (!listChange) return;

  const actions = listChange.split('|');
  for (const action of actions) {
    if (action.startsWith('start')) {
      // Extract type and level (format: "startUL:0", "startOL_RTL:1", "startUL_DASH:0")
      const parts = action.split(':');
      const typeInfo = parts[0].replace('start', '');
      const level = parts[1] || '0';

      // Parse style flags (DASH, CIRCLE, SQUARE for UL; UPPER_ALPHA, LOWER_ROMAN, etc. for OL)
      let listStyle = '';
      if (typeInfo.includes('_DASH')) {
        listStyle = ' style="list-style-type: \'− \'"';
      } else if (typeInfo.includes('_CIRCLE')) {
        listStyle = ' style="list-style-type: circle"';
      } else if (typeInfo.includes('_SQUARE')) {
        listStyle = ' style="list-style-type: square"';
      } else if (typeInfo.includes('_UPPER_ALPHA')) {
        listStyle = ' style="list-style-type: upper-alpha"';
      } else if (typeInfo.includes('_LOWER_ALPHA')) {
        listStyle = ' style="list-style-type: lower-alpha"';
      } else if (typeInfo.includes('_UPPER_ROMAN')) {
        listStyle = ' style="list-style-type: upper-roman"';
      } else if (typeInfo.includes('_LOWER_ROMAN')) {
        listStyle = ' style="list-style-type: lower-roman"';
      }

      if (typeInfo.includes('UL_RTL') || (typeInfo.includes('UL') && typeInfo.includes('_RTL'))) {
        htmlLines.push(`<ul dir="rtl"${listStyle}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      } else if (typeInfo.includes('OL_RTL')) {
        htmlLines.push(`<ol dir="rtl"${listStyle}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      } else if (typeInfo.includes('UL')) {
        htmlLines.push(`<ul${listStyle}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      } else {
        htmlLines.push(`<ol${listStyle}>`);
        listStack.push(`${typeInfo.toLowerCase()}:${level}`);
      }
    } else if (action === 'endLIST') {
      const top = listStack.pop();
      if (!top) continue;
      const listType = top.split(':')[0];
      // Close the nested list
      if (listType.startsWith('u')) htmlLines.push('</ul>');
      else htmlLines.push('</ol>');
      // Close the parent <li> that contained the nested list
      htmlLines.push('</li>');
    } else if (action.startsWith('end')) {
      const top = listStack.pop();
      if (!top) continue;
      const listType = top.split(':')[0];
      if (listType.startsWith('u')) htmlLines.push('</ul>');
      else htmlLines.push('</ol>');
    }
  }
}

/**
 * Close all open lists in the stack.
 *
 * @param {Array} listStack - Stack of open lists
 * @param {Array} htmlLines - Array of HTML lines being built
 */
function closeAllLists(listStack, htmlLines) {
  if (!listStack || !htmlLines) return;

  while (listStack.length > 0) {
    const top = listStack.pop();
    if (!top) continue;
    if (top.startsWith('u')) htmlLines.push('</ul>');
    else htmlLines.push('</ol>');
  }
}

module.exports = {
  inferNestingLevel,
  detectBulletStyle,
  detectNumberStyle,
  isNumberedList,
  detectListChange,
  handleListState,
  closeAllLists
};
