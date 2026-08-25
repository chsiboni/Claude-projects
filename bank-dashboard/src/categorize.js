'use strict';

const { db } = require('./db');

/**
 * Charges the bank posts for a credit card's monthly total. When card sources
 * are connected too, these line-items duplicate the individual card charges,
 * so they are tracked separately and kept out of spend totals.
 */
const CARD_SETTLEMENT = 'חיוב כרטיס אשראי';
const INCOME = 'הכנסה';
const INTERNAL = 'העברות ומזומן';
const OTHER = 'אחר';

/** Categories that are movements of your own money, not spending. */
const NON_SPEND = new Set([INCOME, CARD_SETTLEMENT, 'חיסכון והשקעות']);

/** First rule that matches wins, so the specific ones come before the broad ones. */
const RULES = [
  [CARD_SETTLEMENT, ['לאומי קארד', 'ישראכרט', 'כרטיסי אשראי', 'כאל', 'מקס איט', 'max it', 'ויזה כ.א.ל', 'אמריקן אקספרס', 'דיינרס', 'חיוב כרטיס']],
  [INCOME, ['משכורת', 'משכורתו', 'שכר', 'הכנסה', 'זיכוי', 'החזר', 'ביטוח לאומי', 'קצבה', 'דמי לידה', 'דיבידנד', 'ריבית זכות']],
  ['משכנתא והלוואות', ['משכנתא', 'הלוואה', 'הלואה', 'פרעון הלוואה', 'החזר הלוואה', 'אשראי לזמן']],
  ['חיסכון והשקעות', ['פקדון', 'פיקדון', 'קרן השתלמות', 'קופת גמל', 'גמל', 'ניירות ערך', 'תיק השקעות', 'חסכון', 'חיסכון', 'מיטב דש', 'אלטשולר', 'ילין לפידות', 'פסגות', 'הראל פיננסים', 'interactive brokers', 'etoro', 'blink']],
  ['ביטוח', ['ביטוח', 'הראל', 'כלל ביטוח', 'מגדל ביטוח', 'מנורה מבטחים', 'הפניקס', 'איילון', 'שירביט', 'ליברה', 'wobi', 'ווישור']],
  ['בריאות ותרופות', ['סופר פארם', 'super-pharm', 'super pharm', 'ניו פארם', 'בי דראג', 'מכבי', 'כללית', 'מאוחדת', 'לאומית שירותי', 'קופת חולים', 'רופא', 'שיניים', 'מרפאה', 'בית חולים', 'אופטיק', 'אופטומטר', 'פיזיותרפ', 'פסיכולוג', 'תרופות', 'מגן דוד אדום']],
  ['חשבונות ותשתיות', ['חברת החשמל', 'חשמל', 'מקורות', 'תאגיד המים', 'תאגיד מים', 'מי אביבים', 'מי שבע', 'הגיחון', 'ארנונה', 'עיריית', 'מועצה מקומית', 'סופרגז', 'אמישראגז', 'פזגז', 'דורגז', 'ועד בית']],
  ['תקשורת', ['סלקום', 'פרטנר', 'פלאפון', 'הוט מובייל', 'הוט', 'yes ', 'בזק', '012 ', '019 ', 'רמי לוי תקשורת', 'גולן טלקום', 'we4g', 'triple c', 'סלולר']],
  ['מנויים ודיגיטל', ['netflix', 'נטפליקס', 'spotify', 'ספוטיפיי', 'youtube', 'יוטיוב', 'apple.com', 'apple ', 'icloud', 'google', 'microsoft', 'openai', 'anthropic', 'claude', 'chatgpt', 'adobe', 'dropbox', 'canva', 'disney', 'hbo', 'amazon prime', 'linkedin', 'notion', 'github', 'figma', 'מנוי']],
  ['דלק ותחבורה', ['פז ', 'פז יעל', 'סונול', 'דור אלון', 'yellow', 'אלונית', 'דלקן', 'תחנת דלק', 'רב קו', 'רב-קו', 'רכבת ישראל', 'אגד', 'מטרופולין', 'gett', 'uber', 'אובר', 'יאנגו', 'yango', 'מוניות', 'פנגו', 'pango', 'סלופארק', 'cellopark', 'חניון', 'חניה', 'כביש 6', 'כביש חוצה ישראל', 'נתיבי איילון']],
  ['רכב', ['מוסך', 'טסט רכב', 'צמיגים', 'חלפים', 'ליסינג', 'שכירות רכב', 'אלבר', 'sixt', 'קל אוטו', 'דלק מוטורס', 'משרד הרישוי', 'אגרת רישוי']],
  ['סופרמרקט', ['שופרסל', 'רמי לוי', 'ויקטורי', 'יינות ביתן', 'טיב טעם', 'אושר עד', 'מגה בעיר', 'am:pm', 'אמפמ', 'מרקט', 'מחסני השוק', 'יוחננוף', 'קווין', 'מכולת', 'חצי חינם', 'זול ובגדול', 'שוק העיר', 'ירקות']],
  ['מסעדות ובתי קפה', ['ארומה', 'קפה', 'לנדוור', 'גרג', 'קופיקס', 'מסעד', 'פיצה', 'בורגר', 'מקדונלד', 'ברגר', 'סושי', 'המבורגר', 'wolt', 'וולט', '10bis', 'תן ביס', 'תן-ביס', 'שווארמה', 'פלאפל', 'חומוס', 'ביסטרו', 'מאפה', 'קונדיטור', 'starbucks', 'kfc', 'דומינו', 'מזנון']],
  ['חינוך וילדים', ['גן ילדים', 'גנון', 'צהרון', 'בית ספר', 'תלמוד תורה', 'אוניברסיט', 'מכללה', 'קורס', 'חוג ', 'משרד החינוך', 'ועד הורים', 'מעון', 'קייטנה', 'שכר לימוד']],
  ['בילויים ופנאי', ['סינמה', 'יס פלאנט', 'רב חן', 'תיאטרון', 'הופעה', 'כרטיסים', 'eventim', 'טיקט', 'בריכה', 'חדר כושר', 'הולמס פלייס', 'גימבורי', 'לונה פארק', 'מוזיאון', 'סופר גיים', 'פארק']],
  ['נסיעות וטיסות', ['אל על', 'טיסה', 'ישראייר', 'ארקיע', 'wizz', 'ryanair', 'booking', 'airbnb', 'expedia', 'מלון', 'צימר', 'issta', 'איסתא', 'דקה 90', 'גוליבר', 'skyscanner', 'נתב"ג', 'שדה תעופה']],
  ['קניות וביגוד', ['zara', 'castro', 'קסטרו', 'fox ', 'פוקס', 'h&m', 'renuar', 'רנואר', 'american eagle', 'terminal x', 'טרמינל איקס', 'גולף', 'דלתא', 'נעלי', 'aliexpress', 'עלי אקספרס', 'shein', 'amazon', 'ebay', 'asos', 'סטימצקי', 'צומת ספרים', 'תכשיט', 'בגדים']],
  ['בית וריהוט', ['איקאה', 'ikea', 'הום סנטר', 'ace hardware', 'א.ס.י', 'ורדינון', 'רהיטים', 'מזרן', 'מחסני חשמל', 'אלקטרה', 'ביתילי', 'קרמיקה', 'שיפוצים', 'חשמלאי', 'אינסטלטור', 'נגר']],
  ['עמלות בנק וריבית', ['עמלה', 'עמלות', 'עמלת', 'דמי ניהול', 'ריבית חובה', 'ריבית', 'דמי כרטיס', 'הפרשי המרה', 'דמי טיפול', 'דמי פנקס']],
  [INTERNAL, ['העברה', 'העברת', 'משיכת מזומן', 'כספומט', 'הפקדה', 'ביט ', 'paybox', 'פייבוקס', 'מזומן', 'שיק', "צ'ק", 'משיכה']],
];

/** Longest patterns first, so "רמי לוי תקשורת" beats "רמי לוי". */
const COMPILED = RULES.map(([category, patterns]) => ({
  category,
  patterns: [...patterns].sort((a, b) => b.length - a.length).map((p) => p.toLowerCase()),
}));

let overridesCache = null;

function loadOverrides() {
  if (!overridesCache) {
    overridesCache = db.prepare('SELECT pattern, category FROM category_overrides').all();
  }
  return overridesCache;
}

function invalidateOverrides() {
  overridesCache = null;
}

/**
 * @param {string} description merchant text as it appears on the statement
 * @param {number} amount signed, so a positive amount can imply income
 */
function categorize(description, amount) {
  const text = `${description || ''}`.toLowerCase().trim();
  if (!text) return OTHER;

  for (const { pattern, category } of loadOverrides()) {
    if (text.includes(pattern.toLowerCase())) return category;
  }

  for (const { category, patterns } of COMPILED) {
    if (patterns.some((p) => text.includes(p))) {
      // A refund at a shop is still that shop's category, not income.
      if (category === INCOME && amount < 0) continue;
      return category;
    }
  }

  return amount > 0 ? INCOME : OTHER;
}

/** All categories the rules can produce, for the UI's filter list. */
function allCategories() {
  return [...new Set([...RULES.map(([c]) => c), OTHER])];
}

module.exports = {
  categorize,
  allCategories,
  invalidateOverrides,
  NON_SPEND,
  CARD_SETTLEMENT,
  INCOME,
  INTERNAL,
  OTHER,
};
