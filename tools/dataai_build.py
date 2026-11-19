"""Build everything /data-ai/ needs from the practice files:

  data-ai/files/NN-clean.xlsx   end of round 2 (duplicates removed, text numbers/dates fixed, neighbourhood names unified)
  data-ai/files/NN-rich.xlsx    end of round 3 (clean + the new columns the round teaches)
  data-ai/files.js              facts per file for the pages: columns, formulas with real column letters, expected answers

Every formula shown on the site is evaluated here in a real Excel (Windows + pywin32), so a formula that
does not work fails the build instead of reaching a participant.

Run:  python tools/dataai_build.py
"""
import os, re, json, datetime, collections, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'data', 'practice-files', 'files')
OUT = os.path.join(ROOT, 'data-ai', 'files')
TMP = os.path.join(os.environ.get('TEMP', '/tmp'), 'dataai_build')
HOOD = 'שכונה'
FIX = 'שכונה מתוקנת'
ASK = 'לבירור'
HOODS = {'אבנ"ח': 'אבני חן', 'אבני-חן': 'אבני חן', 'מרכז': 'מרכז העיר', 'רעות.': 'רעות', 'בוכמאן': 'בוכמן',
         'מורייה': 'מוריה', 'נביאים': 'הנביאים', 'שבטים': 'השבטים', 'הצפורים': 'הציפורים', 'מכבים-רעות': ASK}
MONTH = lambda c: '=IF([[%s]]2="","",YEAR([[%s]]2)&"-"&TEXT(MONTH([[%s]]2),"00"))' % (c, c, c)
XL = lambda c: '=IF([[%s]]2="","",XLOOKUP([[%s]]2,מיפוי!A:A,מיפוי!B:B,"לא ידוע"))' % (c, c)

# cols: five new columns per file, always in this order:
#   A calculation (basic) · B date part or second simple column (basic) · C yes/no flag (medium)
#   D grouping (medium; with "map" it is a lookup table) · E several conditions in one LET (advanced)
# p: three cross-tabs (title, row field, aggregation, value field). kpis: three numbers for the product round.
FILES = {
'01': dict(short='מוקד 106', topic='קריאות למוקד 106', role='מנהל/ת המוקד', who='כל מי שאין לו קובץ מהתחום שלו', key=0,
  ask='איזה נושא ואיזו שכונה חורגים הכי הרבה מה-SLA?',
  q3=[('כמה קריאות יש בנושא תאורה?', '=COUNTIF([[נושא ראשי]]:[[נושא ראשי]],"תאורה")'),
      ('מה משך הטיפול הממוצע, בשעות?', '=AVERAGE([[משך טיפול]]:[[משך טיפול]])'),
      ('כמה קריאות עדיין בטיפול?', '=COUNTIF([[סטטוס]]:[[סטטוס]],"בטיפול")')],
  mid=('מה משך הטיפול הממוצע בקריאות תברואה?', '=AVERAGEIF([[נושא ראשי]]:[[נושא ראשי]],"תברואה",[[משך טיפול]]:[[משך טיפול]])'),
  adv='איזה גורם מטפל חורג הכי הרבה מיעד ה-SLA?',
  cols=[('חריגה מהיעד (שעות)', 'משך הטיפול פחות יעד ה-SLA', '=IF(OR([[משך טיפול]]2="",[[יעד SLA]]2=""),"",[[משך טיפול]]2-[[יעד SLA]]2)'),
        ('חודש פתיחה', 'שנה וחודש מתוך תאריך הפתיחה', MONTH('תאריך ושעת פתיחה')),
        ('חרג מ-SLA', 'כן אם משך הטיפול ארוך מהיעד', '=IF(OR([[משך טיפול]]2="",[[יעד SLA]]2=""),"",IF([[משך טיפול]]2>[[יעד SLA]]2,"כן","לא"))'),
        ('סוג ערוץ', 'ערוצי הקליטה בשתי קבוצות: דיגיטלי ואנושי', XL('ערוץ קליטה')),
        ('מצב טיפול', 'פתוח, בזמן, חריגה, או חריגה חמורה (יותר מפי 2 מהיעד)', '=LET(d,[[משך טיפול]]2,t,[[יעד SLA]]2,IF(OR(d="",t=""),"פתוח",IF(d>2*t,"חריגה חמורה",IF(d>t,"חריגה","בזמן"))))')],
  map=('ערוץ קליטה', [('אפליקציה עירונית', 'דיגיטלי'), ('וואטסאפ', 'דיגיטלי'), ('אתר העירייה', 'דיגיטלי'), ('טלפון', 'אנושי'), ('פנייה פרונטלית', 'אנושי')]),
  p=[('קריאות שחרגו מה-SLA לפי נושא', 'נושא ראשי', 'yes', 'חרג מ-SLA'), ('קריאות שחרגו מה-SLA לפי שכונה', FIX, 'yes', 'חרג מ-SLA'), ('קריאות לפי חודש', 'חודש פתיחה', 'count', None)],
  kpis=[('מספר הקריאות', '=COUNTA([[מס\' קריאה]]:[[מס\' קריאה]])-1'), ('קריאות שחרגו מה-SLA', '=COUNTIF([[חרג מ-SLA]]:[[חרג מ-SLA]],"כן")'), ('משך טיפול ממוצע (שעות)', '=AVERAGE([[משך טיפול]]:[[משך טיפול]])')],
  filters=['נושא ראשי', FIX, 'חודש פתיחה'], charts=['קריאות שחרגו מה-SLA לפי נושא', 'מספר הקריאות לפי חודש']),

'02': dict(short='חינוך', topic='מוסדות החינוך', role='מנהל/ת אגף החינוך', who='חינוך', key=None,
  ask='איפה הצפיפות או האלימות עולות משנה לשנה?',
  q3=[('כמה שורות יש לבתי ספר יסודיים?', '=COUNTIF([[שלב חינוך]]:[[שלב חינוך]],"יסודי")'),
      ('מה אחוז הנוכחות הממוצע?', '=AVERAGE([[אחוז נוכחות]]:[[אחוז נוכחות]])'),
      ('כמה אירועי אלימות דווחו בסך הכול?', '=SUM([[אירועי אלימות]]:[[אירועי אלימות]])')],
  mid=('כמה תלמידים למדו בשנת תשפ"ו?', '=SUMIF([[שנת לימודים]]:[[שנת לימודים]],"תשפ""ו",[[מס\' תלמידים]]:[[מס\' תלמידים]])'),
  adv='באילו מוסדות מספר אירועי האלימות עלה משנה לשנה?',
  cols=[('אחוז תלמידי שילוב', 'תלמידי השילוב חלקי כלל התלמידים', '=IF(OR([[מס\' תלמידים]]2="",[[מס\' תלמידים]]2=0),"",[[תלמידי שילוב]]2/[[מס\' תלמידים]]2)'),
        ('אלימות ל-100 תלמידים', 'אירועי האלימות ביחס לגודל המוסד', '=IF(OR([[מס\' תלמידים]]2="",[[מס\' תלמידים]]2=0),"",[[אירועי אלימות]]2/[[מס\' תלמידים]]2*100)'),
        ('כיתות צפופות', 'כן אם הממוצע לכיתה מעל 31', '=IF([[ממוצע תלמידים]]2="","",IF([[ממוצע תלמידים]]2>31,"כן","לא"))'),
        ('קבוצת שלב', 'שלבי החינוך בשתי קבוצות: יסודי ועל-יסודי', XL('שלב חינוך')),
        ('מצב מוסד', 'צפיפות (מעל 31 בכיתה) ואלימות (מעל 0.8 אירועים ל-100 תלמידים) יחד', '=LET(c,[[ממוצע תלמידים]]2,s,[[מס\' תלמידים]]2,v,[[אירועי אלימות]]2,IF(OR(c="",s="",s=0),"",IF(AND(c>31,v/s*100>0.8),"צפוף ואלימות גבוהה",IF(c>31,"צפוף",IF(v/s*100>0.8,"אלימות גבוהה","תקין")))))')],
  map=('שלב חינוך', [('יסודי', 'יסודי'), ('חטיבת ביניים', 'על-יסודי'), ('תיכון', 'על-יסודי')]),
  p=[('אירועי אלימות לפי שנת לימודים', 'שנת לימודים', 'sum', 'אירועי אלימות'), ('ממוצע תלמידים לכיתה לפי שכונה', FIX, 'avg', 'ממוצע תלמידים'), ('מספר התלמידים לפי שנת לימודים', 'שנת לימודים', 'sum', "מס' תלמידים")],
  kpis=[('תלמידים בתשפ"ו', '=SUMIF([[שנת לימודים]]:[[שנת לימודים]],"תשפ""ו",[[מס\' תלמידים]]:[[מס\' תלמידים]])'), ('ממוצע תלמידים לכיתה', '=AVERAGE([[ממוצע תלמידים]]:[[ממוצע תלמידים]])'), ('אירועי אלימות', '=SUM([[אירועי אלימות]]:[[אירועי אלימות]])')],
  filters=['שנת לימודים', 'שלב חינוך', FIX], charts=['אירועי אלימות לפי שנת לימודים', 'ממוצע תלמידים לכיתה לפי שכונה']),

'03': dict(short='פינוי פסולת', topic='פינוי הפסולת', role='מנהל/ת אגף שפ"ע', who='שפ"ע ותברואה', key=None,
  ask='מה העלות לטון, לפי קבלן ולפי שכונה?',
  q3=[("כמה שורות יש לקבלן א'?", '=COUNTIF([[קבלן פינוי]]:[[קבלן פינוי]],"קבלן א\'")'),
      ('מה סך עלות הפינוי?', '=SUM([[עלות פינוי]]:[[עלות פינוי]])'),
      ('כמה קריאות מוקד יש בממוצע בשורה?', '=AVERAGE([[קריאות מוקד]]:[[קריאות מוקד]])')],
  mid=('מה סך עלות הפינוי בשכונת מוריה?', '=SUMIF([[שכונה]]:[[שכונה]],"מוריה",[[עלות פינוי]]:[[עלות פינוי]])'),
  adv='איזה קבלן יקר יותר לטון, ובאילו שכונות?',
  cols=[('עלות לטון (₪)', 'עלות הפינוי חלקי סך הטונות מכל הסוגים', '=IF(SUM([[פסולת מעורבת]]2:[[גזם]]2)=0,"",[[עלות פינוי]]2/SUM([[פסולת מעורבת]]2:[[גזם]]2))'),
        ('שנה', 'השנה מתוך עמודת החודש', '=IF([[חודש]]2="","",RIGHT([[חודש]]2,4))'),
        ('מיחזור נמוך', 'כן אם האריזות והנייר הם פחות מ-10% מהפסולת', '=IF(SUM([[פסולת מעורבת]]2:[[גזם]]2)=0,"",IF(([[אריזות]]2+[[נייר]]2)/SUM([[פסולת מעורבת]]2:[[גזם]]2)<0.1,"כן","לא"))'),
        ('רמת קריאות', 'קריאות המוקד בשלוש מדרגות: עד 20, עד 40, מעל 40', '=IF([[קריאות מוקד]]2="","",IF([[קריאות מוקד]]2<=20,"נמוכה",IF([[קריאות מוקד]]2<=40,"בינונית","גבוהה")))'),
        ('מצב שירות', 'עלות (מעל 340 ₪ לטון) וקריאות (מעל 40) יחד', '=LET(t,SUM([[פסולת מעורבת]]2:[[גזם]]2),k,[[קריאות מוקד]]2,IF(t=0,"",LET(c,[[עלות פינוי]]2/t,IF(AND(c>340,k>40),"יקר והרבה קריאות",IF(c>340,"יקר",IF(k>40,"הרבה קריאות","תקין"))))))')],
  p=[('עלות ממוצעת לטון לפי קבלן', 'קבלן פינוי', 'avg', 'עלות לטון'), ('עלות הפינוי לפי שכונה', FIX, 'sum', 'עלות פינוי'), ('עלות הפינוי לפי שנה', 'שנה', 'sum', 'עלות פינוי')],
  kpis=[('סך עלות הפינוי (₪)', '=SUM([[עלות פינוי]]:[[עלות פינוי]])'), ('עלות ממוצעת לטון (₪)', '=AVERAGE([[עלות לטון]]:[[עלות לטון]])'), ('קריאות מוקד', '=SUM([[קריאות מוקד]]:[[קריאות מוקד]])')],
  filters=['קבלן פינוי', FIX, 'שנה'], charts=['עלות ממוצעת לטון לפי קבלן', 'עלות הפינוי לפי שכונה']),

'04': dict(short='רווחה', topic='תיקי הרווחה', role='מנהל/ת אגף הרווחה', who='רווחה', key=0,
  ask='איפה מרוכזים התיקים הארוכים, ובאיזה תחום?',
  q3=[('כמה תיקים פעילים יש?', '=COUNTIF([[סטטוס]]:[[סטטוס]],"פעיל")'),
      ('כמה מפגשים יש בממוצע לתיק?', '=AVERAGE([[מס\' מפגשים]]:[[מס\' מפגשים]])'),
      ('כמה תיקים יש בתחום "סיוע כלכלי"?', '=COUNTIF([[תחום טיפול]]:[[תחום טיפול]],"סיוע כלכלי")')],
  mid=('כמה מפגשים יש בממוצע בתיקי "מוגבלויות"?', '=AVERAGEIF([[תחום טיפול]]:[[תחום טיפול]],"מוגבלויות",[[מס\' מפגשים]]:[[מס\' מפגשים]])'),
  adv='באילו תחומי טיפול ושכונות מרוכזים התיקים הפעילים והארוכים?',
  cols=[('ימים עד סגירה', 'תאריך הסגירה פחות תאריך הפתיחה', '=IF(OR([[תאריך פתיחה]]2="",[[תאריך סגירה]]2=""),"",[[תאריך סגירה]]2-[[תאריך פתיחה]]2)'),
        ('חודש פתיחה', 'שנה וחודש מתוך תאריך הפתיחה', MONTH('תאריך פתיחה')),
        ('תיק ארוך', 'כן אם היו 30 מפגשים או יותר', '=IF([[מס\' מפגשים]]2="","",IF([[מס\' מפגשים]]2>=30,"כן","לא"))'),
        ('סוג הפניה', 'הגורמים המפנים בשתי קבוצות: עצמי ומוסדי', XL('גורם מפנה')),
        ('מצב תיק', 'סטטוס ומספר מפגשים יחד', '=LET(s,[[סטטוס]]2,m,[[מס\' מפגשים]]2,IF(OR(s="",m=""),"",IF(AND(s="פעיל",m>=30),"פעיל וארוך",IF(s="פעיל","פעיל","לא פעיל"))))')],
  map=('גורם מפנה', [('פנייה עצמית', 'עצמי'), ('קופת חולים', 'מוסדי'), ('בית ספר', 'מוסדי'), ('בית חולים', 'מוסדי'), ('מוקד 106', 'מוסדי'), ('משטרה', 'מוסדי')]),
  p=[('מספר התיקים לפי תחום טיפול', 'תחום טיפול', 'count', None), ('תיקים ארוכים לפי שכונה', FIX, 'yes', 'תיק ארוך'), ('תיקים שנפתחו לפי חודש', 'חודש פתיחה', 'count', None)],
  kpis=[('מספר התיקים', '=COUNTA([[מס\' תיק]]:[[מס\' תיק]])-1'), ('תיקים פעילים', '=COUNTIF([[סטטוס]]:[[סטטוס]],"פעיל")'), ('ממוצע מפגשים לתיק', '=AVERAGE([[מס\' מפגשים]]:[[מס\' מפגשים]])')],
  filters=['תחום טיפול', FIX, 'סטטוס'], charts=['מספר התיקים לפי תחום טיפול', 'תיקים שנפתחו לפי חודש']),

'05': dict(short='היתרי בנייה', topic='בקשות להיתר בנייה', role='מהנדס/ת העיר', who='הוועדה המקומית והנדסה', key=0,
  ask='איזה מסלול רישוי הכי איטי, ואיפה?',
  q3=[('כמה בקשות אושרו?', '=COUNTIF([[החלטה]]:[[החלטה]],"אושר")'),
      ('כמה ימים עוברים בממוצע עד החלטה?', '=AVERAGE([[ימים עד החלטה]]:[[ימים עד החלטה]])'),
      ('מה סך האגרות וההיטלים?', '=SUM([[אגרות]]:[[אגרות]])')],
  mid=('כמה ימים עוברים בממוצע עד החלטה במסלול "רישוי רגיל"?', '=AVERAGEIF([[מסלול רישוי]]:[[מסלול רישוי]],"רישוי רגיל",[[ימים עד החלטה]]:[[ימים עד החלטה]])'),
  adv='איזה מסלול רישוי חורג הכי הרבה מהזמן, ובאילו שכונות?',
  cols=[('אגרה למ"ר (₪)', 'האגרות וההיטלים חלקי השטח המבוקש', '=IF(OR([[שטח מבוקש]]2="",[[שטח מבוקש]]2=0),"",[[אגרות]]2/[[שטח מבוקש]]2)'),
        ('חודש הגשה', 'שנה וחודש מתוך תאריך ההגשה', MONTH('תאריך הגשה')),
        ('חרג מהיעד', 'כן אם ההחלטה לקחה יותר מהיעד למסלול (לדוגמה: 270 יום ברישוי רגיל, 45 בשאר)', '=IF([[ימים עד החלטה]]2="","",IF([[ימים עד החלטה]]2>IF([[מסלול רישוי]]2="רישוי רגיל",270,45),"כן","לא"))'),
        ('קבוצת מהות', 'מהות הבקשה בשלוש קבוצות', XL('מהות הבקשה')),
        ('מצב בקשה', 'ממתינה, בזמן, חריגה, או חריגה גדולה (יותר מפי 1.5 מהיעד)', '=LET(d,[[ימים עד החלטה]]2,y,IF([[מסלול רישוי]]2="רישוי רגיל",270,45),IF(d="","ממתינה",IF(d>1.5*y,"חריגה גדולה",IF(d>y,"חריגה","בזמן"))))')],
  map=('מהות הבקשה', [('בנייה חדשה', 'בנייה גדולה'), ('התחדשות עירונית - תמ"א 38/פינוי בינוי', 'בנייה גדולה'), ('שינוי שימוש', 'בנייה גדולה'), ('תוספת בנייה למגורים', 'תוספות'), ('בריכת שחייה', 'תוספות'), ('פרגולה / סככה', 'עבודות קלות'), ('חיפוי ושינויים בחזית', 'עבודות קלות')]),
  p=[('ממוצע ימים עד החלטה לפי מסלול', 'מסלול רישוי', 'avg', 'ימים עד החלטה'), ('בקשות שחרגו מהיעד לפי שכונה', FIX, 'yes', 'חרג מהיעד'), ('בקשות שהוגשו לפי חודש', 'חודש הגשה', 'count', None)],
  kpis=[('מספר הבקשות', '=COUNTA([[מס\' בקשה]]:[[מס\' בקשה]])-1'), ('ממוצע ימים עד החלטה', '=AVERAGE([[ימים עד החלטה]]:[[ימים עד החלטה]])'), ('בקשות שחרגו מהיעד', '=COUNTIF([[חרג מהיעד]]:[[חרג מהיעד]],"כן")')],
  filters=['מסלול רישוי', FIX, 'קבוצת מהות'], charts=['ממוצע ימים עד החלטה לפי מסלול', 'בקשות שהוגשו לפי חודש']),

'06': dict(short='ארנונה', topic='חיובי הארנונה לשנת 2026', role='גזבר/ית העירייה', who='כספים וגבייה', key=0,
  ask='איפה מרוכז החוב, ולפי איזה סיווג ואמצעי תשלום?',
  q3=[('כמה נכסים מסווגים כמגורים?', '=COUNTIF([[סיווג נכס]]:[[סיווג נכס]],"מגורים")'),
      ('מה סך יתרת החוב?', '=SUM([[יתרת חוב]]:[[יתרת חוב]])'),
      ('כמה נכסים משלמים בהוראת קבע?', '=COUNTIF([[אמצעי תשלום]]:[[אמצעי תשלום]],"הוראת קבע")')],
  mid=('מה סך יתרת החוב בשכונת מוריה?', '=SUMIF([[שכונה]]:[[שכונה]],"מוריה",[[יתרת חוב]]:[[יתרת חוב]])'),
  adv='מה שיעור הגבייה לפי אמצעי תשלום?',
  cols=[('אחוז ששולם', 'מה ששולם עד כה חלקי הסכום לתשלום', '=IF(OR([[לתשלום]]2="",[[לתשלום]]2=0),"",[[שולם עד כה]]2/[[לתשלום]]2)'),
        ('יש הנחה', 'כן אם לנכס יש הנחה כלשהי', '=IF([[סוג הנחה]]2="","",IF([[סוג הנחה]]2="ללא","לא","כן"))'),
        ('יש חוב', 'כן אם יתרת החוב גדולה מאפס', '=IF([[יתרת חוב]]2="","",IF([[יתרת חוב]]2>0,"כן","לא"))'),
        ('מדרגת חוב', 'שלוש מדרגות: ללא חוב, עד 1,000, מעל 1,000', '=IF([[יתרת חוב]]2="","",IF([[יתרת חוב]]2=0,"ללא חוב",IF([[יתרת חוב]]2<=1000,"עד 1,000","מעל 1,000")))'),
        ('הנחה וחוב גבוה', 'כן אם לנכס יש הנחה וגם חוב מעל 1,000 ₪', '=LET(h,[[סוג הנחה]]2,d,[[יתרת חוב]]2,IF(OR(h="",d=""),"",IF(AND(h<>"ללא",d>1000),"כן","לא")))')],
  p=[('יתרת החוב לפי סיווג נכס', 'סיווג נכס', 'sum', 'יתרת חוב'), ('יתרת החוב לפי אמצעי תשלום', 'אמצעי תשלום', 'sum', 'יתרת חוב'), ('יתרת החוב לפי שכונה', FIX, 'sum', 'יתרת חוב')],
  kpis=[('סך יתרת החוב (₪)', '=SUM([[יתרת חוב]]:[[יתרת חוב]])'), ('נכסים עם חוב', '=COUNTIF([[יש חוב]]:[[יש חוב]],"כן")'), ('שיעור הגבייה', '=SUM([[שולם עד כה]]:[[שולם עד כה]])/SUM([[לתשלום]]:[[לתשלום]])')],
  filters=['סיווג נכס', FIX, 'אמצעי תשלום'], charts=['יתרת החוב לפי סיווג נכס', 'יתרת החוב לפי שכונה']),

'07': dict(short='אירועים', topic='אירועי התרבות והספורט', role='מנהל/ת אגף התרבות', who='דוברות, תרבות ואירועים', key=0,
  ask='אילו אירועים מתמלאים, ואילו מפסידים?',
  q3=[('כמה הרצאות התקיימו?', '=COUNTIF([[סוג אירוע]]:[[סוג אירוע]],"הרצאה")'),
      ('כמה אנשים נכחו בסך הכול?', '=SUM([[נכחו בפועל]]:[[נכחו בפועל]])'),
      ('מה ציון המשוב הממוצע?', '=AVERAGE([[ציון משוב]]:[[ציון משוב]])')],
  mid=('מה סך ההכנסות מאירועים בהיכל התרבות?', '=SUMIF([[מקום]]:[[מקום]],"היכל התרבות",[[הכנסות]]:[[הכנסות]])'),
  adv='אילו סוגי אירועים מתמלאים, ואילו עולים יותר ממה שהם מכניסים?',
  cols=[('אחוז תפוסה', 'הנוכחים בפועל חלקי הקיבולת', '=IF(OR([[קיבולת]]2="",[[קיבולת]]2=0),"",[[נכחו בפועל]]2/[[קיבולת]]2)'),
        ('חודש', 'שנה וחודש מתוך תאריך האירוע', MONTH('תאריך')),
        ('התמלא', 'כן אם התפוסה 75% או יותר', '=IF(OR([[קיבולת]]2="",[[קיבולת]]2=0),"",IF([[נכחו בפועל]]2/[[קיבולת]]2>=0.75,"כן","לא"))'),
        ('תחום', 'סוגי האירועים בשלוש קבוצות: תרבות, ספורט וקהילה', XL('סוג אירוע')),
        ('מצב אירוע', 'תפוסה (75% ומעלה) ומאזן (הכנסות מול תקציב) יחד', '=LET(g,[[קיבולת]]2,b,[[תקציב האירוע]]2,IF(OR(g="",g=0,b=""),"",LET(o,[[נכחו בפועל]]2/g,p,[[הכנסות]]2-b,IF(AND(o>=0.75,p>=0),"מלא ומאוזן",IF(o>=0.75,"מלא ומפסיד",IF(p>=0,"לא מלא ומאוזן","לא מלא ומפסיד"))))))')],
  map=('סוג אירוע', [('מופע / הופעה', 'תרבות'), ('הצגת ילדים', 'תרבות'), ('הרצאה', 'תרבות'), ('סדנה', 'תרבות'), ('מרוץ / אירוע ספורט', 'ספורט'), ('ליגת נוער', 'ספורט'), ('אירוע קהילתי בפארק', 'קהילה'), ('יריד', 'קהילה'), ('טקס עירוני', 'קהילה')]),
  p=[('אחוז תפוסה ממוצע לפי סוג אירוע', 'סוג אירוע', 'avg', 'אחוז תפוסה'), ('הכנסות לפי מקום', 'מקום', 'sum', 'הכנסות'), ('נוכחים לפי חודש', 'חודש', 'sum', 'נכחו בפועל')],
  kpis=[('מספר האירועים', '=COUNTA([[מס\' אירוע]]:[[מס\' אירוע]])-1'), ('סך הנוכחים', '=SUM([[נכחו בפועל]]:[[נכחו בפועל]])'), ('הכנסות פחות תקציב (₪)', '=SUM([[הכנסות]]:[[הכנסות]])-SUM([[תקציב האירוע]]:[[תקציב האירוע]])')],
  filters=['סוג אירוע', 'מקום', 'קהל יעד'], charts=['אחוז תפוסה ממוצע לפי סוג אירוע', 'נוכחים לפי חודש']),

'08': dict(short='פיקוח', topic='דוחות הפיקוח העירוני', role='מנהל/ת הפיקוח העירוני', who='פיקוח ואכיפה', key=0,
  ask='מה שיעור הגבייה, לפי סוג עבירה ולפי שכונה?',
  q3=[('כמה דוחות ניתנו על "חניה במקום אסור"?', '=COUNTIF([[תיאור עבירה]]:[[תיאור עבירה]],"חניה במקום אסור")'),
      ('מה סך הקנסות שניתנו?', '=SUM([[סכום קנס]]:[[סכום קנס]])'),
      ('כמה דוחות שולמו?', '=COUNTIF([[סטטוס]]:[[סטטוס]],"שולם")')],
  mid=('מה סך הקנסות בדוחות ששולמו?', '=SUMIF([[סטטוס]]:[[סטטוס]],"שולם",[[סכום קנס]]:[[סכום קנס]])'),
  adv='מה שיעור התשלום לפי סוג עבירה, ואיפה הוא נמוך במיוחד?',
  cols=[('ימים עד תשלום', 'תאריך התשלום פחות תאריך העבירה', '=IF(OR([[תאריך עבירה]]2="",[[תאריך תשלום]]2=""),"",[[תאריך תשלום]]2-[[תאריך עבירה]]2)'),
        ('חודש עבירה', 'שנה וחודש מתוך תאריך העבירה', MONTH('תאריך עבירה')),
        ('שולם', 'כן אם הסטטוס הוא "שולם"', '=IF([[סטטוס]]2="","",IF([[סטטוס]]2="שולם","כן","לא"))'),
        ('קבוצת עבירה', 'העבירות בארבע קבוצות: חניה, כלבים, פסולת וניקיון, רעש', XL('תיאור עבירה')),
        ('מצב גבייה', 'סטטוס הדוח ומהירות התשלום יחד (איחור: יותר מ-60 יום)', '=LET(s,[[סטטוס]]2,a,[[תאריך עבירה]]2,b,[[תאריך תשלום]]2,IF(s="","",IF(s="שולם",IF(OR(a="",b=""),"שולם",IF(b-a>60,"שולם באיחור","שולם בזמן")),IF(OR(s="בוטל",s="בקשה לביטול"),"ביטול","פתוח"))))')],
  map=('תיאור עבירה', [('חניה במקום אסור', 'חניה'), ('חניה ללא תשלום', 'חניה'), ('חניה על מדרכה', 'חניה'), ('חניה בחניית נכים', 'חניה'), ('כלב ללא רצועה', 'כלבים'), ('אי איסוף צואת כלב', 'כלבים'), ('השלכת פסולת ברשות הרבים', 'פסולת וניקיון'), ('פסולת בניין', 'פסולת וניקיון'), ('מטרד רעש', 'רעש')]),
  p=[('סך הקנסות לפי סוג עבירה', 'תיאור עבירה', 'sum', 'סכום קנס'), ('דוחות שלא שולמו לפי שכונה', FIX, 'no', 'שולם'), ('דוחות לפי חודש', 'חודש עבירה', 'count', None)],
  kpis=[('מספר הדוחות', '=COUNTA([[מס\' דוח]]:[[מס\' דוח]])-1'), ('סך הקנסות (₪)', '=SUM([[סכום קנס]]:[[סכום קנס]])'), ('דוחות ששולמו', '=COUNTIF([[שולם]]:[[שולם]],"כן")')],
  filters=['קבוצת עבירה', FIX, 'סטטוס'], charts=['סך הקנסות לפי סוג עבירה', 'דוחות לפי חודש']),

'09': dict(short='משאבי אנוש', topic='עובדי העירייה', role='מנהל/ת משאבי אנוש', who='משאבי אנוש', key=0,
  ask='איפה עומס השעות הנוספות וימי המחלה?',
  q3=[('כמה עובדים יש באגף חינוך?', '=COUNTIF([[אגף]]:[[אגף]],"חינוך")'),
      ('כמה ימי מחלה יש בממוצע לעובד?', '=AVERAGE([[ימי מחלה]]:[[ימי מחלה]])'),
      ('כמה שעות הדרכה ניתנו בסך הכול?', '=SUM([[שעות הדרכה]]:[[שעות הדרכה]])')],
  mid=('כמה שעות נוספות יש בממוצע באגף שפ"ע?', '=AVERAGEIF([[אגף]]:[[אגף]],"שפ""ע",[[שעות נוספות]]:[[שעות נוספות]])'),
  adv='באילו אגפים יש גם הרבה שעות נוספות וגם הרבה ימי מחלה?',
  cols=[('שעות נוספות בשנה', 'הממוצע החודשי כפול 12', '=IF([[שעות נוספות]]2="","",[[שעות נוספות]]2*12)'),
        ('שנת תחילת עבודה', 'השנה מתוך תאריך תחילת העבודה', '=IF([[תאריך תחילת עבודה]]2="","",YEAR([[תאריך תחילת עבודה]]2))'),
        ('עומס גבוה', 'כן אם יש 15 שעות נוספות בחודש או יותר, וגם 8 ימי מחלה או יותר', '=IF(OR([[שעות נוספות]]2="",[[ימי מחלה]]2=""),"",IF(AND([[שעות נוספות]]2>=15,[[ימי מחלה]]2>=8),"כן","לא"))'),
        ('קבוצת ותק', 'שלוש קבוצות: עד 5 שנים, 5 עד 15, מעל 15', '=IF([[ותק]]2="","",IF([[ותק]]2<5,"עד 5 שנים",IF([[ותק]]2<=15,"5 עד 15 שנים","מעל 15 שנים")))'),
        ('פרופיל עובד', 'שעות נוספות, מחלה והדרכה יחד', '=LET(o,[[שעות נוספות]]2,s,[[ימי מחלה]]2,t,[[שעות הדרכה]]2,IF(OR(o="",s=""),"",IF(AND(o>=15,s>=8),"עומס ומחלה",IF(o>=15,"עומס",IF(t<10,"מעט הדרכה","רגיל")))))')],
  p=[('ממוצע שעות נוספות לפי אגף', 'אגף', 'avg', 'שעות נוספות - ממוצע'), ('ממוצע ימי מחלה לפי אגף', 'אגף', 'avg', 'ימי מחלה'), ('ממוצע ציון הערכה לפי קבוצת ותק', 'קבוצת ותק', 'avg', 'ציון הערכה')],
  kpis=[('מספר העובדים', '=COUNTA([[מס\' עובד]]:[[מס\' עובד]])-1'), ('ממוצע שעות נוספות בחודש', '=AVERAGE([[שעות נוספות - ממוצע]]:[[שעות נוספות - ממוצע]])'), ('עובדים בעומס גבוה', '=COUNTIF([[עומס גבוה]]:[[עומס גבוה]],"כן")')],
  filters=['אגף', 'דירוג', 'קבוצת ותק'], charts=['ממוצע שעות נוספות לפי אגף', 'ממוצע ימי מחלה לפי אגף']),

'10': dict(short='גנים ציבוריים', topic='תחזוקת הגנים הציבוריים', role='מנהל/ת מחלקת גנים ונוף', who='גנים ונוף', key=None,
  ask='כמה מים צורכים לדונם, ואילו גינות חריגות?',
  q3=[('מה סך צריכת המים, במ"ק?', '=SUM([[צריכת מים]]:[[צריכת מים]])'),
      ('כמה תקלות השקיה היו בסך הכול?', '=SUM([[תקלות השקיה]]:[[תקלות השקיה]])'),
      ('כמה שורות יש ל"פארק מוריה"?', '=COUNTIF([[שם הגינה]]:[[שם הגינה]],"פארק מוריה")')],
  mid=('מה סך צריכת המים בשכונת מוריה?', '=SUMIF([[שכונה]]:[[שכונה]],"מוריה",[[צריכת מים]]:[[צריכת מים]])'),
  adv='באילו גינות צריכת המים לדונם חריגה, והאם יש שם גם תקלות השקיה?',
  cols=[('מים לדונם (מ"ק)', 'צריכת המים חלקי שטח הגינה', '=IF(OR([[שטח]]2="",[[שטח]]2=0),"",[[צריכת מים]]2/[[שטח]]2)'),
        ('שנה', 'השנה מתוך עמודת החודש', '=IF([[חודש]]2="","",RIGHT([[חודש]]2,4))'),
        ('צריכה חריגה', 'כן אם הצריכה מעל 170 מ"ק לדונם', '=IF(OR([[שטח]]2="",[[שטח]]2=0),"",IF([[צריכת מים]]2/[[שטח]]2>170,"כן","לא"))'),
        ('סוג שטח', 'פארק או גינה, לפי תחילת השם', '=IF([[שם הגינה]]2="","",IF(LEFT([[שם הגינה]]2,4)="פארק","פארק","גינה"))'),
        ('מצב גינה', 'צריכת מים (מעל 170 מ"ק לדונם) ותקלות השקיה (3 ומעלה) יחד', '=LET(a,[[שטח]]2,t,[[תקלות השקיה]]2,IF(OR(a="",a=0),"",LET(w,[[צריכת מים]]2/a,IF(AND(w>170,t>=3),"צריכה חריגה ותקלות",IF(w>170,"צריכה חריגה",IF(t>=3,"תקלות","תקין"))))))')],
  p=[('ממוצע מים לדונם לפי גינה', 'שם הגינה', 'avg', 'מים לדונם'), ('צריכת המים לפי שכונה', FIX, 'sum', 'צריכת מים'), ('צריכת המים לפי שנה', 'שנה', 'sum', 'צריכת מים')],
  kpis=[('סך צריכת המים (מ"ק)', '=SUM([[צריכת מים]]:[[צריכת מים]])'), ('תקלות השקיה', '=SUM([[תקלות השקיה]]:[[תקלות השקיה]])'), ('שורות עם צריכה חריגה', '=COUNTIF([[צריכה חריגה]]:[[צריכה חריגה]],"כן")')],
  filters=['שם הגינה', FIX, 'שנה'], charts=['ממוצע מים לדונם לפי גינה', 'צריכת המים לפי שנה']),
}

HEAD_FONT = Font(bold=True, color='FFFFFF')
HEAD_FILL = PatternFill('solid', fgColor='1F4E79')
NEW_FILL = PatternFill('solid', fgColor='B4502F')


def blank(v):
    return v is None or (isinstance(v, str) and v.strip() == '')


def kind(v):
    if isinstance(v, bool): return 'str'
    if isinstance(v, (int, float)): return 'num'
    if isinstance(v, (datetime.datetime, datetime.date)): return 'date'
    return 'str'


def load(n):
    ws = openpyxl.load_workbook(os.path.join(SRC, n + '.xlsx')).worksheets[0]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    fmts = [c.number_format for c in ws[2]]
    return rows[0], rows[1:], fmts


def clean(head, rows, key):
    """-> clean rows + what was fixed. The neighbourhood column is left as typed; the fix goes to a new column."""
    hood = head.index(HOOD) if HOOD in head else None
    major = []
    for i in range(len(head)):
        c = collections.Counter(kind(r[i]) for r in rows if not blank(r[i]))
        major.append(c.most_common(1)[0][0] if c else 'str')
    st = dict(raw=len(rows), empty=0, dups=0, text=0, missing=0)
    out, seen = [], set()
    for r in rows:
        if all(blank(v) for v in r[1:]):
            st['empty'] += 1; continue
        r = list(r)
        for i, v in enumerate(r):
            if blank(v): r[i] = None; continue
            if not isinstance(v, str) or i == hood: continue
            s = v.strip()
            if major[i] == 'num':
                st['text'] += 1
                try:
                    f = float(s); r[i] = None if f != f else (int(f) if f == int(f) else f)
                except ValueError: r[i] = None
            elif major[i] == 'date':
                st['text'] += 1
                m = re.fullmatch(r'(\d\d)\.(\d\d)\.(\d\d)', s)
                r[i] = datetime.datetime(2000 + int(m.group(3)), int(m.group(2)), int(m.group(1))) if m else None
            else: r[i] = s
        k = r[key] if key is not None else tuple(r)
        if k in seen: st['dups'] += 1; continue
        seen.add(k); out.append(r)
    st['clean'] = len(out)
    st['missing'] = sum(1 for r in out for v in r if v is None)
    spell = None
    if hood is not None:
        c = collections.Counter(r[hood] for r in out if r[hood] is not None)
        fix = lambda v: HOODS.get(v.strip(), v.strip())
        spell = [(v, fix(v), n) for v, n in sorted(c.items(), key=lambda t: (fix(t[0]), -t[1]))]
        st['spell_before'] = len(c)
        st['spell_after'] = len(set(fix(v) for v in c))
        for r in out: r.append(fix(r[hood]) if r[hood] is not None else None)
    return out, st, spell


def sheet(wb, title, head, rows, fmts=None, new_from=None, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title; ws.sheet_view.rightToLeft = True
    ws.append(head)
    for r in rows: ws.append(r)
    for i, c in enumerate(ws[1]):
        c.font = HEAD_FONT; c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.fill = NEW_FILL if new_from is not None and i >= new_from else HEAD_FILL
        ws.column_dimensions[L(i + 1)].width = max(14, min(34, len(str(head[i])) + 6))
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = 'A2'
    if fmts:
        for i, f in enumerate(fmts):
            if f and f != 'General':
                for row in ws.iter_rows(min_row=2, min_col=i + 1, max_col=i + 1):
                    row[0].number_format = f
    return ws


def letters(head):
    def sub(f):
        def one(m):
            p = m.group(1)
            hit = [i for i, h in enumerate(head) if h == p] or [i for i, h in enumerate(head) if h.startswith(p)]
            if not hit: raise KeyError('no column %r' % p)
            return L(hit[0] + 1)
        return re.sub(r'\[\[(.+?)\]\]', one, f)
    return sub


def is_err(v):
    return isinstance(v, int) and v < -2146826000


def main():
    import win32com.client
    os.makedirs(OUT, exist_ok=True); os.makedirs(TMP, exist_ok=True)
    only = sys.argv[1:]
    xl = win32com.client.DispatchEx('Excel.Application'); xl.Visible = False; xl.DisplayAlerts = False
    facts = {}
    try:
        for n, F in FILES.items():
            if only and n not in only: continue
            head, rows, fmts = load(n)
            crows, st, spell = clean(head, rows, F['key'])
            chead = head + ([FIX] if spell else [])
            base = len(chead)
            # ---- clean file
            wb = openpyxl.Workbook()
            sheet(wb, 'נתונים', chead, crows, fmts, new_from=len(head) if spell else None, first=True)
            if spell: sheet(wb, 'כתיבים', ['כפי שנכתב', 'שם תקני', 'מספר שורות'], [list(t) for t in spell])
            wb.save(os.path.join(OUT, n + '-clean.xlsx'))
            # ---- rich file: Excel fills the formulas and turns them into values
            rhead = chead + [c[0] for c in F['cols']]
            sub = letters(rhead)
            cols = [dict(name=c[0], what=c[1], f=sub(c[2])) for c in F['cols']]
            wb = openpyxl.Workbook()
            ws = sheet(wb, 'נתונים', rhead, crows, fmts, new_from=base, first=True)
            for ci, c in enumerate(cols):
                if 'אחוז' in c['name']:
                    for row in ws.iter_rows(min_row=2, min_col=base + 1 + ci, max_col=base + 1 + ci): row[0].number_format = '0%'
            if spell: sheet(wb, 'כתיבים', ['כפי שנכתב', 'שם תקני', 'מספר שורות'], [list(t) for t in spell])
            if F.get('map'): sheet(wb, 'מיפוי', [F['map'][0], 'קבוצה'], [list(t) for t in F['map'][1]])
            sheet(wb, 'הסבר', ['עמודה חדשה', 'מה יש בה', 'הנוסחה בשורה 2'], [[c['name'], c['what'], "'" + c['f']] for c in cols])
            tmp = os.path.join(TMP, n + '-rich.xlsx'); wb.save(tmp)
            final = os.path.join(OUT, n + '-rich.xlsx')
            if os.path.exists(final): os.remove(final)
            xwb = xl.Workbooks.Open(tmp); xws = xwb.Worksheets('נתונים')
            for ci, c in enumerate(cols):     # one formula for the whole column: Excel shifts the row number itself
                xws.Range(xws.Cells(2, base + 1 + ci), xws.Cells(len(crows) + 1, base + 1 + ci)).Formula2 = c['f']
            xl.CalculateFull()
            rng = xws.Range(xws.Cells(2, base + 1), xws.Cells(len(crows) + 1, base + len(cols)))
            vals = rng.Value
            for ci, c in enumerate(cols):
                bad = [r[ci] for r in vals if is_err(r[ci])]
                if bad: raise SystemExit('%s: formula for %s gives errors (%d rows): %s' % (n, c['name'], len(bad), c['f']))
                c['filled'] = sum(1 for r in vals if r[ci] not in (None, ''))
            rng.Value = vals
            kpis = []
            for name, f in F['kpis']:
                f = sub(f); v = xws.Evaluate(f)
                if is_err(v): raise SystemExit('%s: kpi %s fails: %s' % (n, name, f))
                kpis.append(dict(name=name, f=f, v=v))
            xwb.SaveAs(final, 51); xwb.Close(False)
            # ---- round 1 answers: on the raw file, as the participant has it
            xwb = xl.Workbooks.Open(os.path.join(SRC, n + '.xlsx'), ReadOnly=True); xws = xwb.Worksheets(1)
            def ask(q, f):
                f = sub(f); v = xws.Evaluate(f)
                if is_err(v): raise SystemExit('%s: %s fails: %s' % (n, q, f))
                return dict(q=q, f=f, v=v)
            q3 = [ask(q, f) for q, f in F['q3']]; mid = ask(*F['mid'])
            xwb.Close(False)
            # ---- cross-tabs from the finished rich file
            rws = openpyxl.load_workbook(final).worksheets[0]
            data = [list(r) for r in rws.iter_rows(values_only=True)]
            rh, rd = data[0], data[1:]
            col = lambda p: ([i for i, h in enumerate(rh) if h == p] or [i for i, h in enumerate(rh) if h.startswith(p)])[0]
            piv = []
            for title, rowf, agg, valf in F['p']:
                ri = col(rowf); vi = col(valf) if valf else None; g = collections.defaultdict(list)
                for r in rd:
                    if r[ri] in (None, ''): continue
                    g[str(r[ri])].append(r[vi] if vi is not None else 1)
                res = []
                for k, vs in g.items():
                    nums = [v for v in vs if isinstance(v, (int, float))]
                    if agg == 'count': v = len(vs)
                    elif agg == 'sum': v = sum(nums)
                    elif agg == 'avg': v = sum(nums) / len(nums) if nums else None
                    else: v = sum(1 for x in vs if x == ('כן' if agg == 'yes' else 'לא'))
                    if v is not None: res.append([k, round(v, 2)])
                timeline = rh[ri] in ('שנה', 'שנת לימודים') or rh[ri].startswith('חודש')
                res.sort(key=(lambda t: t[0]) if timeline else (lambda t: -t[1]))
                piv.append(dict(title=title, rows=rh[ri], agg=agg, val=rh[vi] if vi is not None else None, n=len(res),
                                top=res if timeline and len(res) <= 6 else res[:5], timeline=timeline))
            facts[n] = dict(n=n, short=F['short'], topic=F['topic'], role=F['role'], who=F['who'], ask=F['ask'], id=head[0],
                            head=head, hood=bool(spell), letters={h: L(i + 1) for i, h in enumerate(rhead)},
                            st=st, spell=[list(t) for t in spell] if spell else None,
                            q3=q3, mid=mid, adv=F['adv'], cols=cols, map=dict(src=F['map'][0], pairs=F['map'][1]) if F.get('map') else None,
                            piv=piv, kpis=kpis, filters=F['filters'], charts=F['charts'])
            print(n, st, [c['filled'] for c in cols])
    finally:
        xl.Quit()
    if only:
        print('partial run: files.js not written'); return
    js = '/* generated by tools/dataai_build.py; do not edit by hand */\nwindow.DATAAI_FILES=' + json.dumps(facts, ensure_ascii=False, separators=(',', ':'), default=str) + ';\n'
    open(os.path.join(ROOT, 'data-ai', 'files.js'), 'w', encoding='utf-8', newline='\n').write(js)
    print('files.js', len(js))


if __name__ == '__main__':
    main()
