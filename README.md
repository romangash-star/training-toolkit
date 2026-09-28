# ההכשרות הדיגיטליות — ארגז הכלים

האתר המלווה את ההכשרות של רומן גרינשטיין לרשויות מקומיות. אתר סטטי (HTML/CSS/JS בלבד), בלי build.

| נתיב | מה יש שם |
|---|---|
| `/` | דף הבית: בחירת הכשרה |
| `/ai/` | הכשרת בינה מלאכותית |
| `/maazim/` | מאיצים דיגיטליים |
| `/data/` | קבלת החלטות מבוססת נתונים (3 ימים) |

כלים בתוך `data/`:
- חדשים: `security/`, `question/`, `analysis/`, `chart/`, `prompts/`, `case/`, `cards/`. סגנון משותף ב-`data/assets/dd.css` ו-`dd.js`.
- העתקים של כלי מאיצים: `role-worksheet/`, `problem-tree/`, `challenge/`, `prompt/`, `insight/`, `prototype/`, `pitch/`. נשמרים ב-localStorage במפתחות עם קידומת `dd_`, כך שלא מתערבבים עם מאיצים.

כל הקישורים יחסיים, כך שהאתר עובד גם ב-GitHub Pages תחת `user.github.io/repo/` וגם בדומיין משלכם.

קבצי מקרה הבוחן (אלון-ים) נכנסים ל-`data/case/files/`, ומעדכנים את `href` במערך `FILES` שב-`data/case/index.html`.

הרצה מקומית: `python -m http.server` ופתיחת http://localhost:8000
