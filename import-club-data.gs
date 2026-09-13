/*************************************************************
 *  ИМПОРТ ДАННЫХ КРУЖКА РОБОТОТЕХНИКИ
 *
 *  Данные взяты из файла robotics_club.xlsx (листы «spisok 26-27»,
 *  «Жарыстан орын алған оқушылар», «Robotics жарыстар 26-27», «Annual plan»).
 *  ИИН, даты рождения и номера документов НЕ переносились.
 *
 *  Как пользоваться: добавьте этот файл в тот же проект Apps Script,
 *  что и Code.gs, и один раз запустите функцию importClubData().
 *  Повторный запуск ничего не испортит: существующие записи не дублируются.
 *************************************************************/

var IMPORT_SINCE = '2026-09-01';        // с какой даты считать посещаемость текущего состава
var IMPORT_WRITE_PIN_SHEET = true;   // выписать коды учеников на отдельный лист

/* имя, класс, команда, категория, в составе сейчас (1) или в архиве (0) */
var IMPORT_STUDENTS = [
  ['Базарбай Өміржан', '9А', 'Aqsonix', 'FLL', 1],
  ['Ғабит Абубакр', '9А', 'Aqsonix', 'FLL', 1],
  ['Танкиев Ескендір', '9А', 'Aqsonix', 'FLL', 1],
  ['Хажиолиев Төлеген', '9А', 'Aqsonix', 'FLL', 1],
  ['Турдыбеков Алихан', '9А', 'Aqsonix', 'FLL', 1],
  ['Ахмет Султан', '9Б', '', 'FLL', 1],
  ['Елеусіз Бекарыс', '9Б', '', 'FLL', 1],
  ['Азмуханов Алим', '10А', '', '', 1],
  ['Али Ахмет', '10А', '', 'FTC', 1],
  ['Бисенгалиев Ислам', '10А', '', 'FTC', 1],
  ['Мұрзақан Мадияр', '10А', '', 'FTC', 1],
  ['Нығмет Әбдірасул', '10А', '', 'FTC', 1],
  ['Әменов Омар', '10А', '', 'FTC', 1],
  ['Тлеумбетов Ерамир', '10А', '', 'FTC', 1],
  ['Тулегенов Расул', '10Б', '', 'FTC', 1],
  ['Таяу Сүлеймен', '10Б', '', 'FTC', 1],
  ['Майор Әсет', '10Б', '', 'FTC', 1],
  ['Нагашбаев Амир', '10Б', '', 'FTC', 1],
  ['Ақылбек Сардар', '9А', 'GAMBIT', 'FTC', 1],
  ['Аманғос Омарәлім', '9А', 'GAMBIT', 'FTC', 1],
  ['Еркін Диас', '10А', 'GAMBIT', 'FTC', 1],
  ['Саматов Ерасыл', '10Б', 'GAMBIT', 'FTC', 1],
  ['Есберген Қайсар', '10Б', 'GAMBIT', 'FTC', 1],
  ['Қазанбаев Сүлеймен', '11А', 'GAMBIT', 'FTC', 1],
  ['Қыдырбай Нұрислам', '11А', 'GAMBIT', 'FTC', 1],
  ['Мухтубаев Али', '11А', 'GAMBIT', 'FTC', 1],
  ['Санжар Жұмабек', '11А', 'GAMBIT', 'FTC', 1],
  ['Рат Мадияр', '11А', '', 'FTC', 1],
  ['Барлық Исламбек', '', '', '', 1],
  ['Алпысбай Аслан', '11А', 'GAMBIT', '', 0],
  ['Аманбай Исабек', '11А', 'GAMBIT', '', 0],
  ['Давлетов Есентай', '10А', 'GAMBIT jr', '', 0],
  ['Жалғасбай Әділ', '10Б', 'GAMBIT', '', 0],
  ['Жандосұлы Жансерік', '8A', 'Aqsonix', '', 0],
  ['Наурызғали Бекнар', '11А', 'GAMBIT', '', 0],
  ['Нурмухамедов Асмир', '9A', 'BIL innovators', '', 0],
  ['Нұрмұхан Нұрали', '10А', 'GAMBIT jr', '', 0],
  ['Оспанов Даурен', '11А', 'GAMBIT', '', 0],
  ['Сағынғали Абдуссаттар', '10Б', 'GAMBIT', '', 0],
  ['Серіков Темірлан', '8A', 'Aqsonix', '', 0],
  ['Серікқали Ғани', '11А', 'GAMBIT', '', 0],
  ['Суйесинов Алмаз', '10А', 'GAMBIT jr', '', 0],
  ['Теңелғали Бекарыстан', '10А', 'GAMBIT jr', '', 0],
  ['Төребай Қасымхан', '11А', 'GAMBIT', '', 0],
  ['Умаров Қайсар', '11А', 'GAMBIT', '', 0],
  ['Шүйншбай Ақтілек', '10А', 'GAMBIT', '', 0],
];

/* имя, соревнование, уровень, результат, как записано в таблице */
var IMPORT_AWARDS = [
  ['Ақылбек Сардар', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Ақылбек Сардар', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Ақылбек Сардар', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Ақылбек Сардар', 'INFOMATRIX', 'international', 'bronze', '3rd place'],
  ['Ақылбек Сардар', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Ақылбек Сардар', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Аманғос Омарәлім', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Аманғос Омарәлім', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Аманғос Омарәлім', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Аманғос Омарәлім', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Аманғос Омарәлім', 'INFOMATRIX', 'international', 'bronze', '3rd place'],
  ['Аманғос Омарәлім', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Аманғос Омарәлім', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Базарбай Өміржан', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Базарбай Өміржан', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Базарбай Өміржан', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Базарбай Өміржан', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Базарбай Өміржан', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Базарбай Өміржан', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Ғабит Абубакр', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Ғабит Абубакр', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Ғабит Абубакр', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Ғабит Абубакр', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Ғабит Абубакр', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Жандосұлы Жансерік', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Жандосұлы Жансерік', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Жандосұлы Жансерік', 'Fibonacci Astana', 'republic', 'bronze', 'Judge award, 3 place, Turkey'],
  ['Жандосұлы Жансерік', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Жандосұлы Жансерік', 'INFOMATRIX', 'international', 'part', 'participant'],
  ['Серіков Темірлан', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Серіков Темірлан', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Танкиев Ескендір', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Танкиев Ескендір', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Танкиев Ескендір', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Танкиев Ескендір', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Хажиолиев Төлеген', 'First Aqtobe Regional', 'region', 'honor', 'Core Values Award Finalist'],
  ['Хажиолиев Төлеген', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Хажиолиев Төлеген', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Хажиолиев Төлеген', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Хажиолиев Төлеген', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Турдыбеков Алихан', 'Astana CAFC', 'republic', 'honor', 'Technical excellence award'],
  ['Турдыбеков Алихан', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Турдыбеков Алихан', 'WRO', 'international', 'silver', 'Robo sport 2nd place'],
  ['Ахмет Султан', 'Fibonacci Turkey', 'international', 'bronze', '3 rd place, judge award'],
  ['Елеусіз Бекарыс', 'Fibonacci Turkey', 'international', 'silver', '2nd place'],
  ['Али Ахмет', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Али Ахмет', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Али Ахмет', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Али Ахмет', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Али Ахмет', 'WRO', 'international', 'gold', '1st place'],
  ['Еркін Диас', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Еркін Диас', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Еркін Диас', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Еркін Диас', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Еркін Диас', 'WRO', 'international', 'gold', '1st place'],
  ['Бисенгалиев Ислам', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Бисенгалиев Ислам', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Бисенгалиев Ислам', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Бисенгалиев Ислам', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Бисенгалиев Ислам', 'WRO', 'international', 'honor', '4th place'],
  ['Мұрзақан Мадияр', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Мұрзақан Мадияр', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Мұрзақан Мадияр', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Мұрзақан Мадияр', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Мұрзақан Мадияр', 'WRO', 'international', 'honor', '4th place'],
  ['Нығмет Әбдірасул', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Нығмет Әбдірасул', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Нығмет Әбдірасул', 'Fibonacci Astana', 'republic', 'honor', 'Judge award x3, Turkey'],
  ['Нығмет Әбдірасул', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Нығмет Әбдірасул', 'WRO', 'international', 'honor', '4th place'],
  ['Нурмухамедов Асмир', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Нурмухамедов Асмир', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Нурмухамедов Асмир', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Нурмухамедов Асмир', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Әменов Омар', 'Fibonacci Astana', 'republic', 'part', 'Certificate for participation'],
  ['Әменов Омар', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Әменов Омар', 'WRO', 'international', 'gold', '1st place'],
  ['Таяу Сүлеймен', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Таяу Сүлеймен', 'Bishkek first Kyrgyzstan', 'international', 'honor', 'Robot Performance award Winner'],
  ['Таяу Сүлеймен', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Таяу Сүлеймен', 'WRO', 'international', 'silver', '2nd place'],
  ['Тулегенов Расул', 'First Aqtobe Regional', 'region', 'honor', 'Robot Performance award Winner'],
  ['Тулегенов Расул', 'Astana CAFC', 'republic', 'honor', 'Young innovators award'],
  ['Тлеумбетов Ерамир', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Тлеумбетов Ерамир', 'WRO', 'international', 'honor', '4th place'],
  ['Саматов Ерасыл', 'WRO', 'international', 'honor', '4th place'],
  ['Есберген Қайсар', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Есберген Қайсар', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Есберген Қайсар', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Есберген Қайсар', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Есберген Қайсар', 'WRO', 'international', 'silver', '2nd place'],
  ['Барлық Исламбек', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Барлық Исламбек', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Барлық Исламбек', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Барлық Исламбек', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Давлетов Есентай', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Давлетов Есентай', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Давлетов Есентай', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Давлетов Есентай', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Қазанбаев Сүлеймен', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Қазанбаев Сүлеймен', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Қазанбаев Сүлеймен', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Қыдырбай Нұрислам', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Қыдырбай Нұрислам', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Қыдырбай Нұрислам', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Қыдырбай Нұрислам', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Нұрмұхан Нұрали', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Нұрмұхан Нұрали', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Нұрмұхан Нұрали', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Суйесинов Алмаз', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Суйесинов Алмаз', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Суйесинов Алмаз', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Теңелғали Бекарыстан', 'First Aqtobe Regional', 'region', 'gold', 'Champion\'s award winner'],
  ['Теңелғали Бекарыстан', 'Batys First Mangistau', 'region', 'honor', 'Engineering excellence award'],
  ['Теңелғали Бекарыстан', 'Astana CAFC', 'republic', 'honor', 'Engineering excellence award, Quota to Bulgaria'],
  ['Рат Мадияр', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Рат Мадияр', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Рат Мадияр', 'INFOMATRIX', 'international', 'gold', '1st place'],
  ['Шүйншбай Ақтілек', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Мухтубаев Али', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Мухтубаев Али', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Мухтубаев Али', 'FLL Bulgaria', 'international', 'bronze', '3rd place'],
  ['Сағынғали Абдуссаттар', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Сағынғали Абдуссаттар', 'INFOMATRIX', 'international', 'gold', '1st place'],
  ['Жалғасбай Әділ', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Жалғасбай Әділ', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Жалғасбай Әділ', 'INFOMATRIX', 'international', 'gold', '1st place'],
  ['Наурызғали Бекнар', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Наурызғали Бекнар', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Аманбай Исабек', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Аманбай Исабек', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Оспанов Даурен', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Оспанов Даурен', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Серікқали Ғани', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Серікқали Ғани', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Умаров Қайсар', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Умаров Қайсар', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Төребай Қасымхан', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
  ['Төребай Қасымхан', 'Astana CAFC', 'republic', 'honor', 'Quota to America'],
  ['Алпысбай Аслан', 'Batys First Mangistau', 'region', 'honor', 'Sustain award winner'],
];

/* соревнование, когда, четверть, статус */
var IMPORT_COMPETITIONS = [
  ['Облыстық жоба тақырып ұсыну', '20-25 август', '1 term', 'Done'],
  ['Облыстық жоба жарысы', '10-11 ноябрь', '2 term', ''],
  ['Республикалық жоба жарысы', '25-27 февраль', '3 term', ''],
  ['Fibonacci Italy', 'November 27-5', '2 term', 'Loading'],
  ['First Aktobe Regional (FLL, FTC)', 'December 12-13', '2 term', ''],
  ['Batys First Aktobe (FLL, FTC)', 'January 16-17', '3 term', ''],
  ['Fibonacci Astana', '', '3 term', ''],
  ['Astana CAFC (FLL, FTC)', '', '3 term', ''],
  ['INFOMATRIX', 'March', '3 term', ''],
  ['World robotics olympiad (WRO)', 'April - May', '4 term', ''],
  ['VEX', '', '', ''],
  ['SeaPerch', '', '', ''],
  ['Drone симулятор', '', '', ''],
  ['Startup projects', '', '', ''],
  ['STEAM Baku', '', '', ''],
  ['Летный академия', '', '', ''],
];

/* месяц, неделя, номер недели, роль, задача */
var IMPORT_PLAN = [
  ['SEPTEMBER', 'I', '1', 'Команда жұмысы', 'Introduction, тренермен танысу'],
  ['SEPTEMBER', 'I', '1', 'Captain', 'Дайындыққа оқушылардың келгенін қадағалау'],
  ['SEPTEMBER', 'I', '1', 'Manager', 'Инста аккаунт ашылғанын күту'],
  ['SEPTEMBER', 'I', '1', 'Builder', 'pinterest/youtube idea izdeu'],
  ['SEPTEMBER', 'I', '1', 'Programmer', 'prime lessons'],
  ['SEPTEMBER', 'I', '1', 'SMM', 'Инстаграм аккаунт ашу'],
  ['SEPTEMBER', 'I', '1', 'Slider/Layer', 'Canva pro aккаунт ашу'],
  ['SEPTEMBER', 'II', '2', 'Команда жұмысы', 'Картамен танысу'],
  ['SEPTEMBER', 'II', '2', 'Captain', 'Қадағалау'],
  ['SEPTEMBER', 'II', '2', 'Manager', 'мектепте, сыныпта аккаут актуалность жинау'],
  ['SEPTEMBER', 'II', '2', 'Builder', 'Миссия қанша очко правило оқу'],
  ['SEPTEMBER', 'II', '2', 'Programmer', 'prime lessons'],
  ['SEPTEMBER', 'II', '2', 'SMM', 'Лого дайындау'],
  ['SEPTEMBER', 'II', '2', 'Slider/Layer', 'Canva сабақтарын қарау'],
  ['SEPTEMBER', 'III', '3', 'Команда жұмысы', 'Инновационный проект тақырыбымен танысу'],
  ['SEPTEMBER', 'III', '3', 'Captain', 'Қадағалау'],
  ['SEPTEMBER', 'III', '3', 'Manager', 'ағылшынша speaking дамыту'],
  ['SEPTEMBER', 'III', '3', 'Builder', 'robot building idea tauip құрастырып бастау'],
  ['SEPTEMBER', 'III', '3', 'Programmer', 'youtube'],
  ['SEPTEMBER', 'III', '3', 'SMM', 'Сторис салу'],
  ['SEPTEMBER', 'III', '3', 'Slider/Layer', 'Canva сабақтарын қарау'],
  ['SEPTEMBER', 'IV', '4', 'Команда жұмысы', 'Инновационный проект тақырыбын зерттеу'],
  ['SEPTEMBER', 'IV', '4', 'Captain', 'Қадағалау'],
  ['SEPTEMBER', 'IV', '4', 'Manager', 'FLL Қазақстан бойынша қанша команда бар білу'],
  ['SEPTEMBER', 'IV', '4', 'Builder', 'Насадка ойлап табу'],
  ['SEPTEMBER', 'IV', '4', 'Programmer', 'youtube'],
  ['SEPTEMBER', 'IV', '4', 'SMM', '2 пост салу'],
  ['SEPTEMBER', 'IV', '4', 'Slider/Layer', 'Canva сабақтарын қарау'],
  ['OCTOBER', 'I', '5', 'Команда жұмысы', 'Команданы First тіркеу'],
  ['OCTOBER', 'I', '5', 'Captain', 'Команданы First тіркеу'],
  ['OCTOBER', 'I', '5', 'Manager', '1 митинг жасау'],
  ['OCTOBER', 'I', '5', 'Builder', 'Mission 1-2-3'],
  ['OCTOBER', 'I', '5', 'Programmer', 'Mission 1-2-3'],
  ['OCTOBER', 'I', '5', 'SMM', '2 пост салу'],
  ['OCTOBER', 'I', '5', 'Slider/Layer', 'Power point сабақтарын қарау'],
  ['OCTOBER', 'II', '6', 'Команда жұмысы', 'Активити жасау'],
  ['OCTOBER', 'II', '6', 'Captain', 'Дұрыс орындалмай жатқан жұмысты өзі жасау'],
  ['OCTOBER', 'II', '6', 'Manager', '2 митинг Қазақстандағы мықты командалармен'],
  ['OCTOBER', 'II', '6', 'Builder', 'points көбейту'],
  ['OCTOBER', 'II', '6', 'Programmer', 'points көбейту'],
  ['OCTOBER', 'II', '6', 'SMM', 'Команда таныстыру видео салу'],
  ['OCTOBER', 'II', '6', 'Slider/Layer', 'Power point сабақтарын қарау'],
  ['OCTOBER', 'III', '7', 'Команда жұмысы', 'Robot game уақытқа жасау'],
  ['OCTOBER', 'III', '7', 'Captain', 'Активити жасау'],
  ['OCTOBER', 'III', '7', 'Manager', 'Археолог адаммен сөйлесу'],
  ['OCTOBER', 'III', '7', 'Builder', '100 points жеткізу'],
  ['OCTOBER', 'III', '7', 'Programmer', '100 points жеткізу'],
  ['OCTOBER', 'III', '7', 'SMM', '150 подписчик жинау'],
  ['OCTOBER', 'III', '7', 'Slider/Layer', 'Power point сабақтарын қарау'],
  ['OCTOBER', 'IV', '8', 'Команда жұмысы', 'Жарысқа тіркелу'],
  ['OCTOBER', 'IV', '8', 'Captain', 'Активити жасау'],
  ['OCTOBER', 'IV', '8', 'Manager', 'Музейге бару соны ұйымдастыру'],
  ['OCTOBER', 'IV', '8', 'Builder', '120 points жеткізу'],
  ['OCTOBER', 'IV', '8', 'Programmer', '120 points жеткізу'],
  ['OCTOBER', 'IV', '8', 'SMM', '2 пост 150 лайк жинау'],
  ['OCTOBER', 'IV', '8', 'Slider/Layer', 'Инновационный проект слайд бітіру'],
  ['NOVEMBER', 'I', '9', 'Команда жұмысы', 'Балалардың спич тексеру'],
  ['NOVEMBER', 'I', '9', 'Captain', 'Активити жасау'],
  ['NOVEMBER', 'I', '9', 'Manager', 'Мектепке облысына команданы реклама жасау'],
  ['NOVEMBER', 'I', '9', 'Builder', '200 points жеткізу'],
  ['NOVEMBER', 'I', '9', 'Programmer', '200 points жеткізу'],
  ['NOVEMBER', 'I', '9', 'SMM', 'Әр пост 100 лайктан аз болмау'],
  ['NOVEMBER', 'I', '9', 'Slider/Layer', 'Робот дизайн слайд жасау'],
  ['NOVEMBER', 'II', '10', 'Команда жұмысы', 'Жарыс алдындағы дайындық'],
  ['NOVEMBER', 'II', '10', 'Captain', 'Общий команда Robot game дайындық'],
  ['NOVEMBER', 'II', '10', 'Manager', 'Мектепке Ақтөбе облысына команданы реклама жасау'],
  ['NOVEMBER', 'II', '10', 'Builder', '250 points жеткізу'],
  ['NOVEMBER', 'II', '10', 'Programmer', '250 points жеткізу'],
  ['NOVEMBER', 'II', '10', 'SMM', 'Сторис салу'],
  ['NOVEMBER', 'II', '10', 'Slider/Layer', 'Журнал толтыру'],
  ['NOVEMBER', 'III', '11', 'Команда жұмысы', '2 video call, 2 рет музейге бару, Robot game 265 points, Инновационный проект макет дайын болу керек'],
  ['NOVEMBER', 'III', '11', 'Manager', 'Мектепке Ақтөбе облысына команданы реклама жасау'],
  ['NOVEMBER', 'III', '11', 'Builder', '265 points жеткізу'],
  ['NOVEMBER', 'III', '11', 'Programmer', '265 points жеткізу'],
  ['NOVEMBER', 'III', '11', 'SMM', 'Сторис салу'],
  ['NOVEMBER', 'III', '11', 'Slider/Layer', 'Бос адамдармен макет жасау'],
  ['NOVEMBER', 'IV', '12', 'Команда жұмысы', '3 video call, 2 рет музейге бару, Robot game 265 points, Инновационный проект макет, robot design slide дайын болу керек'],
  ['NOVEMBER', 'IV', '12', 'Manager', 'FLL November region jarys'],
  ['NOVEMBER', 'IV', '12', 'Builder', 'Проверьте робота на поле 10+ раз подряд на стабильность.'],
  ['NOVEMBER', 'IV', '12', 'SMM', 'Подготовьте единый стиль команды (футболки, логотип, плакат).'],
  ['DECEMBER', 'I', '13', 'Команда жұмысы', 'Спикерлердің сөйлеу мәнерін (intonation, gestures, smile) жаттықтыру.)'],
  ['DECEMBER', 'I', '13', 'Manager', 'Инновацияны жергілікті мысалмен байланыстыру (мысалы: Kazakhstan minerals, Smart irrigation system).'],
  ['DECEMBER', 'I', '13', 'Programmer', 'Роботтың қозғалысын тұрақты ету (PID calibration, gyro turning).'],
  ['DECEMBER', 'I', '13', 'Slider/Layer', 'PowerPoint-та дизайнды STEM стилінде әдемі қыл.'],
  ['DECEMBER', 'II', '14', 'Команда жұмысы', 'Команда ұраны ойлап тап: “Dig Deep. Think Smart. Build the Future!”'],
  ['DECEMBER', 'II', '14', 'Captain', 'Футболка / худи дизайнын жаса (логотип, ұран, түстер).'],
  ['DECEMBER', 'II', '14', 'Manager', '3D модель немесе слайд-анимация жасап, презентацияға қос.'],
  ['DECEMBER', 'II', '14', 'Builder', 'Миссияларды топтастыру (бір сапарда 2–3 миссия орындау).'],
  ['DECEMBER', 'II', '14', 'Programmer', 'FLL SPIKE Prime Gyro Alignment'],
  ['DECEMBER', 'II', '14', 'SMM', '📸 Инстаграм парақшасын белсенді жүргіз: Story, reels, robot progress, team spirit.'],
  ['DECEMBER', 'III', '15', 'Manager', 'Нағыз экспертпен (археолог, инженер, эколог) сұхбаттас — жюриге ұнайды.'],
  ['DECEMBER', 'III', '15', 'Builder', 'Күн сайын 15–20 минут “code + test + fix”'],
  ['DECEMBER', 'III', '15', 'SMM', 'Қысқа видео-интервью түсір (әр қатысушы өз рөлін таныстырады).'],
];

function normName_(s) {
  return String(s || '').toLowerCase().replace(/\u0451/g, '\u0435').replace(/\s+/g, ' ').trim();
}

function importClubData() {
  var existing = {}, byId = {};
  rows_('people').forEach(function (p) {
    if (String(p.role) === 'student') existing[normName_(p.name)] = p;
    byId[String(p.id)] = p;
  });

  var created = [], updated = 0;
  IMPORT_STUDENTS.forEach(function (row) {
    var name = row[0], extra = { cls: row[1], team: row[2], category: row[3], active: row[4] ? 1 : 0 };
    if (!row[4]) extra.note = 'Из прошлогоднего состава';
    var hit = existing[normName_(name)];
    if (hit) {
      update_('people', hit._row, {
        cls: extra.cls, team: extra.team, category: extra.category,
        active: extra.active, note: hit.note || extra.note || ''
      });
      updated++;
      return;
    }
    var pin = genPin_(4);
    var p = createPerson_('student', name, pin, IMPORT_SINCE, extra);
    existing[normName_(name)] = p;
    created.push({ name: name, pin: pin, cls: extra.cls, team: extra.team });
  });

  /* --- олимпиады и награды --- */
  var have = {};
  rows_('awards').forEach(function (a) { have[String(a.studentId) + '|' + String(a.title)] = true; });
  var addedAwards = 0, skipped = [];
  IMPORT_AWARDS.forEach(function (row) {
    var p = existing[normName_(row[0])];
    if (!p) { skipped.push(row[0]); return; }
    var key = String(p.id) + '|' + row[1];
    if (have[key]) return;
    have[key] = true;
    append_('awards', {
      id: newId_('a'), studentId: String(p.id), date: '', title: row[1],
      subject: '\u0420\u043e\u0431\u043e\u0442\u043e\u0442\u0435\u0445\u043d\u0438\u043a\u0430',
      level: row[2], result: row[3], note: row[4],
      updatedAt: nowIso_(), updatedBy: '\u0418\u043c\u043f\u043e\u0440\u0442'
    });
    addedAwards++;
  });

  /* --- календарь соревнований и годовой план --- */
  writeSetting_('competitions', IMPORT_COMPETITIONS.map(function (c) {
    return { title: c[0], when: c[1], term: c[2], status: c[3] };
  }));
  writeSetting_('plan', IMPORT_PLAN.map(function (x) {
    return { month: x[0], week: x[1], n: x[2], role: x[3], task: x[4] };
  }));

  /* --- лист с кодами --- */
  if (IMPORT_WRITE_PIN_SHEET && created.length) {
    var bk = book_();
    var sh = bk.getSheetByName('\u041a\u043e\u0434\u044b') || bk.insertSheet('\u041a\u043e\u0434\u044b');
    sh.clear();
    sh.appendRow(['\u0423\u0447\u0435\u043d\u0438\u043a', '\u041a\u043b\u0430\u0441\u0441', '\u041a\u043e\u043c\u0430\u043d\u0434\u0430', '\u041a\u043e\u0434']);
    created.forEach(function (c) { sh.appendRow([c.name, c.cls, c.team, c.pin]); });
    sh.getRange('D:D').setNumberFormat('@');
  }

  var msg = '\u0418\u043c\u043f\u043e\u0440\u0442 \u0437\u0430\u0432\u0435\u0440\u0448\u0451\u043d.\n\n' +
    '\u0414\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u043e \u0443\u0447\u0435\u043d\u0438\u043a\u043e\u0432: ' + created.length + '\n' +
    '\u041e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u043e: ' + updated + '\n' +
    '\u0417\u0430\u043f\u0438\u0441\u0435\u0439 \u043e \u043d\u0430\u0433\u0440\u0430\u0434\u0430\u0445: ' + addedAwards + '\n' +
    '\u0421\u043e\u0440\u0435\u0432\u043d\u043e\u0432\u0430\u043d\u0438\u0439 \u0432 \u043a\u0430\u043b\u0435\u043d\u0434\u0430\u0440\u0435: ' + IMPORT_COMPETITIONS.length + '\n' +
    '\u041f\u0443\u043d\u043a\u0442\u043e\u0432 \u0433\u043e\u0434\u043e\u0432\u043e\u0433\u043e \u043f\u043b\u0430\u043d\u0430: ' + IMPORT_PLAN.length +
    (created.length ? '\n\n\u041a\u043e\u0434\u044b \u0443\u0447\u0435\u043d\u0438\u043a\u043e\u0432 \u2014 \u043d\u0430 \u043b\u0438\u0441\u0442\u0435 \u00ab\u041a\u043e\u0434\u044b\u00bb. \u0420\u0430\u0437\u0434\u0430\u0439\u0442\u0435 \u0438\u0445 \u0438 \u0443\u0434\u0430\u043b\u0438\u0442\u0435 \u043b\u0438\u0441\u0442.' : '');
  return tell_(msg);
}
