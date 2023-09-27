Библиотека функций для работы с гугл таблицами на основе гугловской библиотеки "sheets". 
Работает быстрее на больших объемах данных, в сравнении с ванильными функциями SpreadsheetApp.getActiveSpreadsheet().getSheetByName().getRange().getValues(), но имеет суточным лимит на вызовы.

В этой версии библиотеки количество вызовов sheets api уменьшено в двое, добавлены новые функции. Обращение к методам реализовано через цепочку прототипов.

Из-за ошибки гугла от 6 июня 2012 г. библиотека не будет показывать всплывающие подсказки, свои методы и автозаполнение, но будет корректно работать, если знать, что вызывать.
Рекомендую разместить код библиотеки непосредственно в код вашего скрипта - тогда подсказки будут работать.

# Методы библиотеки:

```
GetTable() - возвращает файл гугл таблицы. Принимает ID файла. Если ID не указан - возвращает контейнер скрипта
```

## Методы GetTable():

```
GetSheet(sheet_id) - возвращает лист гугл таблицы. Принимает ID листа или его имя.

CreateSheet(sheet_id, sheet_name, sheet_index = 0) - Создает лист гугл таблицы. Принимает ID нового листа, имя нового листа, индекс нового листа. Возвращает созданный лист

DeleteSheet(sheet_id) - Удаляет лист гугл таблицы. Принимает ID листа. Возвращает true, если успешно.

ChangeEditorsInProtectedRanges(gmails, adding) - Добавляет или удаляет из всех защищенных диапазонов указанный список пользователей. Принимает массив адресов и {Boolean}  adding: true - добавляем. false - убираем.
```


## Методы GetSheet():

```
GetValues(firstRow = 1, firstCol = 1, rows = "", columns = 99) - Возвращает данные с листа

SetValues(output_arr, firstRow = 1, firstCol = 1) - Помещает двумерный массив на лист

Clear(firstRow = 1, firstCol = 1, rows = "", columns = 99) - Стирает данные с листа гугл таблиц.

ClearContent(firstRow = 1, firstCol = 1, rows = 0, columns = 99) - Стирает введенные пользователем данные с листа гугл таблиц, оставляя флажки, выпадающие списки, форматирование и тд

DeleteDuplicates(first_row = 1, first_col = 1, rows, columns) - Удаляет дубликаты строк в выбраном диапазоне.

DeleteRows(rows_arr) - Удаляет строки. Принимает одномерный массив номеров

ConvertSheet(name, target_folder, format = "pdf") - создает в папке новый файл в выбранном формате из указаного листа гугл таблицы. Возвращает созданный файл

ConvertSheetAndDownload(format = "pdf") - Скачать лист в выбранном формате. форматы pdf, xlsx, ods, zip, csv
```

## Методы GetValues():

`
GetValuesFromBasicFilter() - Возвращает данные, которые отображает на листе базовый фильтр
`

# Примеры использования:

```
const file = GetTable().file // объект текущего файла
const file = GetTable("1uE93B1mO7e4ZKYlcqUevmKCYTeI4t-5OS8ILXparctV8AsYr8W").file // объект указанного файла

const sheet = GetTable().GetSheet(0).sheet // объект листа
const sheet = GetTable().GetSheet("List1").sheet // объект листа

const values = GetTable().GetSheet(0).GetValues() // получить все данные с листа
const values = GetTable().GetSheet(0).GetValues(3,4,false, 5) // получить данные с листа, начиная со строки 3 и колонки 4 (D4) пять колонок до самого низа листа
const values = GetTable().GetSheet(0).GetValues().GetValuesFromBasicFilter() // получить данные с листа, которые отображает на листе базовый фильтр
```
