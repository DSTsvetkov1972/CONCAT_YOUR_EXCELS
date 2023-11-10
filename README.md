# CONCAT_YOUR_EXCELS
## Программа позволяет быстро объеденять в одну таблицу большое количество экселевских таблиц раскиданных в разных экселевских книгах на разных листах.
## Заголовки таблиц могут располагаться в разных строках (где-то, например, на третьей строке, а где-то на пятой) и начинаться с разных колонок. Названия колонок могут иметь разный порядок. Таблицы могут иметь как одинаковые названия колонок так и разные. Пользователь может выбрать листы в книгах Эксель таблицы с которых должны быть объединены.
## Отбор листов экселевских книг с таблицами для объединения и поиск заголовков осуществляется средствами Power Query, что удобнее, чем обработка на Python.<br>Само объединение осуществляется с помощью скрипта на Python, что значительно быстрее, чем аналогичная обработка в Power Query.
## Результат выгружается в CSV-файл, так как итоговая таблица может иметь количество строк превышающее допустимое в Эксель.
## Из итогового результата автоматически вычищаются пустые строки и колонки.
---
### Как пользоваться:

1. Создайте парку проекта
2. В папке проекта создайте папку **Исходники**. В папку **Исходники** поместите экселевские файлы. Экселевские файлы могут расоплогаться как непосредственно в самой папке **Исходники**, так и во вложенных папках (допускается любой уровень вложенности).<br>**ВНИМАНИЕ: ячейка A1 каждого листа, таблицу с которого нужно будет добавить в итоговую, должна быть не пустой или иметь форматирование отличное от форматирования по умолчанию или хотя-бы однажды быть измененной.** 
*Это связано с особенностью импорта данных с листа эксель Power Query. Подробнее можно прочитать по ссылке: [Импорт данных в Power Query и Power BI из листа Excel: ловушка UsedRange](http://excel-inside.pro/blog/2017/05/23/%D0%B8%D0%BC%D0%BF%D0%BE%D1%80%D1%82-%D0%B4%D0%B0%D0%BD%D0%BD%D1%8B%D1%85-%D0%B2-power-query-%D0%B8-power-bi-%D0%B8%D0%B7-excel-%D0%BB%D0%BE%D0%B2%D1%83%D1%88%D0%BA%D0%B0-usedrange/?ysclid=losc79am53362322864)*
3. Скопируйте в папку проекта файлы 
    - **_CONCAT_YOUR_EXCELS (предобработка).xlsm**
    - **_CONCAT_YOUR_EXCELS.exe**
4. Откройте **_CONCAT_YOUR_EXCELS (предобработка).xlsm** на листе **Settings**
5. В ячейку A2 введите путь к папке проекта без "\" на конце. Ячейка A4 содержит количество строк в которых будут искаться заголовки в исходных таблицах. Слишком большое значение в этой ячейке может быть причиной замедления предобработки для больших таблиц.
6. Перейдите на лист **Sheets_in_book** и нажмите кнопку **Получить список листов в книгах**. В получившейся таблице перечислены все книги эксель в папке **Исходники**. В колонке **Отбираем для слияния** проставьте **ДА** для листов, таблицы с которых должны быть добавлены в итоговую таблицу. Для листов которые не нужно обрабатывать ячейка в этой колонки должна быть пустой.
7. Перейдите на лист **Columns in book** и нажмите кнопку **Посмотерь количество столбцов**. Если количество столбцов на каких-либо листах покажется вам подозрительным, откройте соответствующие книги и при необходимости откорректируйте находящуюся там информацию.
8. Откройте какую-либо книгу из папки **Исходники**. Определите признак заголовка. Например для книги **Test1.xlsx** для листа **Ноябрь** признак заголовка "№" в колонке 2. Обратите внимание: ячейка A1 изначально была пустой и былы отформатирована в желтый цвет для корректной обработки UserRange. Перейдите на лист **Settings** и в внесите в таблицу номер колонки "2" и значение "пп №". Перейдите на лист **Titles_in_book** и нажмите кнопку **Подтянуть заголовки**. В таблице отобразятся номер строки и заголовки для листов в которых удастя найти заголовки по указонному признаку. Откройте любую книгу для которой заголовки найти не удалось и посмотрите по какому признаку заголовки можно определить для этого листа. Например для **Test2.xlsx** для листа **Октябрь 2023** 3 колонка значение "Пункт". Добавьте в таблицу на листе **Settings** книги **_CONCAT_YOUR_EXCELS (предобработка).xlsm** строку "4","Пункт". Перейдите на лист **Titles_in_book** и нажмите кнопку **Подтянуть заголовки**. Повторите процедуру для листов для которых не удалось найти заголовки, пока вся таблица не окажется заполненной. Если по непонятным причинам остаются незаполненные строки, проверьте чтобы на соответствующих листах была заполнена ячейка A1 и значение ячейки A4 книги **_CONCAT_YOUR_EXCELS (предобработка).xlsm** на листе **Settings** было достаточным, чтобы строки с заголовками были просмотрены в процессе предобработки.
9. Сохраните книгу **_CONCAT_YOUR_EXCELS (предобработка).xlsm**
10. Если файл **_CONCAT_YOUR_EXCELS.csv** существует в папке проекта, убедитесь что он закрыт и запустите программу **_CONCAT_YOUR_EXCELS.py**
11. В случае если некоторые таблицы невозможно обработать (например по причине того что для них не найдены заголовки), выполнение программы будет прерываться всплывающим окном с предупреждением. Если нажать "ОК", выполнение продолжится, но таблица не будет добавлена в результат. Выполнение можно прервать полностью если закрыть окно программы.<br>
Результат выполнения программы будет сохранен в файле **_CONCAT_YOUR_EXCELS.csv** 
