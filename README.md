# Содержание

1. [Аннотация](#аннотация)
2. [Установка](#установка)
3. [Окружение](#окружение)
4. [Работа с шаблоном <b>xlsx</b>](#шаблон_xlsx)  
4.1. [Вид документа](#вид_документа)  
4.2. [Данные о покупателе](#данные_о_покупателе)  
4.3. [Налог](#налог)  
4.4. [Валюта](#валюта)  
4.5. [Мера имерения](#мера_измерения)  
4.6. [Данные о товарах](#данные_о_товарах)  
4.7. [Данные о продавце](#данные_о_продавце)  
4.8. [Системные данные](#системные_данные)  
4.9. [Фундаментальные правила работы с шаблоном](#шаблон_правила)  
5. [Запуск и работа с <b>FacturaPy</b>](#запуск)  
5.1. [Основной сценарий](#основной_сценарий)  
5.2. [Фундаментальные правила работы с <b>FacturaPy</b>](#правила_запуск)  
5.3. [Схема](#схема)  
5.4. [Исключения](#исключения)  
6. [Разработчик](#разработчик)

<br/>

## <a name='аннотация'>1. Аннотация</a>

Внимательно прочтите изложенные здесь инструкции и наслаждайтесь работой с <b>FacturaPy</b>.  
<b>FacturaPy</b> предназначена для автоматизации процесса формирования транспортных документов, необходимых для предоставления отчетности в Республике Польше, посредством web-сайта https://www.fakturowo.pl/.  
<b>FacturaPy</b> является полностью независимой и исполняемыой программой формата .exe. Для ее корректной работы не нужны какие-либо дополнительные файлы, утилиты и т.д. Достаточно просто скачать программу и начать ее использовать.

<br/>

## <a name='установка'>2. Установка</a>

Для установки программы выполните следующие шаги:  
1. Перейдите по предоставленной разработчиком персональной ссылке. При переходе Вы попадаете в приватный репозиторий в <a href="https://github.com/" target="_blank">GitHub</a>, доступ к которому имеется только у разработчиков и пользователей с прямой ссылкой.

2. Кликните на файл <b>FacturaPy</b>  
![image](https://user-images.githubusercontent.com/116018998/232195611-789b8ff8-26ca-4903-9be6-a79f3ba91edc.png)

3. Кликните на кнопку загрузки <b>Download</b>  
![image](https://user-images.githubusercontent.com/116018998/232195725-a0a4da17-1776-43ef-9037-ddcc70d73a86.png)

4. Создайте папку в любом удобном для вас месте на ПК.
5. Перейдите в папку загрузок, в которую скачался файл <b>FacturaPy</b> на шаге 3, и переместите данный файл в созданную Вами папку на шаге 4.
6. Создайте в данной папке еще одну папку с именем <b>xlsx</b>.
7. Вернитесь к шагу 1.
8. Перейдите в директорий <b>xlsx</b>  
![image](https://user-images.githubusercontent.com/116018998/232198657-b9d8b1f9-339d-4d6b-bc69-257baf5e1c72.png)

9. Перейдите в файл <b>book_1.xlsx</b>  
![image](https://user-images.githubusercontent.com/116018998/232198893-40d5b37a-9971-4da3-8740-95fe9f630123.png)

10. Кликните на кнопку загрузки <b>Download</b>  
![image](https://user-images.githubusercontent.com/116018998/232199126-c9a8eac1-473d-4394-81c2-030bed8f1b5d.png)

11. Перейдите в папку загрузок, в которую скачался шаблон <b>book_1.xlsx</b>, и переместите данный файл в созданную Вами папку <b>xlsx</b> на шаге 6.

В результате у вас должна быть следующая структура:  
![image](https://user-images.githubusercontent.com/116018998/232199226-4b088841-a4c9-4ac7-a849-ebe0cce83f0a.png)
![image](https://user-images.githubusercontent.com/116018998/232199274-a7cdd5bf-0e32-46f5-ac31-e40961256301.png)

<br/>

## <a name='окружение'>3. Окружение</a>

Для корректной работы <b>FacturaPy</b> необходим установленный на устройстве браузер <b>Chrome</b> версии <b>111.0.5563.</b>[последние три цифры не имеют значения]  
![image](https://user-images.githubusercontent.com/116018998/229341814-799df1fa-5e66-4c04-a1f8-332f5d5e48bd.png)

Для того, чтобы узнать версию браузера Chrome, необходимо выполнить следующие шаги:
1. Нажать на кнопку "меню" в правом верхнем углу  
![image](https://user-images.githubusercontent.com/116018998/229342222-7f978dea-e86c-4bbf-a52c-35f2717e2a69.png)

2. Нажать на "Справка"  
![image](https://user-images.githubusercontent.com/116018998/229342406-9d30971e-1834-44dc-88ef-8f5dcdd502e9.png)

3. В раскрывшемся меню нажать на "О браузере Google Chrome"  
![image](https://user-images.githubusercontent.com/116018998/229342477-b311103c-7d32-4b44-9979-a556043869f0.png)

В результате в окне должна отображаться версия Вашего браузера Chrome  
![image](https://user-images.githubusercontent.com/116018998/229341814-799df1fa-5e66-4c04-a1f8-332f5d5e48bd.png)

Если версия отличатеся от трубуемой, перейдите по этой <a href="https://google-chrome.ru.uptodown.com/windows/download/95500106" target="_blank">ссылке</a> и установите нужную версию.

<br/>

## <a name='шаблон_xlsx'>4. Работа с шаблоном xlsx</a>

Используйте шаблон <b>book_1.xlsx</b> для работы с <b>FacturaPy</b>. <b>FacturaPy</b> получает шаблон в качестве входного файла (датафрейма), из которого извлекает данные для формирования транспортных документов. В <b>FacturaPy</b> переданы строгие адреса ячеек для каждого поля на сайте https://www.fakturowo.pl/, ввиду чего разметка шаблона не должна меняться. Заголовки таблиц и данные для них должны всегда оставаться в строго определенных местах.

### <a name='вид_документа'>4.1. Вид документа</a>

Ячейка для вида документа должна находиться в ячейке "В1". В данную ячейку переданы имеющиеся на сйте https://www.fakturowo.pl/ виды докуентов, содержащиеся в поле:  
![image](https://user-images.githubusercontent.com/116018998/232202469-ee10fd70-5631-4d8d-942d-51cd8ee605b4.png)

Ячейка имеет формат раскрывающегося списка. Для того, чтобы сформировать нужный документ просто раскройте список и выберете его:  
![image](https://user-images.githubusercontent.com/116018998/232202416-9c738532-9fad-4c5e-bcdd-cefb9eca3b47.png)

### <a name='данные_о_покупателе'>4.2. Данные о покупателе</a>

Данные о покупателе должны находиться в диапозоне "В2:В6":  
![image](https://user-images.githubusercontent.com/116018998/232202586-3dccae25-6d18-4cd6-ac74-e61d4f477a48.png)

По техническому заданию заказчика в ячейки переданы определенные значения для двух компаний. С целью недопущения изменения формулы ячейки представлены также в формате раскрывающегося списка. Для того, чтобы выбрать компании, раскройте список с именем компании и выберите нужную:  
![image](https://user-images.githubusercontent.com/116018998/232202732-c4842777-f1e6-4d10-a780-8f8f1d58a3f7.png)

Данные в остальные поля подятнуться автоматически в виде единственного варианта выбора, при этом их также нужно выбрать:  
![image](https://user-images.githubusercontent.com/116018998/232202849-55911e42-4bc9-44ff-8702-e2c97a85f9ba.png)

Кроме этого, в данные ячейки также можно ввести иные данные без потери данных об определенных двух компаниях:  
![image](https://user-images.githubusercontent.com/116018998/232203183-ee8ef2fb-4bd0-45af-a3ec-1dd11ea3b2fa.png)

Переданные в ячейки данные будут заполнены в следующие поля на сайте:  
![image](https://user-images.githubusercontent.com/116018998/232203650-ff936bf5-431c-4bcf-b39e-283ee3223065.png)

### <a name='налог'>4.3. Налог</a>

Данные о размере налога должны находится в ячейке "В7":  
![image](https://user-images.githubusercontent.com/116018998/232203980-5bf36ed7-0a77-4d18-83ef-1a5d574f7d8c.png)

Данная ячейка представлена в формате раскрывающегося списка и имеет предопределенные значения. Для выбора размера налога раскройте список и выберите его:  
![image](https://user-images.githubusercontent.com/116018998/232204395-91dbde8c-c3ae-479a-965b-e2900cd845db.png)

В данную ячейку можно поместить свое значение, при этом такое значение должно полностью соответствовать значению, представленному для соответствующего поля на сайте:  
![image](https://user-images.githubusercontent.com/116018998/232204795-d7659a8f-2c5f-43f4-9863-a0b81f8ff9a0.png)
Но вводить свои значения не рекомендуется по причине различного форматирования значений на сайте и в шаблоне, в результате чего нужный налог может быть не выбран, а <b>FacturaPy</b> может привести к сбою.

### <a name='валюта'>4.4. Валюта</a>

Данные о валюте должны находиться в ячейке "В8":  
![image](https://user-images.githubusercontent.com/116018998/232205288-77b407c9-674c-4aee-a295-4571fdd2817e.png)

Правила заполнения и обработки аналогичны ячейке для [налога](#налог).

### <a name='мера_измерения'>4.5. Мера измерения</a>

Данные о валюте должны находиться в ячейке "В9":  
![image](https://user-images.githubusercontent.com/116018998/232205406-f81cb148-43c2-4333-aa4a-51a01f4bdcaf.png)

Правила заполнения и обработки аналогичны ячейке для [налога](#налог).

### <a name='данные_о_товарах'>4.6. Данные о товарах</a>

Данные о наименовании товара, его колиестве и цене должныи следовать в столбцах "А", "В" и "С" со строки 13:  
![image](https://user-images.githubusercontent.com/116018998/232201114-336690e8-0d6a-434d-aef9-e359c276c6f1.png)

Данные в указанные ячейки вводятся самостоятельно пользователем. Вы также можете выполнить команды копирования и вставки, при этом заголовки таблицы должны соответствовать полям вставленных элементов.  
На excel листе может быть сколько угодно позиций. Все позиции будут переданы в соответствующие поля на сайте:  
![image](https://user-images.githubusercontent.com/116018998/232205799-9842e900-4732-4969-ae8a-8df08790890e.png)

### <a name='данные_о_продавце'>4.7. Данные о продавце</a>

Техническое задание и пользование сайтом https://www.fakturowo.pl/ предполагает, что в полях о продавце данные сохранены и заполнены:  
![image](https://user-images.githubusercontent.com/116018998/232206110-0db86a9e-d800-4e48-993f-991a8cb2a76c.png)

<b>FacturaPy</b> каждый раз при запуске осуществляет авторизацию пользователя, ввиду чего данные о продавце должны быть сохранены пользователем в личном кабинете на сайте. Поэтому данные о продавце в шаблон не передаются, а соответствующие ячейки для заполнения в шаблоне отсутствуют.  
Если Вы хотите ввести данный функционал, обратитесь к [разработчику](#разработчик).

### <a name='системные_данные'>4.8. Системные данные</a>

Для того, чтобы не нагружать <b>FacturaPy</b> и обеспечить быстроту ее действия, все данные для полей в шаблоне помещены в скрытый запароленный лист excel. Строго не рекомендуется раскрывать данный лист в шаблоне и самостоятельно вносить в него какие-либо изменения. При желании расширить / изменить шаблон обратитесь к [разработчику](#разработчик).

### <a name='шаблон_правила'>4.9. Фундаментальные правила работы с шаблоном</a>

- не изменять структуру шаблона
- не смещать / не изменять адреса ячеек
- при наличии в ячейке предопределенных значений (раскрывающийся список) рекомендуется выбирать значение из списка
- не оставлять пустые значения в ячейках
- один шаблон может содержать сколько угодно листов, Вы можете добавлять дополнительные листы / удалять лишние листы по Вашему усмотрению
- один лист - один транспортный документ
- шаблон должен находиться в папке <b>xlsx</b>, которая, в свою очередь, должна находиться в корне программы ([инструкция по установке](#установка))  
- в папке <b>xlsx</b> может находиться сколько угодно шаблонов
- если вы внесли изменения в шаблон, и он стал некорректно обрабатываться, просто скачайте шаблон с изначальными настройками, выполнив шаги 7-11 [инструкции по установке](#установка) или обратитесь к [разработчику](#разработчик).

<br/>

## <a name='запуск'>5. Запуск и работа с <b>FacturaPy</b></a>

### <a name='основной_сценарий'>5.1. Основной сценарий</a>

После [установки](#установка) <b>FacturaPy</b>, [проверки версии Chrome](#окружение) и внесения необходимых данных для транспортных документов в шаблон:
1. Поместите входной excel файл в папку <b>xlsx</b>
2. Запустите <b>FacturaPy</b>
3. При первом запуске введите логин и пароль от https://www.fakturowo.pl/
4. Сохраните (или нет) логин и пароль от https://www.fakturowo.pl/  
![image](https://user-images.githubusercontent.com/116018998/232208834-f03a7b9f-b930-4de1-b08d-7785e22373ae.png)

5. Следуйте инструкциями, отображаемым в консоли

После запуска <b>FacturaPy</b> программа проверит входные excel файлы и все лиисты, находящиеся в них.  
В случае успешной проверки в консоли отобразится следующая информация:  
![image](https://user-images.githubusercontent.com/116018998/232208894-4032e5ec-f334-4828-9a56-156a8dfb740c.png)


После этого <b>FacturaPy</b> самостоятельно откроет Chrome со стартовыми настройками и начнет свое выполнения, имитируя действия человека по заполнению полей на сайте. 
![image](https://user-images.githubusercontent.com/116018998/232209721-6ae8d5f6-5bf4-45d9-ac21-5503b674e7da.png)

Когда <b>FacturaPy</b> занесет последнюю позицию товара, переданную в шаблон, ее выполнение приостановится, а пользователю предоставится возможность проверить и убедиться в корректности введенных в поля на сайте данных, а также при необходимости внести изменения.  
После этого введти в консоль <b>"next"</b>, и <b>FacturaPy</b> самостоятельно скачает сформированный транспортный документ, и продолжит свое выполнение:  
![image](https://user-images.githubusercontent.com/116018998/232209952-3c0d7777-0917-4128-9594-1f0282bf4f46.png)

Когда <b>FacturaPy</b> сформирует последний транспортный документ, то есть дойдет до конца последнего листа последнего файла, консоль уведомит об успешном выполнении программы, а пользователь в папке загрузок получит сформированные транспортные документы.

### <a name='правила_запуск'>5.2. Фундаментальные правила работы с <b>FacturaPy</b></a>

- не рекомендуется вводить данные в полях / нажимать на кнопки на сайте / выполнять иные действия, которые будут иметь результат, пока <b>FacturaPy</b> Вас сама об этом не попросит
- не рекуомендуется оставлять пустые поля на сайте
- всегда следуйте инстракциям, отображаемым в консоли
- если в Chrome долгое время (более 5 секунд) не происходит никаких изменений, вернитесь в консоль и выполните отображаемую инструкцию

### <a name='схема'>5.3. Схема</a>

![Untitled Workspace](https://user-images.githubusercontent.com/116018998/229340061-820b693f-3646-4113-b7cb-ddb0ffe9635c.png)

### <a name='исключения'>5.4. Исключения</a>

- <b>Некорректный входной excel файл.</b>  
После запуска <b>FacturaPy</b> программа проверит входные excel файлы и все лиисты, находящиеся в них. 
В случае наличия некорректных либо пустых полей программа приостановится и попросит пользователя самостоятельно устранить недостатки, а также отобразит адреса ячеек с пустыми или некорректными полями:  
![image](https://user-images.githubusercontent.com/116018998/232209058-65a74e14-12ae-4700-bab1-5f555ebc7852.png)  
Это значит, что одна или несколько ячеек в диапазоне "А13:С13" - пусты.  
После устранения недостатков, сохраните документ, введите "next" и нажмите Enter для возобновления <b>FacturaPy</b>:  
![image](https://user-images.githubusercontent.com/116018998/232209153-e890840c-59f5-4a20-b8fa-c92c1b00c564.png)

- <b>Неверный логин или пароль.</b>  
Если введеный логин или пароль неверны, сайт отобразит сообщение об ошибке:  
![image](https://user-images.githubusercontent.com/116018998/232210451-a4ca09d3-853a-4f87-811c-0a5f2377bfa1.png)  
А <b>FacturaPy</b> попросит повторно ввести логин и пароль:  
![image](https://user-images.githubusercontent.com/116018998/232210502-54041645-a371-447b-af19-df7da58073b4.png)  
После ввода верного логина и пароля <b>FacturaPy</b> продолжит свое выполнение.

- <b>Пустые ообязательные поля на сайте.</b>  
Если введенные на сайте обязательные поля остались пусты или некорректны:  
![image](https://user-images.githubusercontent.com/116018998/232210716-8615a006-ef73-4635-8fe4-a08aa33fa6f5.png)  
На сайте отобразится сообщение об ошибке:  
![image](https://user-images.githubusercontent.com/116018998/232210741-f3c35c09-24ec-4b43-bd79-2384e8259cbe.png)  
А <b>FacturaPy</b> попросит повторно проверить корректность введенных данных:  
![image](https://user-images.githubusercontent.com/116018998/232210772-343f0cc9-0847-48ad-9376-5255642a230e.png)  
После корректного ввода программа продолжит свое выполнение.

- <b>Иное.</b>
При наличии ошибок (исключений), неописанных в настоящем туториале, и при невозможном самостоятельном их устранении, а также при наличии таких фатальных ошибок, как вылет программы, обратитесь к [разработчику](#разработчик).

<br/>

## <a name='разработчик'>6. Разработчик</a>

- Телеграм: @alr_ks (прямая ссылка: https://t.me/alr_ks)
- Почта: alexandr.kosyrew@mail.ru

