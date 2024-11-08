# mini-excel-python
# Описание проекта: Учет сотрудников и товаров
Этот проект — это приложение на Python с использованием библиотеки Tkinter, которое предназначено для управления данными сотрудников и товаров компании. Программа сохраняет данные в файле Excel (используя библиотеку openpyxl), предоставляя возможности для добавления, редактирования, удаления, поиска записей и выполнения некоторых статистических расчетов. Проект может быть полезен для компаний, которым требуется простое и удобное решение для ведения базовой базы данных персонала и товарных запасов.

## Основной функционал
Приложение разделено на две основные категории: сотрудники и товары. Для каждой из них предусмотрены следующие функции:

1. Управление сотрудниками
   
Добавление сотрудника: Запись данных о новом сотруднике, включая ФИО, табельный номер, дату рождения, должность, дату приема на работу и зарплату. 
Удаление сотрудника: Удаление записи о сотруднике по его табельному номеру.
Редактирование сотрудника: Внесение изменений в существующие записи сотрудника.
Статистика по зарплате сотрудников: Расчет минимальной, максимальной и медианной зарплат сотрудников.

2. Управление товарами
   
Добавление товара: Запись информации о новом товаре, включая наименование, количество, цену. Пользователь также может отметить, закуплен ли товар.
Удаление товара: Удаление записи о товаре по его наименованию.
Редактирование товара: Внесение изменений в существующие записи о товаре.

## Технические особенности
Проект написан на языке Python с использованием следующих библиотек:

Tkinter: для создания графического интерфейса.
openpyxl: для работы с Excel-файлами, в которых хранятся данные о сотрудниках и товарах.
statistics: для расчета медианной зарплаты сотрудников.
Приложение сохраняет данные в формате Excel, что делает их легкодоступными для редактирования и анализа в популярных офисных программах. Интерфейс прост и интуитивно понятен, что позволяет легко использовать программу даже пользователям без опыта работы с базами данных.

## Пример использования
1.Открытие приложения и добавление данных о новом сотруднике или товаре.

2.Выполнение операций редактирования и удаления, а также быстрый поиск нужных записей.

# Заключение
Данный проект представляет собой удобное и легкое в использовании приложение для управления данными сотрудников и товаров компании. Возможности поиска и статистического анализа делают его полезным инструментом для компаний, которым необходима простая, но функциональная система учета данных.
