# get_a_pollutant

<br>Пользовательская функция Microsoft Excel <i><b>"get_pollutant"</i></b> возвращает характеристику загрязняющего вещества по его коду. 
<br>Принимает 2 аргумента:
- <i>Code</i> – код загрязняющего вещества (primary key);
- <i>Parametr</i> (default = 1) - каждой возврщаемой характеристике вещества соответствует числовой код:
1 – name pollutant; 2 – PDKmr; 3 – PDKss; 4 – PDKsg; 5 – OBUV; 6 – class of a danger.; 7 – agregat; 8 – PDV?; 9 –  VOC?; 10 – ch. formula

Характеристики веществ хранятся в БД Microsoft SQL Server Management Studio 18. Информация передается с помощью хранимой процедуры  <i><b>"proc_GetSubstances"</i></b>
