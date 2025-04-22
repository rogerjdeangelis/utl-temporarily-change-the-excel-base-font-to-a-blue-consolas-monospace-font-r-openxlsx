%let pgm=utl-temporarily-change-the-excel-base-font-to-a-blue-consolas-monospace-font-r-openxlsx;

%stop_submission;

Temporarily change the excel base font to a blue consolas monospace font r openxlsx

github
https://tinyurl.com/3uv9fvkx
https://github.com/rogerjdeangelis/utl-temporarily-change-the-excel-base-font-to-a-blue-consolas-monospace-font-r-openxlsx

related
https://tinyurl.com/ycr9r2da
https://communities.sas.com/t5/ODS-and-Base-Reporting/Modifying-a-style-for-Excel/m-p/963485#M26857


/**************************************************************************************************************************/
/*       INPUT                   |         PROCESS                         |             OUTPUT                           */
/*       =====                   |         =======                         |             =======                          */
/*                               |                                         |                                              */
/*  NAME      SEX    AGE         | %utlfkil(d:/xls/class.xlsx)             | ----------------                             */
/*                               |                                         | |Consolas | 14 |                             */
/* Alfred      M      14         | %utl_rbeginx;                           | -----------------------+                     */
/* Alice       F      13         | parmcards4;                             | | A1| fx    |DAYNUM    |                     */
/* Barbara     F      13         | library(openxlsx)                       | -----------------------------------+         */
/* Carol       F      14         | library(haven)                          | [_] |    A     |    B    |    C    |         */
/* Henry       M      14         | source("c:/oto/fn_tosas9x.R")           | -----------------------------------|         */
/*                               | have<-read_sas("d:/sd1/have.sas7bdat")  |  1  | NAME     |   SEX   |   AGE   |         */
/*                               | print(have)                             |  -- |----------+---------+---------|         */
/* options validvarname=upcase;  | wb <- createWorkbook()                  |  2  |Alfred    | M       | 14      |         */
/* libname sd1 "d:/sd1";         | addWorksheet(wb, "CLASS")               |  -- |----------+---------+---------|         */
/* data sd1.have;                | # Change the default font               |  3  |Alice     | F       | 13      |         */
/*  set sashelp.class            | modifyBaseFont(                         |  -- |----------+---------+---------|         */
/*   (obs=5 keep=name age sex);  |  wb                                     |  4  |Barbara   | F       | 13      |         */
/* run;quit;                     | ,fontSize = 14                          |  -- |----------+---------+---------|         */
/*                               | ,fontColour = "#0072B2"                 |  5  |Carol     | F       | 14      |         */
/*                               | ,fontName = "Consolas")                 |  -- |----------+---------+---------|         */
/*                               | writeData(wb, "CLASS", have)            |  6  |Henry     | M       | 14      |         */
/*                               | saveWorkbook(                           |  -- |----------+---------+---------|         */
/*                               |  wb                                     | [CLASS]                                      */
/*                               | ,"d:/xls/class.xlsx"                    |                                              */
/*                               | ,overwrite = TRUE)                      |                                              */
/*                               | ;;;;                                    |                                              */
/*                               | %utl_rendx;                             |                                              */
/*                               |                                         |                                              */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.have;
 set sashelp.class
  (obs=5 keep=name age sex);
run;quit;

/**************************************************************************************************************************/
/*    NAME      SEX    AGE                                                                                                */
/*                                                                                                                        */
/*   Alfred      M      14                                                                                                */
/*   Alice       F      13                                                                                                */
/*   Barbara     F      13                                                                                                */
/*   Carol       F      14                                                                                                */
/*   Henry       M      14                                                                                                */
/**************************************************************************************************************************/

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

%utlfkil(d:/xls/class.xlsx)

%utl_rbeginx;
parmcards4;
library(openxlsx)
library(haven)
source("c:/oto/fn_tosas9x.R")
have<-read_sas("d:/sd1/have.sas7bdat")
print(have)
wb <- createWorkbook()
addWorksheet(wb, "CLASS")
# Change the default font
modifyBaseFont(
 wb
,fontSize = 14
,fontColour = "#0072B2"
,fontName = "Consolas")
writeData(wb, "CLASS", have)
saveWorkbook(
 wb
,"d:/xls/class.xlsx"
,overwrite = TRUE)
;;;;
%utl_rendx;

/**************************************************************************************************************************/
/* ----------------                                                                                                       */
/* |Consolas | 14 |                                                                                                       */
/* -----------------------+                                                                                               */
/* | A1| fx    |DAYNUM    |                                                                                               */
/* -----------------------------------+                                                                                   */
/* [_] |    A     |    B    |    C    |                                                                                   */
/* -----------------------------------|                                                                                   */
/*  1  | NAME     |   SEX   |   AGE   |                                                                                   */
/*  -- |----------+---------+---------|                                                                                   */
/*  2  |Alfred    | M       | 14      |                                                                                   */
/*  -- |----------+---------+---------|                                                                                   */
/*  3  |Alice     | F       | 13      |                                                                                   */
/*  -- |----------+---------+---------|                                                                                   */
/*  4  |Barbara   | F       | 13      |                                                                                   */
/*  -- |----------+---------+---------|                                                                                   */
/*  5  |Carol     | F       | 14      |                                                                                   */
/*  -- |----------+---------+---------|                                                                                   */
/*  6  |Henry     | M       | 14      |                                                                                   */
/*  -- |----------+---------+---------|                                                                                   */
/* [CLASS]                                                                                                                */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
