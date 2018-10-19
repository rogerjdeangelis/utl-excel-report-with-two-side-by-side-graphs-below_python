Excel report with two side by side graphs below using python

This techniques adds graphics to an existing excel workbooks

I don't believe this is NOT possible with SAS 9.4M2 'ODS Excel' , and
I had issuses with R XLConnect. Python opempyxl seemed to work.

github
https://tinyurl.com/yd7drugh
https://github.com/rogerjdeangelis/utl-excel-report-with-two-side-by-side-graphs-below_python

SAS  Forum
https://tinyurl.com/ya7ot45c
https://communities.sas.com/t5/Graphics-Programming/Ods-Excel-proc-report-and-2-proc-gchart/m-p/505890

Repository macros
https://tinyurl.com/y9nfugth
https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories

INPUT
=====

 WORK.REG total obs=10

  NAME_M     SEX_M    AGE_M    NAME_F     SEX_F    AGE_F

  Alfred       M        14     Alice        F        13
  Henry        M        14     Barbara      F        13
  James        M        12     Carol        F        14
  Jeffrey      M        13     Jane         F        12
  John         M        12     Janet        F        15
  Philip       M        16     Joyce        F        11
  Robert       M        12     Judy         F        14
  Ronald       M        15     Louise       F        12
  Thomas       M        11     Mary         F        15
  William      M        15                            .


EXAMPLE OUTPUT (just an example)
----------------------------------

   EXCEL : D:/XLS/SBYSOUT.XLSX

   WANT (NEW EXCEL SHEET WITH 'proc report'(not png) as histogram

   EXCEL   A           B            C
   ROW  ---------|-----------|---------------

   1    Country     Product     Actual Sales

   2    CANADA      BED           $47,729.00
   3                CHAIR         $50,239.00
   4                DESK          $52,187.00
   5                SOFA          $50,135.00
   6                TABLE         $46,700.00
   7    GERMANY     BED           $46,134.00
   8               CHAIR          $47,105.00
   9                DESK          $48,502.00


  Frequency               Frequency
  6 +                     6 +
    |                       |
  4 +                     4 +  *****   *****
    |                       |  *****   *****
  2 +  *****   *****      2 +  *****   *****
    |  *****   *****        |  *****   *****
    -----------------       -----------------
       BED     CHAIR            DESK   TABLE

 +------+
 |SHEET1|
 +------+


PROCESS
=======

* for development delete old copies;
%utlfkil(d:\xls\sbys.xlsx);
%utlfkil(d:\png\men.png);
%utlfkil(d:\png\women.png);

ods excel file="d:/xls/rpt.xlsx";
ods excel options(sheet_interval="none" sheet_name="sheet1" start_at="B1");
proc report data =reg;
  col('Men' name_m age_m sex_m)
    ('Women' name_f age_f sex_f );
run;quit;

ods excel close;

filename outfile "d:/png/male.png";

 goptions
    reset=goptions
    rotate=portrait
    gsfmode = replace
    device  = png
    gsfname = outfile
    vsize=2.5in
    hsize=2.5in
    htext=3
    display;   /* turn off the display of individual plots */
  run;quit;

proc gchart data=reg;
   vbar age_m / name="men";
run; quit;

filename outfile clear;

filename outfile "d:/png/female.png";

proc gchart data=reg ;
   vbar age_f /name="women";
run; quit;

filename outfile clear;

%utl_submit_py64("
import openpyxl;
from openpyxl import load_workbook;
wb = load_workbook(filename = 'd:/xls/rpt.xlsx');
ws = wb['sheet1'];
img = openpyxl.drawing.image.Image('d:/png/male.png');
ws.add_image(img,'A14');
img = openpyxl.drawing.image.Image('d:/png/female.png');
ws.add_image(img,'D14');
wb.save('d:/xls/rpt.xlsx');
");


*                _               _       _
 _ __ ___   __ _| | _____     __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \   / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/  | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|   \__,_|\__,_|\__\__,_|

;


data men(rename=(name=name_m age=age_m sex=sex_m   )
  keep=name age sex i);

do i=1 to 10;
set sashelp.class(where=(sex='M'));
i=i;
output;
end;

run;

data women(rename=(name=name_f age=age_f sex=sex_f )
  keep=name age sex i);
do i=1 to 10;
set sashelp.class(where=(sex='F') );
i=i;
output;
end;
run;

data reg;
merge  men women;by i;
run;




