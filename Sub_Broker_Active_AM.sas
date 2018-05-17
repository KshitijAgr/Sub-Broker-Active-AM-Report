
/*Importing Excel Files TO SAS*/

%let location= /folders/myfolders/sasuser.v94/data/Sub-Broker active AM Report/;

%macro Import(file_name,ds_name,db_name);
   proc import
   datafile="&location&file_name"
   out=&ds_name
   dbms=&db_name
   replace;
   run;
   proc sort data=&ds_name;
   by TM_CODE TM_NAME;
   run;
%mend;

%import(Rad1.CSV,Rad1,CSV);
%import(Rad2.CSV,Rad2,CSV);
%import(Rad3.CSV,Rad3,CSV);
%import(Rad4.CSV,Rad4,CSV);
%import(Rad5.CSV,Rad5,CSV);
%import(Rad6.xls,Rad6,xls);

/*Merging of All DataSets*/
data RAD;
merge RAD1 RAD2 RAD3 RAD4 RAD5 RAD6;
by tm_code tm_name; 
run;

/*Preparing DataSets*/
data RAD;
length TM_NAME $30.;
format TM_NAME $30.;
length SH_NAME $30.;
format SH_NAME $30.;
set rad;
run;

/*Creating Macros For dates*/
 
%global day yday month pre_month;

%macro dates;
%let day=%sysfunc(today(),date11.);
%let yday=%sysfunc(intnx(day,"&day"d,-1),date11.);
%let month=%sysfunc(today(),monname11.);
%let pre_month=%sysfunc(intnx(month,"&day"d,-1),monname11.);
/* %put &day &yday &month &pre_month*/
%mend;
%dates;


/*Title and Footnote*/

title height=12pt "Sub-Broker Active AM Report As on &day";
footnote" ";
footnote2 height=12pt "Note: This is a system generated mail. Please do not respond to this mail.";

 PROC REPORT DATA=WORK.RAD LS=132 PS=60  SPLIT="/" CENTER Missing;
 COLUMN  TM_CODE TM_NAME SH_NAME AM_CODE AM_NAME
 ("INWARDS" INWARDDAY INWARDMON INWARDLMON DIFF_INWARD)
 ("NCA" NCADAY  NCAMON NCALMON DIFF_NCA);
  
 DEFINE  TM_CODE / GROUP noprint FORMAT= $8. WIDTH=8     SPACING=2   LEFT "TM_CODE" ;
 DEFINE  TM_NAME / GROUP FORMAT= $30. WIDTH=30    SPACING=2   LEFT "TM Name" ;
 DEFINE  SH_NAME / GROUP FORMAT= $30. WIDTH=30    SPACING=2   LEFT "SH Name" ;
 DEFINE  AM_CODE / DISPLAY FORMAT= $8. WIDTH=8     SPACING=2   LEFT "AM/Code" ;
 DEFINE  AM_NAME / DISPLAY FORMAT= $23. WIDTH=23    SPACING=2   LEFT "AM Name" ;
 	
 DEFINE  INWARDDAY / ANALYSIS FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "&yday" ;
 DEFINE  INWARDMON / ANALYSIS FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "&month" ;
 DEFINE  INWARDLMON / ANALYSIS FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "&pre_month" ;
 DEFINE  DIFf_INWARD / Computed FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "Difference/In MTD" ;
 compute DIFF_INWARD;
 DIFF_INWARD = INWARDMON.sum - INWARDLMON.sum;
 endcomp;
 
 DEFINE  NCADAY / ANALYSIS FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "&yday" ;
 DEFINE  NCAMON / ANALYSIS FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "&month" ;
 DEFINE  NCALMON / ANALYSIS FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "&pre_month" ;
 DEFINE  DIFF_NCA / COMPUTED FORMAT= BEST12. WIDTH=12    SPACING=2   RIGHT "Difference/In MTD" ;
 compute DIFF_NCA;
 DIFF_NCA = NCAMON.sum - NCALMON.sum;
 endcomp;

 Break after SH_NAME/Summarize style={foreground=green};

 Break after TM_NAME/Summarize style={foreground=Red};

RBREAK after/Summarize style={foreground=Blue};

Compute after SH_NAME;
SH_NAME = trim(SH_NAME)||" Total";
endcomp;

Compute after TM_NAME;
TM_NAME = trim(TM_NAME)||" Total";
endcomp;

Compute after;
TM_NAME = "Grand Total";
endcomp;

 RUN;
 
