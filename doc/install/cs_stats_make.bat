@ECHO OFF

SET APP_HOME=C:\cs_excel

CD %APP_HOME%

SET JAVA_HOME=%APP_HOME%\\zulu-8\bin
SET CLASSPATH=%APP_HOME%\classes;

for %%f in (%APP_HOME%\lib\*.jar) do call set CLASSPATH=%%CLASSPATH%%%%f;


%JAVA_HOME%\java.exe -Xms32m -Xmx2024m -cp %CLASSPATH%  com.seomse.shop.cs.CsExcelStats %APP_HOME%\CS_Stats.xlsx
