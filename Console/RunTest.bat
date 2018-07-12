echo off

set projectLocation=%1
set testngxml=%2

cd %projectLocation%
mvn clean test -DsuiteXmlFile=%testngxml%

rem mvn validate -DsuiteXmlFile=TestNG.xml
rem mvn compile -DsuiteXmlFile=TestNG.xml
rem mvn test -DsuiteXmlFile=TestNG.xml
rem mvn package -DsuiteXmlFile=TestNG.xml
rem mvn integration-test -DsuiteXmlFile=TestNG.xml
rem mvn verify -DsuiteXmlFile=TestNG.xml
rem mvn install -DsuiteXmlFile=TestNG.xml
rem mvn deploy -DsuiteXmlFile=TestNG.xml
rem mvn clean
rem mvn site
rem mvn test -DsuiteXmlFile=TestNG.xml >> Console\Log.log

rem https://maven.apache.org/guides/getting-started/maven-in-five-minutes.html
rem https://mvnrepository.com/
rem set classpath=%projectLocation%\src\test\java;%projectLocation%\Library\*

rem Jenkins
rem http://localhost:8080/job/Trail_Demo/job/CH_MavenTestNG/build?token=DemoRun&PARAMETER=Value
rem http://localhost:8080/job/Trail_Demo/job/CH_MavenTestNG/buildWithParameters?token=DemoRun&PARAMETER=TestNG_SIT.xml
