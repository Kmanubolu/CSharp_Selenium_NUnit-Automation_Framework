cd\
cd D:\Selenium\Libs
java -Dwebdriver.chrome.driver=D:\Selenium\Drivers\chromedriver.exe -jar selenium-server-standalone-3.12.0.jar -role webdriver -hub http://localhost:4444/grid/register -port 5566 -browser browserName=chrome,maxInstances=2,maxSession=2