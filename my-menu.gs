function onOpen() {
  var ui = DocumentApp.getUi();
  ui.createMenu('<ЗАМЕНЫ>')
          .addItem('Помощь', 'showSidebar')
          .addSeparator()
          .addSubMenu(ui.createMenu('Замены')
            .addItem('Замены Правила 1', 'Zamena1')
            .addItem('Замены Правила 2', 'Zamena2')
            .addItem('Замены Правила 3', 'Zamena3')
            .addItem('Замены Правила 4', 'Zamena4')
            .addItem('Замена тире  150', 'Zamena5')
          )        
               
        
        .addSeparator()
        
        .addSubMenu(DocumentApp.getUi().createMenu('Системная')
              .addItem('Прочесть перед работой в этом меню', 'myAttention')
              .addSubMenu(DocumentApp.getUi().createMenu('Уверен, что тебе сюда надо?')
                          .addItem('Создание файла Правил 1', 'Start1')
                          .addItem('Создание файла Правил 2', 'Start2')
                          .addItem('Создание файла Правил 3', 'Start3')
                          .addItem('Создание файла Правил 4', 'Start4')
                          .addItem('Создание файла Правил 5', 'Start5')
                          ))

        
         
        
      .addToUi();
}



//*********************************************************************************************************************
// ПРИМЕР КАК СДЕЛАТЬ РЕГУЛЯРКУ ДЛЯ GOOGLE DOCS
// https://webapps.stackexchange.com/questions/99922/how-to-use-text-from-capturing-groups-in-google-docs-regex-replace
// !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


//****************** пункт меню применение замен к тексту

//function Zamena1()
//{
//   var nameFile = '_!_Pravila1'; // имя открываемого файла с правилами 1
//   mySearchReplace(nameFile); // запускаем функцию которая открывает файл с правилами замен и производит замены в активном тексте
//   
//}



function Zamena1()
{
   var nameFile = '_!_Pravila1'; // имя открываемого файла с правилами 1
   mySearchReplace(nameFile); // запускаем функцию которая открывает файл с правилами замен и производит замены в активном тексте
   
}

function Zamena2()
{
   var nameFile = '_!_Pravila2'; // имя открываемого файла с правилами 1
   mySearchReplace(nameFile); // запускаем функцию которая открывает файл с правилами замен и производит замены в активном тексте
   
}

function Zamena3()
{
   var nameFile = '_!_Pravila3'; // имя открываемого файла с правилами 1
   mySearchReplace(nameFile); // запускаем функцию которая открывает файл с правилами замен и производит замены в активном тексте
   
}

function Zamena4()
{
   var nameFile = '_!_Pravila4'; // имя открываемого файла с правилами 1
   mySearchReplace(nameFile); // запускаем функцию которая открывает файл с правилами замен и производит замены в активном тексте
   
}

function Zamena5()
{
   var nameFile = '_!_Pravila5'; // имя открываемого файла с правилами 1
   mySearchReplace(nameFile); // запускаем функцию которая открывает файл с правилами замен и производит замены в активном тексте
   
}

//****************** тело функции -- обработка замен, парсим правила и делаем замены в тексте 
// Cоздание файла и его заполнение файла правилами замен из предустановленного массива в функции Start
// передаем имя файла и массив значений что и на что будем менять


function mySearchReplace(nameFile)
{
  var foundFile = DriveApp.getFilesByName(nameFile).hasNext(); // нашли его на диске и в переменную foundFile передали false или true
    if (foundFile == true) // проверяем нашли файл или нет, если нашли то 
      {
         // ветка НАШЛИ файл - значит НЕ создаем его 
         // Logger.log('файл найден и поэтому НЕ создан');
         
         var fileRead = DriveApp.getFilesByName(nameFile); // получаем файл по имени
         var fileId = fileRead.next().getId(); // получаем ID файла
         var file = DocumentApp.openById(fileId); // открываем файл с правилами замен по ID 
         var content = file.getBody().getParagraphs(); // считываем содержимое файла 
         
         for(var i=0; i<content.length; i++) // парсим считанные данные по разделителю @
           {
             var myString = content[i].getText();
             var rules = myString.split('@')
             if(rules.length > 1)
               {
                 var activeDoc = DocumentApp.getActiveDocument(); // открываем активный документ 
                 var body =  activeDoc.getBody();
                 var paragraphs = body.getParagraphs();
                 //Logger.log(paragraphs.length); 
                 for (var j=0; j<paragraphs.length; j++) {
                   var text = paragraphs[j].getText();
                   var regexp = new RegExp(rules[0], "g"); // пределываем правило замен (что менять) в Regформат иначе не работает 
                   //Logger.log(regexp);
                   paragraphs[j].replaceText("(.*)", text.replace(regexp, rules[1])); // производим замены при помощи регулярки - первое поле рег выражение, полученное через преообразование , второе поле рег выражение в явном виде
                 }
                 showDialogReplaceOk(); // выводим соощение замены произведены
                 // activeDoc.getBody().replaceText(rules[0], rules[1]); // проводим замены в документе 
                 // Logger.log(content[i].getText()); 
                 // Logger.log(rules[0]); 
                 // Logger.log(rules[1]); 
               }           
                
           }
          
       }
     else 
        { 
          // ветка НЕ нашли файл - значит можно сделать чтобы писать файл с правилами замен не найден создайте его
          showDialogfilenotfound();
        }
}

// **************** myCreateFileAndWritePravila(nameFile, rules)
// создание файлов с правилами замен
// получает имя файла и список правил
// создает файл с именем'_!_Pravila1 (2,3,4,5)и пишет туда правила замен
// 

function myCreateFileAndWritePravila(nameFile, rules)
{
    
    // Logger.log(nameFile);
    
    var foundFile = DriveApp.getFilesByName(nameFile).hasNext(); // нашли его на диске и в переменную foundFile передали false или true
    
    // Logger.log('продолжаем');
    // Logger.log(foundFile);
    
    if (foundFile === false) // проверяем нашли файл или нет, так как файл НЕ нашли  - значит СОЗДАЕМ его и заполняем правилами замен
       {
          var doc = DocumentApp.create(nameFile); // создали файл с именем nameFile
          //Logger.log('файл не найден и поэтому СОЗДАН');          
           
          AlertCriate(); // сообщение что создан
          
          // распарсиваем массив правил из предустановленной переменной массив rules в стартовой фунции 
           for(var i=0; i<rules.length; i++)
           {
              var find = rules[i].find;
              var replace = rules[i].replace;
              var pravilo = find  + '@' + replace; // формируем правило 
              doc.getBody().appendParagraph(pravilo); // заносим правило в файл
              // Logger.log(find);
              // Logger.log(replace);
            }
        }
        else 
        {
          FileFound();
        }
 }



//****************** пункт меню создание файлов с правилами замен

// ------ Создание 1 нового файла с правилами замен

function Start1() 
{
  var nameFile = '_!_Pravila1'; // имя создаваемого файла

  var rules = 
  [
   // замена табуляции
          {"find":'test',"replace":'Ok'}, 
  // замена табуляции
          {"find":'(\\t)',"replace":''}, 
  
  // замена буквы тире на алт 0151
          {"find":'(^–|^-)',"replace":'—'},          
  // замена буквы ё
          {"find":'ё',"replace":'e'},
  // замена буквы тире с пробелом на алт 0151 с пробелом
          {"find":'( - )'  ,"replace":' — '},
  
  // замена открывающих кавычек
          {"find":'("|”|“|\‘|\’|\')([А-Яа-яA-Za-z0-9_])',"replace":'«$2'},
  
  // замена закрывающих кавычек        
          {"find":'("|”|“|\‘|\’|\')([,| |.|?|!])',"replace":'»$2'}
  
  ]
      
// {"find":"что","replace":"на что"},{"find":"ё","replace":"е"},

   
  myCreateFileAndWritePravila(nameFile, rules);
}

// ------ Создание 2-го нового файла с правилами замен

function Start2() 
{
  var nameFile = '_!_Pravila2'; // имя создаваемого файла
  var rules = [{"find":"22","replace":"33"},{"find":"44","replace":"55"},{"find":"66","replace":"77"}]
  myCreateFileAndWritePravila(nameFile, rules);
  
}


// ------ Создание 3-го нового файла с правилами замен

function Start3() 
{
  var nameFile = '_!_Pravila3'; // имя создаваемого файла
  var rules = [{"find":"f","replace":"r"},{"find":"f","replace":"r"},{"find":"f","replace":"r"}]
  myCreateFileAndWritePravila(nameFile, rules);
}

// ------ Создание 4-го нового файла с правилами замен

function Start4() 
{
  var nameFile = '_!_Pravila4'; // имя создаваемого файла
  var rules = [{"find":"f","replace":"r"},{"find":"f","replace":"r"},{"find":"f","replace":"r"}]
  myCreateFileAndWritePravila(nameFile, rules);
}

// ------ Создание 5-го нового файла с правилами замен

function Start5() 
{
  var nameFile = '_!_Pravila5'; // имя создаваемого файла
  //var rules = [{"find":"f","replace":"r"},{"find":"f","replace":"r"},{"find":"f","replace":"r"}]
  var rules = 
  [
   // замена табуляции
          {"find":'test',"replace":'Ok'}, 
  // замена табуляции
          {"find":'(\\t)',"replace":''}, 
  // замена тире на алт 0151
          {"find":'(^–|^-)',"replace":'–'},          
  // замена тире с пробелом на алт 0151 с пробелом
          {"find":'( - )'  ,"replace":' – '},
  // замена буквы ё
          {"find":'ё',"replace":'e'},
  // замена открывающих кавычек
          {"find":'("|”|“|\‘|\’|\')([А-Яа-яA-Za-z0-9_])',"replace":'«$2'},
  // замена закрывающих кавычек        
          {"find":'("|”|“|\‘|\’|\')([,| |.|?|!])',"replace":'»$2'}
  
  ]
  myCreateFileAndWritePravila(nameFile, rules);
}



// ******************** диалоги

// сообщение что создан
function AlertCriate() 
    {
      DocumentApp.getUi().alert('Файл с правилами создан под именем _!_Pravila ');
    }
    
// сообщение что создан
function myAttention() 
    {
      DocumentApp.getUi().alert
      ('Это раздал предназначен для создания новых файлов с правилами замен. Файлы называются _!_Pravila и размещаются в корне Disk Google. Если файл уже есть новый создан не будет. Удалите старый и тогда можно создать новый');
    }

// Файл с правилами замен не найден
function showDialogfilenotfound() {
    DocumentApp.getUi().alert('Файл с правилами замен не найден. Самостоятельно сойдайте файл с правилами замен');
}

// Файл с правилами замен не найден
function FileFound() {
    DocumentApp.getUi().alert('Файл с правилами замен УЖЕ существует. Если Вы хотите создать новый то найдите на Google Disk файл с именем _!_Pravila и удалите его');
}



// замены произведены
function showDialogReplaceOk() {
  var html = HtmlService.createHtmlOutputFromFile('filenotfound')
      .setWidth(400)
      .setHeight(50);
  DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Замены выполнены. Все Ok!');
}
// ******************  боковое меню


function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Правила работы')
      .setWidth(300);
     DocumentApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}

