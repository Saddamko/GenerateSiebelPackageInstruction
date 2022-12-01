/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package ua.biz.myshop.generatesiebelpackageinstruction;

/**
 *
 * @author boss
 */
// Java program to create a Word document
// Importing Spire Word libraries
import com.spire.doc.Document;
import com.spire.doc.FieldType;
import com.spire.doc.FileFormat;
import com.spire.doc.HeaderFooter;
import com.spire.doc.Section;
import com.spire.doc.ShapeHorizontalAlignment;
import com.spire.doc.ShapeVerticalAlignment;
import com.spire.doc.Table;
import com.spire.doc.documents.BorderStyle;
import com.spire.doc.documents.BreakType;
import com.spire.doc.documents.BuiltinStyle;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.HorizontalOrigin;
import com.spire.doc.documents.PageSize;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;
import com.spire.doc.documents.TextAlignment;
import com.spire.doc.documents.TextWrappingStyle;
import com.spire.doc.documents.TextWrappingType;
import com.spire.doc.documents.VerticalOrigin;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TextRange;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.siebel.data.*;
import com.spire.doc.CaptionNumberingFormat;
import com.spire.doc.CaptionPosition;
import com.spire.doc.ProtectionType;
import com.spire.doc.documents.TextDirection;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Properties;
import java.util.logging.Level;
import java.util.logging.Logger;

class GenerateSiebelPackageInstruction {

    private static String pkgNo ="290";
    private static final SimpleDateFormat formatterCurrentDate = new SimpleDateFormat("yyyy_MM_dd");
    private static final Date dateCurrent = new Date();
    private static final String sDateCurrent = formatterCurrentDate.format(dateCurrent);
    private static final String FILE_NAME = "e:\\ProjectChangedObjects" + pkgNo;
    private static final String sPathBase = "C:\\Areon\\Configuration\\"+pkgNo+ "_" + sDateCurrent+"\\";
    private static final String FILE_OUT_NAME = sPathBase+"UNIQA. Інструкція до пакету PkgUnq"+pkgNo;
    static boolean bTable = false, bTasks=false, bWorkflowProcess = false, bIntegrationObject = false, bPicture = false, bSystemPreferences = false, bLOV = false, bStateModel = false, bDataMap = false, bJob = false, bEIMConfigFile = false;
    static boolean bEAIDataMap = false, bCommPackage = false, bProfileConfiguration = false, bComponentDefinitions = false, bBusinessRole = false, bManifestFile = false, bManifestObject = false, bJavaScriptFile = false, bSQL = false;
    static boolean bSRF=true, bSavedQueries=false;
    static ArrayList<TParams> tParamsList = new ArrayList<TParams>();
    static ArrayList<TWorkflows> tWFList = new ArrayList<TWorkflows>();
    static ArrayList<TTasks> tTasksList = new ArrayList<TTasks>();
    static ArrayList<TEAIDataMap> tEAIDataMapList = new ArrayList<TEAIDataMap>();
    static ArrayList<TSysPref> tSysPrefList = new ArrayList<TSysPref>();
    static ArrayList<TSQL> tSQLList = new ArrayList<TSQL>();
    static ArrayList<TEIM> tEIMList = new ArrayList<TEIM>();
    static ArrayList<TLOV> tLOVList = new ArrayList<TLOV>();
    static ArrayList<TStateModel> tStateModelList = new ArrayList<TStateModel>();
    static ArrayList<TCommPkg> tCommPkgList = new ArrayList<TCommPkg>();
    static ArrayList<TJavaScript> tJavaScriptList = new ArrayList<TJavaScript>();
    static ArrayList<TPDQ> tPDQList = new ArrayList<TPDQ>();
    static List<String> tTablesList = new ArrayList<String>();
    static List<String> tIOList = new ArrayList<String>();
    static  String SiebelConnectString ="";
    static  String SiebelUser="";
    static  String SiebelUserPassword="";
    static int nT; //Counter of Tables
    static int nR; //Counter of Repository Objects

    // Main driver method
    @SuppressWarnings("empty-statement")
    public static void main(String[] args) throws Exception {
        ReadExcelFile();
        getProperties();
        
//        if (args[0].isEmpty()==false) { pkgNo = args[0];};
        
        getTaskList(tParamsList);
    
        // create a Word document
        Document document = new Document();
        
        CreateCatalog ( sPathBase);
        
//        document.protect(ProtectionType.Allow_Only_Reading);
//        document.setProtectionType(ProtectionType.Allow_Only_Reading);

        // Customize a paragraph style
        ParagraphStyle style = new ParagraphStyle(document);
        // Paragraph name
        style.setName("paraStyle");
        // Paragraph format
        style.getCharacterFormat().setFontName("Calibri");
        // Paragraph font size
        style.getCharacterFormat().setFontSize(11f);
        style.getParagraphFormat().setFirstLineIndent(20);
        style.getParagraphFormat().setTextAlignment(TextAlignment.Auto);
        style.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
        // Adding styles using inbuilt method
        document.getStyles().add(style);
        
         // Customize a paragraph style
        ParagraphStyle Heading1style = new ParagraphStyle(document);
        // Paragraph name
        Heading1style.setName("myHeading_1");
        // Paragraph format
        Heading1style.getCharacterFormat().setFontName("Calibri");
        // Paragraph font size
        //Heading1style.getListFormat().applyNumberedStyle();
        Heading1style.getCharacterFormat().setFontSize(12f);
        Heading1style.getCharacterFormat().setAllCaps(true);
        Heading1style.getCharacterFormat().setBold(true);
        Heading1style.getParagraphFormat().setFirstLineIndent(0);
        Heading1style.getParagraphFormat().setTextAlignment(TextAlignment.Auto);
        Heading1style.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
        Heading1style.getListFormat().setListLevelNumber(1);
        // Adding styles using inbuilt method
        document.getStyles().add(Heading1style);
        
        // Customize a paragraph style
        ParagraphStyle Heading2style = new ParagraphStyle(document);
        // Paragraph name
        Heading2style.setName("myHeading_2");
        // Paragraph format
        Heading2style.getCharacterFormat().setFontName("Calibri");
        // Paragraph font size
        //Heading2style.getListFormat().applyNumberedStyle();
        Heading2style.getCharacterFormat().setFontSize(12f);
        Heading2style.getCharacterFormat().setBold(true);
        Heading2style.getParagraphFormat().setFirstLineIndent(0);
        Heading2style.getParagraphFormat().setTextAlignment(TextAlignment.Auto);
        Heading2style.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
        Heading2style.getListFormat().setListLevelNumber(2);
        // Adding styles using inbuilt method
        document.getStyles().add(Heading2style);

        // Customize a paragraph style
        ParagraphStyle Name = new ParagraphStyle(document);
        // Paragraph name
        Name.setName("nameStyle");
        // Paragraph format
        Name.getCharacterFormat().setFontName("Calibri");
        Name.getCharacterFormat().setBold(true);
        // Paragraph font size
        Name.getCharacterFormat().setFontSize(15f);
        Name.getParagraphFormat().setFirstLineIndent(00);
        Name.getParagraphFormat().setTextAlignment(TextAlignment.Center);
        Name.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        // Adding styles using inbuilt method
        document.getStyles().add(Name);

        // Customize a table paragraph style
        ParagraphStyle TableText = new ParagraphStyle(document);
        // Paragraph name
        TableText.setName("TableText");
        // Paragraph format
        TableText.getCharacterFormat().setFontName("Calibri");
        // Paragraph font size
        TableText.getCharacterFormat().setFontSize(10f);
        TableText.getParagraphFormat().setFirstLineIndent(00);
        TableText.getParagraphFormat().setTextAlignment(TextAlignment.Center);
        TableText.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
        // Adding styles using inbuilt method
        document.getStyles().add(TableText);

        // Customize a table paragraph style
        ParagraphStyle TableHeader = new ParagraphStyle(document);
        // Paragraph name
        TableHeader.setName("TableHeader");
        // Paragraph format
        TableHeader.getCharacterFormat().setFontName("Calibri");
        TableHeader.getCharacterFormat().setBold(true);
        // Paragraph font size
        TableHeader.getCharacterFormat().setFontSize(10f);
        TableHeader.getParagraphFormat().setFirstLineIndent(00);
        TableHeader.getParagraphFormat().setTextAlignment(TextAlignment.Center);
        TableHeader.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        // Adding styles using inbuilt method
        document.getStyles().add(TableHeader);

        // Add a section
        Section section = document.addSection();

        //insert header and footer
        insertHeaderAndFooter(section);

        // Add a heading
        Paragraph heading = section.addParagraph();

        heading.appendText("КЕРІВНИЦТВО");
        heading.appendBreak(BreakType.Line_Break);
        heading.appendText("щодо перенесення конфігурації Siebel CRM");
        heading.appendBreak(BreakType.Line_Break);
        heading.appendField("Зміст", FieldType.Field_TOC);

        // Add a subheading
        Paragraph subheading_table = section.addParagraph();
        subheading_table.appendText("Таблиця змін");

        Table tableChanges = section.addTable(true);
        SimpleDateFormat formatter = new SimpleDateFormat("dd.MM.yyyy");
        Date date = new Date();
        String[][] dataChanges
                = {
                    new String[]{"Дата", "Автор", "Зміни"},
                    new String[]{formatter.format(date), "Вашурін В.", "Створення стартової версії документа"},
                    new String[]{"", "", ""},};

        int rowCountChanges = dataChanges.length;
        int columnCountChanges = dataChanges[0].length;
        tableChanges.resetCells(rowCountChanges, columnCountChanges);

        //fill the data to Table
        for (int i = 0; i < rowCountChanges; i++) {
            for (int j = 0; j < columnCountChanges; j++) {
                Paragraph p;
                p = tableChanges.getRows().get(i).getCells().get(j).addParagraph();
                p.appendText(dataChanges[i][j]);
                if (i == 0) {
                    p.applyStyle("TableHeader");
                } else {
                    p.applyStyle("TableText");
                };
            }
        }

        // Adding sub-headings
        // two paragraph under the first subheading
        Paragraph subheading_common = section.addParagraph();
        subheading_common.appendText("Загальні вимоги");

        // Adding paragraph 1
        Paragraph para_z = section.addParagraph();
        para_z.appendText(
                "Роботи зі встановлення пакета повинен самостійно виконувати Адміністратор системи Siebel CRM "
                        + "або спільно з фахівцями Ареон, які мають достатню компетенцію щодо розгортання функціоналу пакета, "
                        + "за погодженням з власниками бізнес-процесів, відповідно до цієї інструкції "
                        + "та прийнятого регламенту обслуговування системи.");

        Paragraph subheading_goal = section.addParagraph();
        subheading_goal.appendText("Призначення пакета (Release notes)");
        // Adding paragraph 2
        Paragraph para_goal_text = section.addParagraph();
        para_goal_text.appendText(
                "#Release Notes");
        para_goal_text.appendCheckBox();

        Paragraph subheading_opys = section.addParagraph();
        subheading_opys.appendText("Опис пакету");

        Paragraph para_object = section.addParagraph();
        para_object.appendText("Пакет  містить об'єкти:");

        Table tableObjects = section.addTable(true);
        String[][] dataObjects
                = {
                    new String[]{"Категорія об'єкту", "Тип зміненого об'єкту", "Назва зміненого об'єкту"},};

        int rowCountObjects = tParamsList.size() ;
        int columnCountObjects = 3;
        tableObjects.resetCells(rowCountObjects+1, columnCountObjects);
        //fill the header to Table
        int i = 0;
        for (int j = 0; j < 3; j++) {
            Paragraph p;
            p = tableObjects.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataObjects[i][j]);
        }

        //fill the data to Table
        for (i = 0; i < rowCountObjects; i++) {
            for (int j = 0; j < columnCountObjects; j++) {
                Paragraph p;
                p = tableObjects.getRows().get(i+1).getCells().get(j).addParagraph();
                if (j == 0) {
                    p.appendText(tParamsList.get(i).category);
                } else if (j == 1) {
                    p.appendText(tParamsList.get(i).type);
                } else if (j == 2) {
                    p.appendText(tParamsList.get(i).name);
                }
                p.applyStyle("TableText");
            }
        }
        
        
        //BEGIN COMPILE SRF
        if (bSRF)
        {
        Paragraph subheading_import = section.addParagraph();
        subheading_import.appendText("Імпорт репозиторних об'єктів");
        CreateCatalog(sPathBase+"200-repo");
        CreateCatalog(sPathBase+"200-repo\\overwrite");
        

        // Adding one paragraph under the second subheading
        Paragraph para_repo = section.addParagraph();
        para_repo.appendText(
                "На тестовому, а після закінчення перевірки, на продуктивному середовищі, імпортувати в Siebel Tools SIF з папки "
                + "з вмістом пакета в режимі пакетної обробки Siebel Tools. "
                + "Копіювати файли пакета в папку "+sPathBase+ ", зберігши структуру каталогів.  "
                + "Дати повні права на папку " + sPathBase +" для всіх користувачів.\n"
                + "Пакетний імпорт об'єктів виконується:\n"
                + "1.	З папки "+ sPathBase + "\\200-repo\\overwrite в режимі імпорту \"overwrite\":\n"
                + "Командний рядок (в режимі Адміністратор) режиму Overwrite (зразок для TEST середовища, для PROD необхідно змінити пароль):\n"
                + "C:\\Siebel\\16.0.0.0.0\\Tools\\BIN\\siebdev.exe /c \"c:\\Siebel\\16.0.0.0.0\\Tools\\bin\\enu\\tools.cfg\" /u SADMIN /p UNQIP2016 /d ServerDataSrc /batchimport \"Siebel Repository\" overwrite \"" + sPathBase +"200-repo\\overwrite\" "+ sPathBase + "200-repo\\UnqPkg"+pkgNo+"overwrite.log\n"
                + "Після закінчення імпорту потрібно обов'язково досліджувати вміст файлів " + sPathBase +"200-repo\\UnqPkg"+pkgNo+"overwrite.log,\n"
                + "на наявність помилок (див. останній рядок у цих файлах).\n"
                + "Повинна бути підстрока Failed Imports: 0 (без повідомлень про помилки), приклад:\n"
                + "STATUS: Total Files:  "+ Integer.toString(nR) +", Successful Imports: "+ Integer.toString(nR) +", Failed Imports: 0\n"
                + "2.	З папки "+ sPathBase + "\\200-repo\\merge (описи таблиць) в режимі імпорту \"merge\":\n"
                + "Командний рядок (в режимі Адміністратор) режиму Merge (зразок для TEST середовища, для PROD необхідно змінити пароль):\n"
                + "C:\\Siebel\\16.0.0.0.0\\Tools\\BIN\\siebdev.exe /c \"c:\\Siebel\\16.0.0.0.0\\Tools\\bin\\enu\\tools.cfg\" /u SADMIN /p UNQIP2016 /d ServerDataSrc /batchimport \"Siebel Repository\" merge \"" + sPathBase +"200-repo\\merge\" "+ sPathBase + "200-repo\\UnqPkg"+pkgNo+"merge.log\n"
                + "Після закінчення імпорту потрібно обов'язково досліджувати вміст файлів " + sPathBase +"200-repo\\UnqPkg"+pkgNo+"overwrite.log,\n"
                + "на наявність помилок (див. останній рядок у цих файлах).\n"
                + "Повинна бути підстрока Failed Imports: 0 (без повідомлень про помилки), приклад:\n"
                + "STATUS: Total Files:  "+ Integer.toString(nT) +", Successful Imports: "+ Integer.toString(nT) +", Failed Imports: 0\n"
        );

        Paragraph subheading_compile = section.addParagraph();
        subheading_compile.appendText("Компіляція SRF-файлу (репозиторія)");

        // Adding one paragraph under the second subheading
        Paragraph para_compile = section.addParagraph();
        para_compile.appendText(
                  "Далі необхідно провести компіляцію SRF-файлу, щоб у нього увійшли імпортовані з файлів .sif зміни. "
                + "Компіляцію зробити для ENU та RUS мов (зробити два окремих SRF-файли).\n"
                + "Перед компіляцією необхідно перевірити встановлення мови у налаштуваннях Siebel Tools, "
                + "вибрати всі проекти та обрати файл, до якого вивантажити репозиторій.");

        DocPicture picture2 = para_compile.appendPicture("src\\main\\resources\\ToolsCompile2.png");
        picture2.setWidth(400);
        picture2.setHeight(300);
        picture2.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        picture2.addCaption("Екран під час компіляції SRF-файлу (репозиторія)", CaptionNumberingFormat.Alphabetic, CaptionPosition.Below_Item);
        subheading_import.applyStyle("myHeading_1");
        subheading_compile.applyStyle("myHeading_2");
        para_repo.applyStyle("paraStyle");
        para_compile.applyStyle("paraStyle");
        }
        //END COMPILE SRF
        
                //BEGIN SAVED/PREDIFINED QUERIES
        if (bSavedQueries)
        {
        getPDQList(tParamsList);
        Paragraph subheading_predifined_queries = section.addParagraph();
        subheading_predifined_queries.appendText("Налаштування Predefined Queries");
        
        // Adding one paragraph under the second subheading
        Paragraph predifined_queries = section.addParagraph();
        predifined_queries.appendText(
                  "Попередньо визначені запити (PDQ) автоматизують запити, які користувач може виконувати онлайн. "
                          + "Замість того, щоб створювати запит, вводити критерії та запускати запит, "
                          + "користувач вибирає PDQ зі спадного списку Запити. "
                          + "Якщо ви хочете зробити запит загальнодоступним:\n" +
                        "Перейдіть до екрана «Administration - Application screen», а потім до перегляду «Predefined Queries».\n" +
                        "У списку «Predefined Queries» зніміть прапорець у полі «Private» в записі для щойно створеного запиту."
                          + "Зайдить до Administration-Application, Predifined queries, та додайте наступні записи");
            Table tablePDQ = section.addTable(true);
            String[][] dataStateModel
                    = {new String[]{"Type", "Object", "Name","Query"},};

            int rowCountPDQ = tPDQList.size();

            int columnCountPDQ = 4;
            tablePDQ.resetCells(rowCountPDQ+1, columnCountPDQ);

            //fill the header to tableIO
            i = 0;
            for (int j = 0; j < columnCountPDQ; j++) {
                Paragraph p;
                p = tablePDQ.getRows().get(i).getCells().get(j).addParagraph();
                p.applyStyle("TableHeader");
                p.appendText(dataStateModel[i][j]);
            }

            for (i = 0; i < rowCountPDQ; i++) {
                for (int j = 0; j < columnCountPDQ; j++) {
                    Paragraph p;
                    p = tablePDQ.getRows().get(i+1).getCells().get(j).addParagraph();
                    p.applyStyle("TableText");
                    if (j == 0) {
                        p.appendText(tPDQList.get(i).type);
                    } else if (j == 2) {
                        p.appendText(tPDQList.get(i).name);
                    } else if (j == 4) {
                        p.appendText(tPDQList.get(i).script);                        
                    }
                }

            }   
        
        DocPicture picture = predifined_queries.appendPicture("src\\main\\resources\\PDQ.png");
        picture.setWidth(400);
        picture.setHeight(120);
        picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);

        picture.addCaption("Приклад налаштування PDQ", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);
        subheading_predifined_queries.applyStyle("myHeading_1");
        predifined_queries.applyStyle("paraStyle");    
        }
        //END SAVED/PREDIFINED QUERIES


        
        //BEGIN STATE MODEL
        if (bStateModel)
        {
            getStateModelList(tParamsList);
            Paragraph subheading_StateModel = section.addParagraph();
            subheading_StateModel.appendText("Перенос State Model");
            
            CreateCatalog(sPathBase+"300-environment");
            CreateCatalog(sPathBase+"300-environment\\303-State Model");
            
            Paragraph para_StateModel = section.addParagraph();
            para_StateModel.appendText(
                      "Далі імпортуйте State Model. Щоб імпортувати State Model, скористайтеся Application Deployment Manager. "
                    + "Відкрийте клієнт Siebel (ENU) із правами адміністратора. Перейдіть до екрана «Application Deployment Manager», "
                    + "а потім до «Deployment Sessions». Далі, в меню виберіть \"Deploy From Local File\" і папку, "
                    + "де збережені файли, що містять State Model: " +sPathBase+ "300-environment\\303-State Model.\n"
                    + "Необхідно по черзі виконати вказані дії для всіх файлів (див. таблицю нижче)." );
            Table tableStateModel = section.addTable(true);
            String[][] dataStateModel
                    = {new String[]{"State Model", "XML-файл"},};

            int rowCountStateModel = tStateModelList.size();

            int columnCountStateModel = 2;
            tableStateModel.resetCells(rowCountStateModel+1, columnCountStateModel);

            //fill the header to tableIO
            i = 0;
            for (int j = 0; j < columnCountStateModel; j++) {
                Paragraph p;
                p = tableStateModel.getRows().get(i).getCells().get(j).addParagraph();
                p.applyStyle("TableHeader");
                p.appendText(dataStateModel[i][j]);
            }

            for (i = 0; i < rowCountStateModel; i++) {
                for (int j = 0; j < columnCountStateModel; j++) {
                    Paragraph p;
                    p = tableStateModel.getRows().get(i+1).getCells().get(j).addParagraph();
                    p.applyStyle("TableText");
                    if (j == 0) {
                        p.appendText(tStateModelList.get(i).name);
                    } else if (j == 1) {
                        p.appendText(tStateModelList.get(i).file);
                    }
                }

            }   
            
            subheading_StateModel.applyStyle("myHeading_1");
            para_StateModel.applyStyle("paraStyle");
            
        }
        //END STATE MODEL
        
        //BEGIN IO
        if (bIntegrationObject)
        {
        getIOList(tParamsList);
        Paragraph subheading_IO = section.addParagraph();
        subheading_IO.appendText("Налаштування по Integration Objects");

        Paragraph para_IO = section.addParagraph();
        para_IO.appendText("1. Підключіться за допомогою Siebel Tools.\n" +
            "2. У Object Explorer у Siebel Tools виберіть Integration Object, після чого з’явиться список Integration Objects.\r" +
            "3. Клацніть правою кнопкою миші об’єкт інтеграції, який потрібно розгорнути, а потім виберіть «Deploy to run-time Database».\r" +
            "4. Об'єкт інтеграції розгортається.\n" +
            "5. У клієнті Siebel перейдіть до екрана «Administration-Web Services», перегляд «Inbound (або Outbound) Web Services ».\r" +
            "6. Клацніть «Clear Cache», щоб зробити недійсними визначення об’єкта інтеграції та веб-служб у базі даних часу виконання.\r" +
            "Виконати Undeploy/Deploy для:" );

        Table tableIO = section.addTable(true);
        String[][] dataIO
                = {new String[]{"Тип об'єкту", "Назва"},};

        int rowCountIO = tIOList.size();

        int columnCountIO = 2;
        tableIO.resetCells(rowCountIO+1, columnCountIO);

        //fill the header to tableIO
        i = 0;
        for (int j = 0; j < columnCountIO; j++) {
            Paragraph p;
            p = tableIO.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataIO[i][j]);
        }

        for (i = 0; i < rowCountIO; i++) {
            for (int j = 0; j < columnCountIO; j++) {
                Paragraph p;
                p = tableIO.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText("233-Integration Object");
                } else if (j == 1) {
                    p.appendText(tIOList.get(i));
                }
            }

        }   
        
        Paragraph para_IO2 = section.addParagraph();
        para_IO2.appendText("Для перевірки зайти до Deployed Integration Objects вкладки  экрану Siebel Web Services Administration" );
        subheading_IO.applyStyle("myHeading_1");
        para_IO.applyStyle("paraStyle");
        para_IO2.applyStyle("paraStyle");
        
        DocPicture picture = para_IO2.appendPicture("src\\main\\resources\\io.png");
        picture.setWidth(450);
        picture.setHeight(150);
        picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        picture.addCaption("Екран під час налаштування по Integration Objects", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);
        }
        //END IO

        //BEGIN TABLES
        if (bTable)
        {
        CreateCatalog(sPathBase+"200-repo\\merge");    
        getTablesList(tParamsList);
        Paragraph subheading_tables = section.addParagraph();
        subheading_tables.appendText("Apply та Activate таблиць");

        Paragraph para_tables = section.addParagraph();
        para_tables.appendText("У фізичну схему бази даних необхідно внести зміни, "
                + "відповідно до нових описів таблиць в репозиторії (виконати Apply). У Siebel Tools знайдіть таблиці");

        Table tableTables = section.addTable(true);
        String[][] dataTable
                = {new String[]{"Тип об'єкту", "Назва"},};

        int rowCountTable = dataTable.length;

        rowCountTable = tTablesList.size();

        int columnCountTable = 2;
        tableTables.resetCells(rowCountTable+1, columnCountTable);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountTable; j++) {
            Paragraph p;
            p = tableTables.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataTable[i][j]);
        }

        for (i = 0; i < rowCountTable; i++) {
            for (int j = 0; j < columnCountTable; j++) {
                Paragraph p;
                p = tableTables.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText("201-Table");
                } else if (j == 1) {
                    p.appendText(tTablesList.get(i));
                }
            }

        }

        Paragraph para_tables2 = section.addParagraph();
        para_tables2.appendText("та натисніть кнопку «Apply» під користувачем: SIEBEL/SIBIP2016; DSN: SIEBTEST2016_DSN (для тестового середовища), або SIEBPROD2016_DSN для продуктивного середовища. "
                + "Table Space: SIEBEL_DATA, Index Space: SIEBEL_INDEX для всіх середовищ. "
                + "Провести операцію Activate.");
        DocPicture picture = para_tables2.appendPicture("src\\main\\resources\\Apply.png");
        picture.setWidth(250);
        picture.setHeight(350);
        picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        picture.addCaption("Екран Apply та Activate таблиць", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);
        subheading_tables.applyStyle("myHeading_1");
        para_tables.applyStyle("paraStyle");
        para_tables2.applyStyle("paraStyle");
        }
        //END TABLES
        
        //BEGIN LOVS
        if (bLOV)
        {
        getLOVList(tParamsList);
        Paragraph subheading_lov = section.addParagraph();
        subheading_lov.appendText("Встановлення значень List of Values (LOV)");

        Paragraph para_lov = section.addParagraph();
        para_lov.appendText(
                  "Імпортуйте LOV. Щоб імпортувати LOV, скористайтеся Application Deployment Manager. "
                + "Відкрийте клієнт Siebel (ENU) із правами адміністратора. "
                + "Перейдіть до екрана «Application Deployment Manager), "
                + "а потім до «Deployment Sessions». Далі, в меню виберіть  "
                + "\"Deploy From Local File\" і папку, де збережені файли, що містять LOV: "
                + sPathBase +"300-environment\\302-List Of Values.\n" 
                + "Необхідно по черзі виконати вказані дії для всіх файлів (див. таблицю нижче).");   
        
        Table tableLOV = section.addTable(true);
        String[][] dataLOV
                = {new String[]{"Назва LOV", "XML-файл"},};

        int rowCountTableLOV = tLOVList.size();

        int columnCountTableLOV = 2;
        tableLOV.resetCells(rowCountTableLOV+1, columnCountTableLOV);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountTableLOV; j++) {
            Paragraph p;
            p = tableLOV.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataLOV[i][j]);
        }

        for (i = 0; i < rowCountTableLOV; i++) {
            for (int j = 0; j < columnCountTableLOV; j++) {
                Paragraph p;
                p = tableLOV.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText(tLOVList.get(i).name);
                } else if (j == 1) {
                    p.appendText(tLOVList.get(i).file);
                }
            }
        }
        Paragraph para_lov2 = section.addParagraph();
        para_lov2.appendText(
                  "Далі, перейдіть до екрана «Administration Data», а потім до «List of Values», і натисніть кнопку «Clear Cache».");     
        
        subheading_lov.applyStyle("myHeading_1");
        para_lov.applyStyle("paraStyle");  
        para_lov2.applyStyle("paraStyle");
        }
        //END LOVS

        //BEGIN WORKFLOWS
        if (bWorkflowProcess)
        {
        getWorkflowList(tParamsList);
        UpdateWorkflowVersionList();
        Paragraph subheading_wf = section.addParagraph();
        subheading_wf.appendText("Активація потоку операцій (Workflow)");

        Paragraph para_wf = section.addParagraph();
        para_wf.appendText("Щоб активувати потік операцій:  увійдіть у клієнт Siebel (ENU) з правами адміністратора. "
                + "Перейдіть до екрана Administration – Business Process», а потім до Workflow Deployment». "
                + "На аплеті Repository \"Workflow Process\" запитайте в полі Name\" потрібний Workflow і "
                + "натисніть кнопку \"Activate\". Після активації, необхідно перевірити: "
                + "версія активного Workflow має збігатися з версією репозиторії.");

        Table tableWF = section.addTable(true);
        String[][] dataWF
                = {new String[]{"Тип об'єкту", "Потік операцій (Workflow)", "Версія репозиторію"},};

        int rowCountWF = dataWF.length;

        rowCountWF = tWFList.size();

        int columnCountWF = 3;
        tableWF.resetCells(rowCountWF+1, columnCountWF);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountWF; j++) {
            Paragraph p;
            p = tableWF.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataWF[i][j]);
        }

        for (i = 0; i < rowCountWF; i++) {
            for (int j = 0; j < columnCountWF; j++) {
                Paragraph p;
                p = tableWF.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText("Workflow Process");
                } else if (j == 1) {
                    p.appendText(tWFList.get(i).name);
                } else if (j == 2) {
                    p.appendText(tWFList.get(i).version.toString());
                }
            }
        }

        //fill the style to Table
        para_wf.appendText("У Siebel Tools перевірити: якщо встановлено статус Completed для різних версій одного і того ж Workflow, "
                + "то необхідно змінити статус на Not In Use. Тобто, у статусі Completed повинен бути Workflow останньої версії.\n"
                + "При активації WF в клієнті видалити попередні неактивні версії, перезавантаження сервісів не потрібно.");
        DocPicture picture = para_wf.appendPicture("src\\main\\resources\\WF1.png");
        picture.setWidth(400);
        picture.setHeight(150);
        picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        picture.addCaption("Екран під час активації потоку операцій (Workflow)", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);
        
        subheading_wf.applyStyle("myHeading_1");
        para_wf.applyStyle("paraStyle");
        }
        // END WORKFLOWS
        
        //BEGIN TASKS
        if (bTasks)
        {
        getTaskList(tParamsList);
        UpdateTaskVersionList();
        Paragraph subheading_tsk = section.addParagraph();
        subheading_tsk.appendText("Налаштування Task");

        Paragraph para_tsk = section.addParagraph();
        para_tsk.appendText("Виконується після компіляції SRF-файлів та їх заміни."
                + " У Siebel Tools перевірити: якщо встановлено статус Completed для різних версій одного і того ж Task, "
                + "то необхідно змінити статус на Not In Use. Тобто. у статусі \"Completed\" повинен бути тільки Task останньої версії.\n" 
                + "Далі, щоб активувати завдання: увійдіть у клієнт Siebel (ENU) з правами адміністратора. "
                + "Перейдіть до екрана «Administration - Business Process», а потім до подання «Task Deployment». "
                + "У списку «Task Deployment» запитайте у полі «Name» потрібне завдання та натисніть кнопку «Activate». "
                + "Після активації необхідно перевірити: версія активного завдання має збігатися з версією репозиторії.");

        Table tableTsk = section.addTable(true);
        String[][] dataTsk
                = {new String[]{"Тип об'єкту", "Назва", "Версія репозиторію"},};

        int rowCountTsk = tTasksList.size();

        int columnCountTsk = 3;
        tableTsk.resetCells(rowCountTsk+1, columnCountTsk);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountTsk; j++) {
            Paragraph p;
            p = tableTsk.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataTsk[i][j]);
        }

        for (i = 0; i < rowCountTsk; i++) {
            for (int j = 0; j < columnCountTsk; j++) {
                Paragraph p;
                p = tableTsk.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText("215-Task UI");
                } else if (j == 1) {
                    p.appendText(tTasksList.get(i).name);
                } else if (j == 2) {
                    p.appendText(tTasksList.get(i).version.toString());
                }
            }
        }

        //fill the style to Table
        Paragraph para_tsk2 = section.addParagraph();
        para_tsk2.appendText("У Siebel Tools перевірити: якщо встановлено статус Completed для різних версій одного і того ж Task, "
                + "то необхідно змінити статус на Not In Use. Тобто, у статусі Completed повинен бути Task останньої версії.\n"
                + "При активації Task в клієнті видалити попередні неактивні версії, перезавантаження сервісів не потрібно.");
        
        subheading_tsk.applyStyle("myHeading_1");
        para_tsk.applyStyle("paraStyle");
        para_tsk2.applyStyle("paraStyle");
        }
        //END TASKS

        Paragraph subheading_env = section.addParagraph();
        subheading_env.appendText("Перенос оточення");
        
        //BEGIN OTHER
        if (bEIMConfigFile)
        {
        getEIMList(tParamsList);
        Paragraph subheading_EIM2 = section.addParagraph();
        subheading_EIM2.appendText("Інші налаштування");

        Paragraph para_eim = section.addParagraph();
        para_eim.appendText("Скопіюйте із заміною файли *.ifb з директорії "
                + sPathBase + "300-environment\\300-Other "
                + "до директорії C:\\Siebel\\16.0.0.0.0\\ses\\siebsrvr\\ADMIN siebelapp серверу");

        Table tableEIM = section.addTable(true);

        String[][] dataEIM
                = {new String[]{"Тип об'єкту", "Назва файлу",},};

        int rowCountEIM= tEIMList.size();
        int columnCountEIM = 2;
        tableEIM.resetCells(rowCountEIM+1, columnCountEIM);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountEIM; j++) {
            Paragraph p;
            p = tableEIM.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataEIM[i][j]);
        }

        for (i = 0; i < rowCountEIM; i++) {
            for (int j = 0; j < columnCountEIM; j++) {
                Paragraph p;
                p = tableEIM.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText(tEIMList.get(i).type);
                } else if (j == 1) {
                    p.appendText(tEIMList.get(i).name);
                }
            }
        }
        subheading_EIM2.applyStyle("myHeading_2");
        para_eim.applyStyle("paraStyle");
        
        }
        //END OTHER
        
        //BEGIN COMM_TEMPLATES
        if (bCommPackage)
        {
        getCommPkgList(tParamsList);
        Paragraph subheading_CommPkg = section.addParagraph();
        subheading_CommPkg.appendText("Імпорт шаблонів комунікацій");
        
        CreateCatalog(sPathBase+"300-environment\\323-Comm Package");

        Paragraph para_CommPkg = section.addParagraph();
        para_CommPkg.appendText(
                  "Щоб імпортувати Comm Package, відкрийте клієнт Siebel із правами адміністратора. "
                + "Перейдіть до екрана «Application Deploy Management», а потім до «Deployment Sessions». "
                + "Далі, в меню виберіть \"Deploy From Local File\" і папку, де збережений файл: "
                + sPathBase +"300-environment\\323-Comm Package.\n" +
                  "Для перевірки необхідно зайти до «Administration – Communications», "
                + "«All Templates» та виконати пошук шаблонів, зазначених у таблиці нижче:");
        Table tableCommPkg = section.addTable(true);

        String[][] dataCommPkg
                = {new String[]{"Тип об'єкту", "Назва шаблону комунікації","Назва файлу",},};

        int rowCountCommPkg = tCommPkgList.size();
        int columnCountCommPkg = 3;
        tableCommPkg.resetCells(rowCountCommPkg+1, columnCountCommPkg);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountCommPkg; j++) {
            Paragraph p;
            p = tableCommPkg.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataCommPkg[i][j]);
        }

        for (i = 0; i < rowCountCommPkg; i++) {
            for (int j = 0; j < columnCountCommPkg; j++) {
                Paragraph p;
                p = tableCommPkg.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText(tCommPkgList.get(i).type);
                } else if (j == 1) {
                    p.appendText(tCommPkgList.get(i).name);
                }
            }
        }
        subheading_CommPkg.applyStyle("myHeading_2");
        para_CommPkg.applyStyle("paraStyle");         
        }   
        //END   COMM_TEMPLATES
        
        //BEGIN MANIFEST
        if (bJavaScriptFile || bManifestFile || bManifestObject)
        {
        getJavaScriptList(tParamsList);
        Paragraph subheading_JavaScript = section.addParagraph();
        subheading_JavaScript.appendText("Маніфест");
        CreateCatalog(sPathBase+"300-environment\\395-Java Script File");

        Paragraph para_JavaScript = section.addParagraph();
        para_JavaScript.appendText(
                  "Перенести " +sPathBase +"300-environment\\395-Java Script File\\*.js,(дивись таблицю ниже), із збереженням поточної версії файлу  до C:\\Siebel\\16.0.0.0.0\\eappweb\\public\\SCRIPTS\\siebel\\custom на сервері siebelapp.\n" +
                    "В Administration Application - Manifest Administration створити запис\n" +
                    " \n" +
                    "	UI Objects:\n" +
                    "		Inactive = N\n" +
                    "		Type = Applet\n" +
                    "		Usage Type = Physical Renderer\n" +
                    "		Name = XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX\n" +
                    "	Object Expression:		\n" +
                    "		Inactive Flag = N\n" +
                    "		Level = 1\n" +
                    "	Files\n" +
                    "		Inactive = N\n" +
                    "		Name = XXXXXXXXXXXXXXXXXXXXXXXXXX.js");
        DocPicture picture = para_JavaScript.appendPicture("src\\main\\resources\\Манифест1.png");
        picture.setWidth(400);
        picture.setHeight(300);
        picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        picture.addCaption("Приклад налашування маніфесту", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);
        Table tableJavaScript = section.addTable(true);

        String[][] dataJavaScript
                = {new String[]{"Тип об'єкту", "Назва","Назва файлу",},};

        int rowCountJavaScript = tJavaScriptList.size();
        int columnCountJavaScript = 3;
        tableJavaScript.resetCells(rowCountJavaScript+1, columnCountJavaScript);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountJavaScript; j++) {
            Paragraph p;
            p = tableJavaScript.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataJavaScript[i][j]);
        }

        for (i = 0; i < rowCountJavaScript; i++) {
            for (int j = 0; j < columnCountJavaScript; j++) {
                Paragraph p;
                p = tableJavaScript.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText(tJavaScriptList.get(i).type);
                } else if (j == 2) {
                    p.appendText(tJavaScriptList.get(i).name);
                }
            }
        }
        subheading_JavaScript.applyStyle("myHeading_2");
        para_JavaScript.applyStyle("paraStyle");  
        }
        //END   MANIFEST
        
        //BEGIN SYSPREF
        if (bSystemPreferences) 
        {
        getSysPrefList(tParamsList);
        Paragraph subheading_env2 = section.addParagraph();
        subheading_env2.appendText("Встановлення системних налаштувань");

        Paragraph para_env = section.addParagraph();
        para_env.appendText("Увійдіть у клієнт Siebel (ENU) з правами адміністратора. "
                + "Перейдіть до екрану «Administration - Application», "
                + "а потім до виду «System Preferences».\n"
                + "***Дивись коментарі у стовбчику «Опис» щодо встановлення параметрів для продуктивного або тестового середовищя");

        Table tableSysPref = section.addTable(true);

        String[][] dataSysPref
                = {new String[]{"Системне налаштування", "Значення", "Опис"},};

        int rowCountSysPref = tSysPrefList.size();
        int columnCountSysPref = 3;
        tableSysPref.resetCells(rowCountSysPref+1, columnCountSysPref);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountSysPref; j++) {
            Paragraph p;
            p = tableSysPref.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataSysPref[i][j]);
        }

        for (i = 0; i < rowCountSysPref; i++) {
            for (int j = 0; j < columnCountSysPref; j++) {
                Paragraph p;
                p = tableSysPref.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText(tSysPrefList.get(i).name);
                } else if (j == 1) {
                    p.appendText(tSysPrefList.get(i).value);
                } else if (j == 2) {
                    p.appendText(tSysPrefList.get(i).comment);
                }
            }
        }
        subheading_env2.applyStyle("myHeading_2");
        para_env.applyStyle("paraStyle");
        }
        //END SYSPREF
        
        //BEGIN PICTURES
        if (bPicture)
        { 
        Paragraph subheading_pic = section.addParagraph();
        subheading_pic.appendText("Перенесення файлів зображень");
        subheading_pic.applyStyle("myHeading_1");
        
        CreateCatalog(sPathBase+"300-environment\\398-Bitmap File");
        
        Paragraph para_pic = section.addParagraph();
        para_pic.appendText("Скопіюйте із заміною файли з директорії "
                + sPathBase +"300-environment\\398-Bitmap File "
                + "до директорії "
                + "C:\\Siebel\\16.0.0.0.0\\eappweb\\public\\IMAGES "
                + "серверу siebelapp");
        para_pic.applyStyle("paraStyle");    
        }
        //END PICTURES
        
        
        //BEGIN SQL
        if (bSQL)
        {    
        getSQLList(tParamsList);    
        Paragraph subheading_sql = section.addParagraph();
        subheading_sql.appendText("Внесення змін до схем бази даних");
        subheading_sql.applyStyle("myHeading_1");

        Paragraph subheading_sql2 = section.addParagraph();
        subheading_sql2.appendText("Виконання скриптів оновлення бази даних");
        subheading_sql2.applyStyle("myHeading_2");
        
        Paragraph para_sql = section.addParagraph();
        para_sql.appendText("Під користувачем SIEBEL виконати скрипти в базі даних:");
        para_sql.applyStyle("paraStyle");
        
        CreateCatalog(sPathBase+"300-environment\\100-СУБД");
        
        Table tableSQL = section.addTable(true);
        String sPathSQL = sPathBase + "100-СУБД\\";
        String[][] dataSQL
                = {new String[]{"Тип об'єкту", "Назва об'єкту", "Скрипт"},};

        int rowCountSQL = tSQLList.size();
        int columnCountSQL = 3;
        tableSQL.resetCells(rowCountSQL+1, columnCountSQL);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountSQL; j++) {
            Paragraph p;
            p = tableSQL.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataSQL[i][j]);
        }

        for (i = 0; i < rowCountSQL; i++) {
            for (int j = 0; j < columnCountSQL; j++) {
                Paragraph p;
                p = tableSQL.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText(tSQLList.get(i).type);
                } else if (j == 1) {
                    p.appendText(tSQLList.get(i).name);
                } else if (j == 2) {
                    p.appendText(sPathSQL);
                }

            }
        }
    }
        //END SQL

        //BEGIN SRF
        if (bSRF)
        {
        Paragraph subheading_srf = section.addParagraph();
        subheading_srf.appendText("Заміна SRF-файлу (репозиторія)");

        Paragraph subheading_srf2 = section.addParagraph();
        subheading_srf2.appendText("Підготовчі роботи");

        Paragraph para_srf = section.addParagraph();
        para_srf.appendCheckBox();
        para_srf.appendText("Перед встановленням нових srf файлів на систему Siebel CRM необхідно виконати такі дії:\n"
                + "1.	Погодити час зупинки системи;\n"
                + "2.	Вимкнути користувачів від системи;\n"
                + "3.	Зупинити послуги системи Siebel CRM.\n"
                + "4.	Зберегти поточний файл SRF (зробити архівну копію)");

        Paragraph subheading_srf3 = section.addParagraph();
        subheading_srf3.appendText("Заміна SRF-файлу ");
        

        Paragraph para_srf3 = section.addParagraph();
        para_srf3.appendText(
                "Далі необхідно замінити SRF-файли на середовищі, де ведеться установка пакета. "
                + "На тестовому середовищі необхідно зупинити службу Siebel Server із заміною SRF. "
                + "На продуктивному середовищі необхідно по черзі виконати зупинку служби Siebel Server "
                + "із заміною SRF на всіх серверах Siebel (siebelbpm - 192.168.100.79, siebelapp - 192.168.100.80, vicisieb - 192.168.100.120).\n"
                + "Порядок заміни SRF-файлів на продуктивному середовищі наступний:\n"
                + "•	Зупинити службу Siebel Server на сервері vicisieb (192.168.100.120).\n"
                + "•	Замінити файли SRF.\n"
                + "Отриманий в результаті компіляції RUS файл siebel_sia.srf розмістити в папці C:\\Siebel\\16.0.0.0.0\\ses\\siebsrvr\\OBJECTS\\rus, замінивши наявний.\n"
                + "Отриманий в результаті компіляції ENU файл siebel_sia.srf розмістити в папці C:\\Siebel\\16.0.0.0.0\\ses\\siebsrvr\\OBJECTS\\enu, замінивши наявний.\n"
                + "•	Здійснити запуск служби Siebel Server, дочекатися закінчення запуску.\n"
                + "•	Виконати ту ж процедуру із сервером siebelapp (192.168.100.80) та з сервером siebelbpm (192.168.100.79).\n"
                + "Такий режим запуску дозволяє забезпечити безперервність доступності сервісу мобільного додатка з Інтернету.");
        subheading_srf.applyStyle("myHeading_1");
        subheading_srf2.applyStyle("myHeading_2");
        subheading_srf3.applyStyle("myHeading_2");
        para_srf.applyStyle("paraStyle");
        para_srf3.applyStyle("paraStyle");
        }
        //END SRF

        //BEGIN EAI
        if (bEAIDataMap)
        {
        getEAIDataMapList(tParamsList);
        Paragraph subheading_eai = section.addParagraph();
        subheading_eai.appendText("Перенос EAI DataMap");
        CreateCatalog(sPathBase+"300-environment\\318-EAI DataMap");

        Paragraph para_EAI = section.addParagraph();
        para_EAI.appendText("Увійдіть у клієнт Siebel (ENU) з правами адміністратора. "
                + "Перейдіть до екрану «Administration - Integration», а потім до «Data Maps». "
                + "Далі виконайте імпорт файлів (Menu – Import Data map)   з директорії "
                + sPathBase +"300-Оточення\\318-EAI DataMap.");

        Table tableEAI = section.addTable(true);
        String sPathEAIDataMap = sPathBase + "300-environment\\318-EAI DataMap";
        String[][] dataEAI
                = {new String[]{"Тип об'єкту", "Назва картки даних", "Файл"},};

        int rowCountEAI = tEAIDataMapList.size();
        int columnCountEAI = 3;
        tableEAI.resetCells(rowCountEAI+1, columnCountEAI);

        //fill the header to Table
        i = 0;
        for (int j = 0; j < columnCountEAI; j++) {
            Paragraph p;
            p = tableEAI.getRows().get(i).getCells().get(j).addParagraph();
            p.applyStyle("TableHeader");
            p.appendText(dataEAI[i][j]);
        }

        for (i = 0; i < rowCountEAI; i++) {
            for (int j = 0; j < columnCountEAI; j++) {
                Paragraph p;
                p = tableEAI.getRows().get(i+1).getCells().get(j).addParagraph();
                p.applyStyle("TableText");
                if (j == 0) {
                    p.appendText("318-EAI DataMap");
                } else if (j == 1) {
                    p.appendText(tEAIDataMapList.get(i).type);
                } else if (j == 2) {
                    p.appendText(sPathEAIDataMap);
                }

            }
        }
        DocPicture picture = para_EAI.appendPicture("src\\main\\resources\\DataMap1.png");
        picture.setWidth(400);
        picture.setHeight(300);
        picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        picture.addCaption("EAI DataMap", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);

        Paragraph para_EAI_end = section.addParagraph();
        para_EAI_end.appendText("Після імпорту потрібно почистити кеш: "
                + "перейдіть до екрану «Administration - Integration, "
                + "а потім до виду  «EAI Dispatcher Service View  (Rule Sets)» і натисніть «Clear Cache».");
        subheading_eai.applyStyle("myHeading_1");
        para_EAI.applyStyle("paraStyle");
        para_EAI_end.applyStyle("paraStyle");
        }
        //END EAI
        
        //BEGIN JOB
        if (bJob)
        {
            Paragraph subheading_job = section.addParagraph();
            subheading_job.appendText("Створення Job");
            
            Paragraph para_job = section.addParagraph();
            para_job.appendText(
                    "Відкрити в ENU “Administration – Server Management” – “Jobs”\n" +
                    "Створити Job\n" +
                    "Component | Job \"#######################\"\n" +
                    "Delete interval “XX hour”\n" +
                    "Repeat Unit=”Minutes”\n" +
                    "Repeat Interval=”XX”\n" +
                    "Repeating? \"XXX\"\n" +
                    "Додати параметри (Job parameters)\n" +
                    "Workflow Process Name = XXXXXXXXXXXXXXXXXXXXX\n" +
                    "Після цього натиснути \"Submit Job\"");   
            
            DocPicture picture = para_job.appendPicture("src\\main\\resources\\Job1.png");
            picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
            picture = para_job.appendPicture("src\\main\\resources\\Job2.png");
            picture.setWidth(400);
            picture.setHeight(300);
            picture.setHorizontalAlignment(ShapeHorizontalAlignment.Center);
            picture.addCaption("Створення Job", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);
            subheading_job.applyStyle("myHeading_1");
            para_job.applyStyle("paraStyle");            
            
        }
        //END JOB
        

        Paragraph subheading_end = section.addParagraph();
        subheading_end.appendText("Перевірка працездатності");

        Paragraph para_end = section.addParagraph();
        para_end.appendText("Після встановлення пакета, необхідно переконатися, що основний функціонал системи, "
                + "такий як: відкриття стартової сторінки, вхід до системи, вихід із системи, "
                + "відкриття основних видів, друк звітів і т.д. перебуває у робочому стані. "
                + "Для цього перейдіть за посиланням http://siebtapp/fins_rus для тестового середовища TEST, "
                + "(http://siebelapp/fins_rus - для продуктивного середовища PROD) "
                + "та виконайте необхідні перевірки функціоналу, згідно Release notes");

        // Apply built-in style to heading and subheadings
        // so that it is easily distinguishable
        // Apply the style to other paragraphs
        heading.applyStyle("nameStyle");
        subheading_common.applyStyle("myHeading_1");



        subheading_opys.applyStyle("myHeading_1");
        subheading_table.applyStyle("myHeading_1");

        subheading_goal.applyStyle("myHeading_1");
        subheading_end.applyStyle("myHeading_1");
        subheading_env.applyStyle("myHeading_1");

        para_object.applyStyle("paraStyle");


        para_goal_text.applyStyle("paraStyle");
        para_z.applyStyle("paraStyle");
        para_end.applyStyle("paraStyle");



        // Iteration for white spaces
        for (i = 0; i < section.getParagraphs().getCount(); i++) {

            // Automatically add whitespaces
            // to every paragraph in the file
            section.getParagraphs()
                    .get(i)
                    .getFormat()
                    .setAfterAutoSpacing(true);
        }

        
    for (i = 0; i < section.getParagraphs().getCount(); i++) {

            // Automatically add whitespaces
            // to every paragraph in the file
            section.setTextDirection(TextDirection.Left_To_Right);
        }
//                 //get the first table from the first section of the document
//                Table table1 = document.getSections().get(0).getTables().get(0);
//
//                //add a picture to the specified table cell and set picture size
//                DocPicture picture = table1.getRows().get(1).getCells().get(2).getParagraphs().get(0).appendPicture("E:\\sources\\Java\\GenerateSiebelPackageInstruction\\src\\main\\resources\\footer.jpg");
//                picture.setWidth(100);
//                picture.setHeight(100);
        // Save the document
        document.updateTableOfContents();
        document.saveToFile(
                FILE_OUT_NAME+".docx",
                FileFormat.Docx);
    }

    static private void ReadExcelFile() {
        try {

            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME+".xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            iterator.next();

            String sError = null;
            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
//                Iterator<Cell> cellIterator = currentRow.iterator();
                String categoryOfObject = currentRow.getCell(0).toString();
                String typeOfObject = currentRow.getCell(1).toString();
                String nameOfObject = currentRow.getCell(2).toString();
                tParamsList.add(new TParams(categoryOfObject, typeOfObject, nameOfObject));
                System.out.print(categoryOfObject + " " + typeOfObject + " " + nameOfObject);
                System.out.println();
                if (true) {//categoryOfObject.equalsIgnoreCase("200-Репозиторій")
                    switch (typeOfObject) {
                        case ("201-Table"):
                            bTable = true;
                            nT++;
                            break;
                        case ("202-Business Component"):
                            nR++;
                            break;
                        case ("203-Business Object"):
                            nR++;
                            break;
                        case ("205-Link"):
                            nR++;
                            break;
                        case ("207-EIM Interface Table"):
                            nR++;
                            break;
                        case ("211-Applet"):
                            nR++;
                            break;
                        case ("212-Pick List"):
                            nR++;
                            break;
                        case ("213-View"):
                            nR++;
                            break;
                        case ("214-Screen"):
                            nR++;
                            break;
                        case ("215-Task UI"):
                            nR++;
                            bTasks = true;
                            break;
                        case ("216-Symbolic String"):
                            nR++;
                            break;
                        case ("217-Application"):
                            nR++;
                            break;
                        case ("221-Workflow Process"):
                            nR++;
                            bWorkflowProcess = true;
                            break;
                        case ("234-Bitmap Category"):
                            break;
                        case ("233-Integration Object"):
                            bIntegrationObject = true;
                            nR++;
                            break;
                        case ("235-Icon Map"):
                            bPicture = true;
                            break;
                        case ("300-Other"):
                            sError = sError + " объект типа " + categoryOfObject + " " + nameOfObject + " не найден. ";
                            break;
                        case ("301-System Preferences"):
                            bSystemPreferences = true;
                            break;
                        case ("302-List Of Values"):
                            bLOV = true;
                            break;
                        case ("303-State Model"):
                            bStateModel = true;
                            break;
                        case ("308-DataMap"):
                            bDataMap = true;
                            break;
                        case ("313-Siebel job"):
                            bJob = true;
                            break;
                        case ("314-EIM Config File"):
                            bEIMConfigFile = true;
                            break;
                        case ("318-EAI DataMap"):
                            bEAIDataMap = true;
                            break;
                        case ("323-Comm Package"):
                            bCommPackage = true;
                            break;
                        case ("324-Profile Configuration"):
                            bProfileConfiguration = true;
                            break;
                        case ("325-Component Definitions"):
                            bComponentDefinitions = true;
                            break;
                        case ("342-Business Role"):
                            bBusinessRole = true;
                            break;
                        case ("363-Manifest File"):
                            bManifestFile = true;
                            break;
                        case ("364-Manifest Object"):
                            bManifestObject = true;
                            break;
                        case ("395-Java Script File"):
                            bJavaScriptFile = true;
                            break;
                        case ("398-Bitmap File"):
                            bPicture = true;
                            break;
                        case ("112-View"):
                            bSQL = true;
                            break;
                        case ("112-Вид"):
                            bSQL = true;
                            break;
                        case ("131-Package"):
                            bSQL = true;
                            break;
                        case ("151-Alter script"):
                            bSQL = true;
                            break;
                        case ("351-Saved Queries"):
                            bSavedQueries = true;
                            break;    
                        default:
                            sError = sError + " объект типа " + categoryOfObject + " " + nameOfObject + " не найден. ";
                            break;
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
//        Arrays.sort(tParamsList, new SortByCost());      
    }

    public static class TParams {

        private String category;
        private String type;
        private String name;

        public TParams(String category, String type, String name) {
            this.category = category;
            this.type = type;
            this.name = name;
        }
        // Getters + Setters
    }

    public static class TWorkflows {

        private Integer version;
        private String name;

        public TWorkflows(String type, Integer version) {
            this.version = version;
            this.name = name;
        }

        // Getters + Setters
        public Integer getVersion() {
            return version;
        }

        public void setVersion(Integer version) {
            this.version = version;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }

    public static class TTasks {

        private Integer version;
        private String name;

        public TTasks(String type, Integer version) {
            this.version = version;
            this.name = name;
        }

        // Getters + Setters
        public Integer getVersion() {
            return version;
        }

        public void setVersion(Integer version) {
            this.version = version;
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }

    public static class TEAIDataMap {
        private String type;
        public TEAIDataMap(String type) {
            this.type = type;
        }
    }

    public static class TSysPref {
        private String value;
        private String name;
        private String comment;
        public TSysPref(String name, String value, String comment) {
            this.comment = comment;
            this.name = name;
            this.value = value;
        }  
        public TSysPref(String name) {
            this.name = name;
        }
    }
    
    
    public static class TSQL {
        private String type;
        private String name;
        private String script;
        public TSQL(String type, String name, String script) {
            this.type = type;
            this.name = name;
            this.script = script;
        }
    }
    
    public static class TPDQ {
        private String type;
        private String name;
        private String script;
        public TPDQ(String type, String name, String script) {
            this.type = type;
            this.name = name;
            this.script = script;
        }
    }
    
    public static class TJavaScript {
        private String type;
        private String name;
        private String script;
        public TJavaScript(String type, String name, String script) {
            this.type = type;
            this.name = name;
            this.script = script;
        }
    }    
    
    public static class TCommPkg {
        private String type;
        private String name;
        private String file;
        public TCommPkg(String type, String name, String file) {
            this.type = type;
            this.name = name;
            this.file = file;
        }
    }    
    
    public static class TEIM {
        private String type;
        private String name;
        public TEIM(String type, String name) {
            this.type = type;
            this.name = name;
        }
    }
    
        public static class TStateModel{
        private String file;
        private String name;
        public TStateModel(String name, String file) {
            this.file = file;
            this.name = name;
        }
    }
        
    public static class TLOV {
        private String file;
        private String name;
        public TLOV(String name, String file) {
            this.file = file;
            this.name = name;
        }
    }

    public static class SortByCost implements Comparator<TParams> {
        public int compare(TParams a, TParams b) {
            if (a.type.compareTo(b.type) > 0) {
                return -1;
            } else if (a.type == b.type) {
                return 0;
            } else {
                return 1;
            }
        }
    }

    static private void getWorkflowList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("221-Workflow Process")) {
                System.out.println("Workflow is found " + currentElement.name);
                TWorkflows WF = new TWorkflows(currentElement.name, -1);
                WF.name = currentElement.name;
                WF.version = -1;
                tWFList.add(WF);
            }
        }
    }
    
    static private void getPDQList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("351-Saved Queries")) {
                System.out.println("PDQ is found " + currentElement.name);
                TPDQ PDQ = new TPDQ(currentElement.name, "", "");
                PDQ.name = currentElement.name;
                PDQ.type = "351-Saved Queries";
                PDQ.script="";
                tPDQList.add(PDQ);
            }
        }
    }
    
    static private void getCommPkgList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("323-Comm Package")) {
                System.out.println("Workflow is found " + currentElement.name);
                TCommPkg CommPkg = new TCommPkg("323-Comm Package", currentElement.name, "");
                CommPkg.name = currentElement.name;
                CommPkg.type = "323-Comm Package";
                CommPkg.file = "";
                tCommPkgList.add(CommPkg);
            }
        }
    }
    
    static private void UpdateWorkflowVersionList() {
        Iterator<TWorkflows> iterator = tWFList.iterator();
        while (iterator.hasNext()) {
            TWorkflows currentElement = iterator.next();
            try {
                currentElement.version = GetWorkflowVersion(currentElement.name);
            } catch (SiebelException ex) {
                Logger.getLogger(GenerateSiebelPackageInstruction.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
    }
    
    static private void UpdateTaskVersionList() {
        Iterator<TTasks> iterator = tTasksList.iterator();
        while (iterator.hasNext()) {
            TTasks currentElement = iterator.next();
            try {
                currentElement.version = GetTaskVersion(currentElement.name);
            } catch (SiebelException ex) {
                Logger.getLogger(GenerateSiebelPackageInstruction.class.getName()).log(Level.SEVERE, null, ex);
            }
            }
    }

    static private void getTaskList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("215-Task UI")) {
                System.out.println("Task is found " + currentElement.name);
                TTasks Task = new TTasks(currentElement.name, -1);
                Task.name = currentElement.name;
                Task.version = -1;
                tTasksList.add(Task);
            }
        }
    }
    
    static private void getStateModelList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("303-State Model")) {
                System.out.println("State Model is found " + currentElement.name);
                TStateModel StateModel = new TStateModel(currentElement.name, "");
                StateModel.name = currentElement.name;
                StateModel.file = "";
                tStateModelList.add(StateModel);
            }
        }
    }    

    static private void getTablesList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("201-Table")) {
                System.out.println("Table is found " + currentElement.name);
                tTablesList.add(currentElement.name);
            }
        }
    }
    
        static private void getIOList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("233-Integration Object")) {
                System.out.println("IO is found " + currentElement.name);
                tIOList.add(currentElement.name);
            }
        }
    }

    static private void getSysPrefList(ArrayList<TParams> tParamsList) throws SiebelException {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("301-System Preferences")) {
                System.out.println("SysPref is found " + currentElement.name);
                TSysPref SysPref = new TSysPref(currentElement.name, "", "");
                SysPref=GetSysPref(currentElement.name);
                tSysPrefList.add(SysPref);
            }
        }
    }
    


    static private void getSQLList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.category.equals("100-СУБД")) {
                System.out.println("SQL is found " + currentElement.name);
                TSQL SQL = new TSQL(currentElement.type, currentElement.name, "");
                SQL.name = currentElement.name;
                SQL.type = currentElement.type;
                tSQLList.add(SQL);
            }
        }
    }
    
    static private void getJavaScriptList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("395-Java Script File")) {
                System.out.println("JavaScript is found " + currentElement.name);
                TJavaScript JavaScript = new TJavaScript(currentElement.type, currentElement.name, "");
                JavaScript.name = currentElement.name;
                JavaScript.type = "395-Java Script File";
                tJavaScriptList.add(JavaScript);
            }
        }
    }  
    
    static private void getEIMList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("314-EIM Config File")) {
                System.out.println("EIM is found " + currentElement.name);
                TEIM EIM = new TEIM(currentElement.type, currentElement.name);
                EIM.name = currentElement.name;
                EIM.type = "314-EIM Config File";
                tEIMList.add(EIM);
            }
        }
    }
    
    static private void getLOVList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("302-List Of Values")) {
                System.out.println("LOV is found " + currentElement.name);
                TLOV LOV = new TLOV(currentElement.type, currentElement.name);
                LOV.name = currentElement.name;
                LOV.file = "LOV_"+currentElement.name+".xml";
                tLOVList.add(LOV);
            }
        }
    }    
    
        static private void getEAIDataMapList(ArrayList<TParams> tParamsList) {
        Iterator<TParams> iterator = tParamsList.iterator();
        while (iterator.hasNext()) {
            TParams currentElement = iterator.next();
            if (currentElement.type.equals("318-EAI DataMap") || currentElement.type.equals("308-DataMap")) {
                System.out.println("EAI DataMap is found " + currentElement.name);
                TEAIDataMap tEAI = new TEAIDataMap(currentElement.name);
                tEAI.type = currentElement.name;
                tEAIDataMapList.add(tEAI);
            }
        }
    }

    private static void insertHeaderAndFooter(Section section) throws Exception {
        String headerImage = "E:\\sources\\Java\\GenerateSiebelPackageInstruction\\src\\main\\resources\\header.jpg";
        section.getPageSetup().setPageSize(PageSize.A4);
        section.getPageSetup().getMargins().setTop(90f);
        section.getPageSetup().getMargins().setBottom(60f);
        section.getPageSetup().getMargins().setLeft(50f);
        section.getPageSetup().getMargins().setRight(30f);

        HeaderFooter header = section.getHeadersFooters().getHeader();
        HeaderFooter footer = section.getHeadersFooters().getFooter();

        //insert picture and text to header
        Paragraph headerParagraph = header.addParagraph();
        DocPicture headerPicture = headerParagraph.appendPicture(headerImage);
        headerPicture.setAllowOverlap(true);
        headerPicture.setTextWrappingStyle(TextWrappingStyle.Through);
        headerPicture.setTextWrappingType(TextWrappingType.Both);

        //header text
        TextRange text = headerParagraph.appendText("ТОВ \"Ареон Консалтінг\" \n"
                + "Україна, 04210, Київ, вул. Маршала Тимошенка, 21/14\n"
                + " Телефон: +38 (044) 538-08-00");
        text.getCharacterFormat().setFontName("Calibri");
        text.getCharacterFormat().setFontSize(11);
        headerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

        //border
        headerParagraph.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Single);
        headerParagraph.getFormat().getBorders().getBottom().setSpace(0.05F);

        //header picture layout - text wrapping
        headerPicture.setTextWrappingStyle(TextWrappingStyle.Behind);

        //header picture layout - position
        headerPicture.setHorizontalOrigin(HorizontalOrigin.Column);
        headerPicture.setHorizontalAlignment(ShapeHorizontalAlignment.Left);
//        headerPicture.setVerticalOrigin(VerticalOrigin.Top_Margin_Area);
        headerPicture.setVerticalAlignment(ShapeVerticalAlignment.Top);

        //insert picture to footer
        Paragraph footerParagraph = footer.addParagraph();

        //insert page number
        footerParagraph.appendField("Версія "+ pkgNo, FieldType.Field_Info);
        footerParagraph.appendText(" Стр.");
        footerParagraph.appendField("Страница", FieldType.Field_Page);
        footerParagraph.appendText(" з ");
        footerParagraph.appendField("ЧислоСтраниц", FieldType.Field_Num_Pages);
        footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
        footerParagraph.applyStyle("paraStyle");

        //border
        footerParagraph.getFormat().getBorders().getTop().setBorderType(BorderStyle.Single);
        footerParagraph.getFormat().getBorders().getTop().setSpace(0.05F);

    }
    
    static int GetWorkflowVersion(String NameWF) throws   SiebelException
 {
     int n=0;
     int tasks=0;
     		try {
			SiebelDataBean sblConnect = new SiebelDataBean();
			sblConnect.login(SiebelConnectString, SiebelUser, SiebelUserPassword, "enu");
 
			SiebelBusObject BO = sblConnect.getBusObject("PSP Procedure Loader");
			SiebelBusComp BC = BO.getBusComp("Repository Workflow Process");
                        BC.activateField("Version");
                        BC.activateField("Name");
                        BC.activateField("Status");
			BC.clearToQuery();
                        BC.setSearchSpec("Status", "COMPLETED");
                        BC.setSearchSpec("Process Name", "'"+NameWF+"'");

			BC.executeQuery(true);
                        
			if(BC.firstRecord())
			{
                            do {
                                    System.out.println(BC.getFieldValue("Name") +" "+BC.getFieldValue("Version") + " " + BC.getFieldValue("Status"));
                                    n=Integer.parseInt(BC.getFieldValue("Version"));
                            }while(BC.nextRecord());
                        }
                        System.out.println("n:"+n);
                        
			BC = null;
			BO = null;
			sblConnect.logoff();
 
		}
		catch (SiebelException e)
		{           
			e.printStackTrace();
		}
    return n ;           
 }   
    
    static int GetTaskVersion(String NameTask) throws   SiebelException
 {
     int n=0;
     int tasks=0;
     		try {
			SiebelDataBean sblConnect = new SiebelDataBean();
			sblConnect.login(SiebelConnectString, SiebelUser, SiebelUserPassword, "enu");
 
			SiebelBusObject BO = sblConnect.getBusObject("Repository Task");
			SiebelBusComp BC = BO.getBusComp("Repository Task");
                        BC.activateField("Version");
                        BC.activateField("Name");
                        BC.activateField("Status");
			BC.clearToQuery();
                        BC.setSearchSpec("Status", "COMPLETED");
                        BC.setSearchSpec("Task Name", NameTask);

			BC.executeQuery(true);
                        
			if(BC.firstRecord())
			{
                            do {
                                    System.out.println(BC.getFieldValue("Name") +BC.getFieldValue("Version") + " " + BC.getFieldValue("Status"));
                                    n=Integer.parseInt(BC.getFieldValue("Version"));
                            }while(BC.nextRecord());
                        }
                        System.out.println("n:"+n);
                        
			BC = null;
			BO = null;
			sblConnect.logoff();
 
		}
		catch (SiebelException e)
		{           
			e.printStackTrace();
		}
    return n ;           
 }  

static TSysPref GetSysPref(String Name) throws   SiebelException
 {
     TSysPref SysPref= new TSysPref(Name);
     		try {
			SiebelDataBean sblConnect = new SiebelDataBean();
			sblConnect.login(SiebelConnectString, SiebelUser, SiebelUserPassword, "enu");
 
			SiebelBusObject BO = sblConnect.getBusObject("System Preferences");
			SiebelBusComp BC = BO.getBusComp("System Preferences");
                        BC.activateField("Name");
                        BC.activateField("Value");
                        BC.activateField("Comments");
			BC.clearToQuery();
                        BC.setSearchSpec("Name", Name);
//                        BC.SetSearchExpr("[Name] = '" + Name + "'");

			BC.executeQuery(true);
                        
			if(BC.firstRecord())
			{
                            do {
                                    System.out.println(BC.getFieldValue("Name") +BC.getFieldValue("Value") + " " + BC.getFieldValue("Comments"));
                                    SysPref.comment=BC.getFieldValue("Comments");
                                    SysPref.name=BC.getFieldValue("Name");
                                    SysPref.value=BC.getFieldValue("Value");
                            }while(BC.nextRecord());
                        }
                        
			BC = null;
			BO = null;
			sblConnect.logoff();
 
		}
		catch (SiebelException e)
		{           
			e.printStackTrace();
		}
    return SysPref ;           
 }       


public static void CreateCatalog (String sPath)
{
              try {

            Path path = Paths.get(sPath);

            //java.nio.file.Files;
            Files.createDirectories(path);

            System.out.println("Directory is created!");

          } catch (IOException e) {

            System.err.println("Failed to create directory!" + e.getMessage());

          }
}

static void getProperties ()
{
        try (InputStream input = GenerateSiebelPackageInstruction.class.getClassLoader().getResourceAsStream("config.properties")) {
            Properties prop = new Properties();
            if (input == null) {
                System.out.println("Sorry, unable to find config.properties");
                return;
            }
            //load a properties file from class path, inside static method
            prop.load(input);
            //get the property value and print it out
            SiebelConnectString = prop.getProperty("siebel.url");
            SiebelUserPassword=prop.getProperty("siebel.password");
            SiebelUser=prop.getProperty("siebel.user");    
            System.out.println(SiebelConnectString);
            System.out.println(SiebelUser);
            System.out.println(SiebelUserPassword);

        } catch (IOException ex) {
            ex.printStackTrace();
        }
}
}
