/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */
package Logica;

import Clases.Documento;
//import Ventanas.Informacion;
import io.github.bonigarcia.wdm.WebDriverManager;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.Properties;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.Session;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.JOptionPane;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.apache.commons.lang.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.xml.sax.InputSource;

/**
 *
 * @author sopor
 */
public class Logica {

//    public static Informacion info;
    public static void descargarBoletas(String enviadosRecibidos, String fecha, String fecha1, int opcion, String tipo, ArrayList<Documento> arrDocumento, ArrayList<String> arrRuts, int paso, int num2) throws InterruptedException, IOException {
        String rut = "76008058-6";
        String password = "8058";

        String baseURL = "http://enteldte.facturanet.cl/";

        ChromeOptions options = new ChromeOptions();
        HashMap<String, Object> prefs = new HashMap<>();
        prefs.put("profile.default_content_settings.popups", 0);
        prefs.put("download.default_directory", System.getProperty("user.dir") + "\\XMLs");
        prefs.put("download.prompt_for_download", false);
        prefs.put("safebrowsing.enabled", true);
        prefs.put("disable-popup-blocking", true);
        prefs.put("download.extensions_to_open", "application/xml");

        options.setExperimentalOption("prefs", prefs);
        options.addArguments("start-maximized");
        options.addArguments("--host-resolver-rules=MAP www.google-analytics.com 127.0.0.1");
        options.addArguments("--safebrowsing-disable-download-protection");
        options.addArguments("safebrowsing-disable-extension-blacklist");
        options.addArguments("--headless=new");
        options.addArguments("--log-level=3");
        options.addArguments("--silent");
        options.addArguments("--remote-allow-origins=*");

        WebDriverManager.chromedriver().setup();

        WebDriver driver = new ChromeDriver(options);
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, 10);

        driver.get(baseURL);

        wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"username\"]")));

        driver.findElement(By.xpath("//*[@id=\"username\"]")).sendKeys(rut);
        driver.findElement(By.xpath("//*[@id=\"password\"]")).sendKeys(password);
        driver.findElement(By.xpath("//*[@id=\"content\"]/div/div[2]/form/div[3]/button")).click();

        System.out.println("fecha " + fecha);
        System.out.println("fecha1 " + fecha1);
        String[] split = fecha.split("-");
        String[] splitx = fecha1.split("-");

        LocalDate d1 = LocalDate.parse(split[2] + "-" + split[1] + "-" + split[0], DateTimeFormatter.ISO_LOCAL_DATE);
        LocalDate d2 = LocalDate.parse(splitx[2] + "-" + splitx[1] + "-" + splitx[0], DateTimeFormatter.ISO_LOCAL_DATE);

        Duration between = Duration.between(d1.atStartOfDay(), d2.atStartOfDay());
        long toDays = between.toDays();

        System.out.println("toDays " + toDays);
        ZoneId defaultZoneId = ZoneId.systemDefault();

        for (int h = 0; h <= toDays; h++) {
            Date date = Date.from(d1.atStartOfDay(defaultZoneId).toInstant());

            DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
            String strDatex = dateFormat.format(date);

            System.out.println("strDate " + strDatex);
            String ruta = "";
            if (enviadosRecibidos.equals("ENVIADOS")) {
                System.out.println("ENVIADOS");
                ruta = "http://enteldte.facturanet.cl/documento/buscar/index.php";
            } else {
                System.out.println("RECIBIDOS");
                ruta = "http://enteldte.facturanet.cl/documento_recibido/buscar/index.php";
            }

            driver.get(ruta);
            Select select;
            if (enviadosRecibidos.equals("ENVIADOS")) {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("mantenedor_form_M_documento_tido_id")));
            } else {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("mantenedor_form_M_documento_recibido_tido_id")));
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                select = new Select(driver.findElement(By.id("mantenedor_form_M_documento_tido_id")));
                select.selectByIndex(opcion);
            } else {
                select = new Select(driver.findElement(By.id("mantenedor_form_M_documento_recibido_tido_id")));
                select.selectByIndex(opcion);
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_emision__desde\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_emision__desde\"]")).clear();
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_emision__desde\"]")).sendKeys(strDatex);
            } else {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_emision__desde\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_emision__desde\"]")).clear();
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_emision__desde\"]")).sendKeys(strDatex);
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_emision__hasta\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_emision__hasta\"]")).clear();
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_emision__hasta\"]")).sendKeys(strDatex);
            } else {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_emision__hasta\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_emision__hasta\"]")).clear();
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_emision__hasta\"]")).sendKeys(strDatex);
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_recepcion__desde\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_recepcion__desde\"]")).clear();
            } else {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_recepcion__desde\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_recepcion__desde\"]")).clear();
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_recepcion__hasta\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_fecha_recepcion__hasta\"]")).clear();
            } else {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_recepcion__hasta\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_fecha_recepcion__hasta\"]")).clear();
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_folio__hasta\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_docu_folio__hasta\"]")).clear();
            } else {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_folio__desde\"]")));
                driver.findElement(By.xpath("//*[@id=\"mantenedor_form_M_documento_recibido_docr_folio__desde\"]")).clear();
            }
            if (enviadosRecibidos.equals("ENVIADOS")) {
                select = new Select(driver.findElement(By.id("mantenedor_form_M_documento__limit")));
                select.selectByIndex(3);
            } else {
                select = new Select(driver.findElement(By.id("mantenedor_form_M_documento_recibido__limit")));
                select.selectByIndex(3);
            }

            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"div_btn_2\"]/span/button")));
            WebElement element = driver.findElement(By.xpath("//*[@id=\"div_btn_2\"]/span/button"));
            JavascriptExecutor executor = (JavascriptExecutor) driver;
            executor.executeScript("arguments[0].scrollIntoView(true);", element);
            element.click();

            System.out.println("Aqui");
            String cantidad = "";
            try {
                wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//*[@id=\"div_btn_imprimir_todo\"]/span/button")));
                cantidad = driver.findElement(By.xpath("//*[@id=\"div_btn_imprimir_todo\"]/span/button")).getText();
                cantidad = cantidad.replace("Descargar", "");
                cantidad = cantidad.replace("los", "");
                cantidad = cantidad.replace("documentos", "");
                cantidad = cantidad.replace("el", "");
                cantidad = cantidad.replace("documento", "");
                cantidad = cantidad.trim();
            } catch (Exception ex) {
                System.out.println("ex " + ex);
                cantidad = "0";
            }

            if (cantidad.equals("")) {
                cantidad = "1";
            }

            wait = new WebDriverWait(driver, 5);

            int x = 1;
            int i = 0;
            int cont = 0;
            boolean bool = true;

            do {
                System.out.println("------------------------------------------> " + x);
                for (i = 0; i < 20; i++) {
                    try {
                        wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"yui-dt0-bdrow" + i + "-cell0\"]/div")));
                        String text = driver.findElement(By.xpath("//*[@id=\"yui-dt0-bdrow" + i + "-cell0\"]/div")).getText();
                        System.out.println("------> " + text);

                        String attribute = "";
                        if (enviadosRecibidos.equals("ENVIADOS")) {
                            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"yui-dt0-bdrow" + i + "-cell13\"]/div/a[1]")));
                            attribute = driver.findElement(By.xpath("//*[@id=\"yui-dt0-bdrow" + i + "-cell13\"]/div/a[1]")).getAttribute("href");
                        } else {
                            //*[@id="yui-dt0-bdrow0-cell8"]/div/a[1]
                            wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"yui-dt0-bdrow" + i + "-cell8\"]/div/a[1]")));
                            attribute = driver.findElement(By.xpath("//*[@id=\"yui-dt0-bdrow" + i + "-cell8\"]/div/a[1]")).getAttribute("href");
                        }

//                System.out.println("attribute " + attribute);
                        ((JavascriptExecutor) driver).executeScript("window.open()");
                        ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
                        driver.switchTo().window(tabs.get(1));
                        driver.get(attribute);

                        String text1 = driver.findElement(By.tagName("body")).getText();
                        text1 = text1.replace("This XML file does not appear to have any style information associated with it. The document tree is shown below.", "").trim();
                        text1 = text1.replace("&", "Y");

                        Document convertStringToDocument = convertStringToDocument(text1);

                        if (convertStringToDocument != null) {
                            convertStringToDocument.getDocumentElement().normalize();

                            Documento documento = new Documento();

                            if (enviadosRecibidos.equals("ENVIADOS")) {
                                documento.setColumna0("1-1-40-01");
                            } else {
                                documento.setColumna0("2-1-20-01");
                            }

                            org.w3c.dom.NodeList nList = convertStringToDocument.getElementsByTagName("IdDoc");

                            String folio = "";

                            for (int temp = 0; temp < nList.getLength(); temp++) {
                                Node node = nList.item(temp);
                                System.out.println("");    //Just a separator
                                if (node.getNodeType() == Node.ELEMENT_NODE) {
                                    //Print each employee's detail
                                    Element eElement = (Element) node;
//                                System.out.println("Folio : " + eElement.getElementsByTagName("Folio").item(0).getTextContent());
                                    folio = eElement.getElementsByTagName("Folio").item(0).getTextContent();
                                    documento.setColumna10(eElement.getElementsByTagName("Folio").item(0).getTextContent());
                                    documento.setColumna14(eElement.getElementsByTagName("Folio").item(0).getTextContent());

                                    documento.setColumna11(strDatex.replace("-", "/"));
                                    documento.setColumna12(strDatex.replace("-", "/"));
                                }
                            }

                            nList = convertStringToDocument.getElementsByTagName("Totales");

                            String neto = "";
                            String iva = "";
                            String exento = "0";

                            for (int temp = 0; temp < nList.getLength(); temp++) {
                                Node node = nList.item(temp);
                                System.out.println("");    //Just a separator
                                if (node.getNodeType() == Node.ELEMENT_NODE) {
                                    //Print each employee's detail
                                    Element eElement = (Element) node;
//                                System.out.println("MntNeto : " + eElement.getElementsByTagName("MntNeto").item(0).getTextContent());
                                    neto = eElement.getElementsByTagName("MntNeto").item(0).getTextContent();

//                                System.out.println("IVA : " + eElement.getElementsByTagName("IVA").item(0).getTextContent());
                                    iva = eElement.getElementsByTagName("IVA").item(0).getTextContent();

                                    try {
//                                System.out.println("IVA : " + eElement.getElementsByTagName("IVA").item(0).getTextContent());
                                        exento = eElement.getElementsByTagName("MntExe").item(0).getTextContent();
                                    } catch (Exception ex) {

                                    }
                                }
                            }
                            String rutReceptor = "";
                            String razonSocial = "";

                            if (opcion == 4 || opcion == 8 || opcion == 9) {
                                nList = convertStringToDocument.getElementsByTagName("Receptor");

                                for (int temp = 0; temp < nList.getLength(); temp++) {
                                    Node node = nList.item(temp);
                                    System.out.println("");    //Just a separator
                                    if (node.getNodeType() == Node.ELEMENT_NODE) {
                                        Element eElement = (Element) node;
//                                      System.out.println("MntNeto : " + eElement.getElementsByTagName("MntNeto").item(0).getTextContent());
                                        rutReceptor = eElement.getElementsByTagName("RUTRecep").item(0).getTextContent();

//                                        System.out.println("MntNeto : " + eElement.getElementsByTagName("MntNeto").item(0).getTextContent());
                                        razonSocial = eElement.getElementsByTagName("RznSocRecep").item(0).getTextContent();
                                    }
                                }
                            }

                            arrRuts.add(rutReceptor + "@" + razonSocial);

                            String TpoDocRef = "";

                            if (opcion == 8 || opcion == 9) {
                                nList = convertStringToDocument.getElementsByTagName("Referencia");

                                for (int temp = 0; temp < nList.getLength(); temp++) {
                                    Node node = nList.item(temp);
                                    System.out.println("");    //Just a separator
                                    if (node.getNodeType() == Node.ELEMENT_NODE) {
                                        //Print each employee's detail
                                        Element eElement = (Element) node;
//                                System.out.println("MntNeto : " + eElement.getElementsByTagName("MntNeto").item(0).getTextContent());
                                        TpoDocRef = eElement.getElementsByTagName("TpoDocRef").item(0).getTextContent();

                                        if (TpoDocRef.equals("33")) {
                                            TpoDocRef = "VF";
                                        } else if (TpoDocRef.equals("39")) {
                                            TpoDocRef = "VB";
                                        } else if (TpoDocRef.equals("56")) {
                                            TpoDocRef = "VD";
                                        } else if (TpoDocRef.equals("61")) {
                                            TpoDocRef = "VC";
                                        }

//                                System.out.println("IVA : " + eElement.getElementsByTagName("IVA").item(0).getTextContent());
                                        documento.setColumna10(folio);
                                        documento.setColumna14(eElement.getElementsByTagName("FolioRef").item(0).getTextContent());
                                    }
                                }
                            }

                            nList = convertStringToDocument.getElementsByTagName("Receptor");

                            String rutx = "";

                            for (int temp = 0; temp < nList.getLength(); temp++) {
                                Node node = nList.item(temp);
                                System.out.println("");    //Just a separator
                                if (node.getNodeType() == Node.ELEMENT_NODE) {
                                    //Print each employee's detail
                                    Element eElement = (Element) node;
//                                System.out.println("MntNeto : " + eElement.getElementsByTagName("MntNeto").item(0).getTextContent());
                                    rutx = eElement.getElementsByTagName("RUTRecep").item(0).getTextContent();
                                }
                            }

                            if (rutx.equals("")) {
                                rutx = "1";
                            } else {
                                String[] split2 = rutx.split("-");
                                rutx = split2[0].trim();
                            }
                            if (enviadosRecibidos.equals("ENVIADOS")) {
                                switch (opcion) {
                                    case 8:
                                        documento.setColumna2("" + (Integer.parseInt(neto) + Integer.parseInt(exento) + Integer.parseInt(iva)));
                                        documento.setColumna1("0");
                                        break;
                                    default:
                                        documento.setColumna1("" + (Integer.parseInt(neto) + Integer.parseInt(exento) + Integer.parseInt(iva)));
                                        documento.setColumna2("0");
                                        break;
                                }
                            } else {
                                switch (opcion) {
                                    case 8:
                                        documento.setColumna1("" + (Integer.parseInt(neto) + Integer.parseInt(iva)));
                                        documento.setColumna2("0");
                                        break;
                                    default:
                                        documento.setColumna2("" + (Integer.parseInt(neto) + Integer.parseInt(iva)));
                                        documento.setColumna1("0");
                                        break;
                                }
                            }

                            switch (opcion) {
                                case 1:
                                    documento.setColumna3("BOLETA");
                                    break;
                                case 4:
                                    documento.setColumna3("FACTURA");
                                    break;
                                case 8:
                                    documento.setColumna3("NOTACREDITO");
                                    break;
                                case 9:
                                    documento.setColumna3("NOTADEBITO");
                                    break;
                                default:
                                    break;
                            }
                            documento.setColumna4("1");
                            if (enviadosRecibidos.equals("ENVIADOS")) {
                                switch (opcion) {
                                    case 8:
                                        documento.setColumna6("" + (Integer.parseInt(neto) + Integer.parseInt(exento) + Integer.parseInt(iva)));
                                        documento.setColumna5("0");
                                        break;
                                    default:
                                        documento.setColumna5("" + (Integer.parseInt(neto) + Integer.parseInt(exento) + Integer.parseInt(iva)));
                                        documento.setColumna6("0");
                                        break;
                                }
                            } else {
                                switch (opcion) {
                                    case 8:
                                        documento.setColumna5("" + (Integer.parseInt(neto) + Integer.parseInt(iva)));
                                        documento.setColumna6("0");
                                        break;
                                    default:
                                        documento.setColumna6("" + (Integer.parseInt(neto) + Integer.parseInt(iva)));
                                        documento.setColumna5("0");
                                        break;
                                }
                            }
                            if (enviadosRecibidos.equals("ENVIADOS")) {
                                documento.setColumna7("300");
                            } else {
                                documento.setColumna7("");
                            }

                            switch (opcion) {
                                case 1:
                                    documento.setColumna3("BOLETA");
                                    break;
                                case 4:
                                    documento.setColumna3("FACTURA");
                                    break;
                                case 8:
                                    documento.setColumna3("NOTACREDITO");
                                    break;
                                case 9:
                                    documento.setColumna3("NOTADEBITO");
                                    break;
                                default:
                                    break;
                            }

                            switch (opcion) {
                                case 1:
                                    documento.setColumna8("1");
                                    break;
                                case 4:
                                    documento.setColumna8(rutx);
                                    break;
                                case 8:
                                    documento.setColumna8(rutx);
                                    break;
                                case 9:
                                    documento.setColumna8(rutx);
                                    break;
                                default:
                                    break;
                            }

                            if (enviadosRecibidos.equals("ENVIADOS")) {
                                switch (opcion) {
                                    case 1:
                                        documento.setColumna9("VB");
                                        documento.setColumna13("VB");
                                        break;
                                    case 4:
                                        documento.setColumna9("VF");
                                        documento.setColumna13("VF");
                                        break;
                                    case 8:
                                        documento.setColumna9("VC");
                                        documento.setColumna13(TpoDocRef);
                                        break;
                                    case 9:
                                        documento.setColumna9("VD");
                                        documento.setColumna13(TpoDocRef);
                                        break;
                                    default:
                                        break;
                                }
                            } else {
                                switch (opcion) {
                                    case 1:
                                        documento.setColumna9("CB");
                                        documento.setColumna13("CB");
                                        break;
                                    case 4:
                                        documento.setColumna9("CF");
                                        documento.setColumna13("CF");
                                        break;
                                    case 8:
                                        documento.setColumna9("CC");
                                        documento.setColumna13(TpoDocRef);
                                        break;
                                    case 9:
                                        documento.setColumna9("CD");
                                        documento.setColumna13(TpoDocRef);
                                        break;
                                    default:
                                        break;
                                }
                            }

                            documento.setColumna15(neto);
                            documento.setColumna16(exento);
                            documento.setColumna17(iva);
                            documento.setColumna18("" + (Integer.parseInt(neto) + Integer.parseInt(exento) + +Integer.parseInt(iva)));
                            documento.setColumna19("S");
                            documento.setColumna20("N");

                            System.out.println(documento.getColumna0() + " | "
                                    + documento.getColumna1() + " | "
                                    + documento.getColumna2() + " | "
                                    + documento.getColumna3() + " | "
                                    + documento.getColumna4() + " | "
                                    + documento.getColumna5() + " | "
                                    + documento.getColumna6() + " | "
                                    + documento.getColumna7() + " | "
                                    + documento.getColumna8() + " | "
                                    + documento.getColumna9() + " | "
                                    + documento.getColumna10() + " | "
                                    + documento.getColumna11() + " | "
                                    + documento.getColumna12() + " | "
                                    + documento.getColumna13() + " | "
                                    + documento.getColumna14() + " | "
                                    + documento.getColumna15() + " | "
                                    + documento.getColumna16() + " | "
                                    + documento.getColumna17() + " | "
                                    + documento.getColumna18() + " | "
                                    + documento.getColumna19() + " | "
                                    + documento.getColumna20() + " | ");
                            arrDocumento.add(documento);
                        }

                        driver.close();

                        driver.switchTo().window(tabs.get(0));

                        cont++;
                    } catch (Exception ex) {
                        Logger.getLogger(Logica.class.getName()).log(Level.SEVERE, null, ex);
                        bool = false;
                        break;
                    }
                }

                if (bool) {
                    try {
                        wait = new WebDriverWait(driver, 2);
                        wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#yui-dt-pagselect0 > a.yui-dt-next")));
                        driver.findElement(By.cssSelector("#yui-dt-pagselect0 > a.yui-dt-next")).click();
                    } catch (Exception ex) {
                        System.out.println("AHHHHH!");
                        bool = false;
                    }

                    wait = new WebDriverWait(driver, 5);

                    x++;
                    i = 0;
                }
            } while (bool);

            d1 = d1.plusDays(1);
        }

        if (paso == 1) {
            System.out.println("1");

            if (!arrDocumento.isEmpty()) {
                crearExcel(enviadosRecibidos, arrDocumento, opcion, tipo, num2);
                System.out.println("arrRuts " + arrRuts.size());

                if (opcion != 2) {
                    crearExcel2(enviadosRecibidos, arrRuts, opcion);
                }
            }
            System.out.println("2");
        }

        driver.quit();
    }

    public static void crearExcel2(String enviadosRecibidos, ArrayList<String> arrRuts, int opcion) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("Facturas");

        XSSFFont headerFont = workbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.index);
        CellStyle headerCellStyle = spreadsheet.getWorkbook().createCellStyle();
        // fill foreground color ...
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        // and solid fill pattern produces solid grey cell fill
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setFont(headerFont);

//        Informacion.jProgressBar1.setMaximum(arrRuts.size());
        System.out.println(arrRuts);

        Set<String> s = new LinkedHashSet<>(arrRuts);
        arrRuts.clear();
        arrRuts.addAll(s);

        System.out.println("arrRuts " + arrRuts.size());

        System.out.println(arrRuts);

        int filaInicio = 0;
        for (int f = 0; f < arrRuts.size(); f++) {
            System.out.println("f + 1 " + f + 1);

            String get = "";
            String razones = "";

            String getx = arrRuts.get(f);

            System.out.println("getx " + getx);

            String[] splitx = getx.split("@");

            get = splitx[0];
            razones = splitx[1];

            System.out.println("get " + get);
            System.out.println("razones " + razones);

            String[] split = get.split("-");
            String name = split[0];

            Row fila = spreadsheet.createRow(filaInicio);
            for (int c = 0; c < 4; c++) {
                switch (c) {
                    case 0: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(name);
                        break;
                    }
                    case 1: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(razones);
                        break;
                    }
                    case 2: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(razones);
                        break;
                    }
                    case 3: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(get);
                        break;
                    }
                }
            }
            filaInicio++;
        }

        String nombre = "";
        String ruta = "";
        String envidaRecivida = "";
        if (enviadosRecibidos.equals("ENVIADOS")) {
            switch (opcion) {
                case 2:
                    nombre = "BOLETA";
                    ruta = System.getProperty("user.dir") + "\\BOLETAENVIADA\\";
                    break;
                case 5:
                    nombre = "FACTURA";
                    ruta = System.getProperty("user.dir") + "\\FACTURAENVIADA\\";
                    break;
                case 8:
                    nombre = "NOTACREDITO";
                    ruta = System.getProperty("user.dir") + "\\NCENVIADA\\";
                    break;
                case 9:
                    nombre = "NOTADEBITO";
                    ruta = System.getProperty("user.dir") + "\\DNENVIADA\\";
                    break;
                default:
                    break;
            }
            envidaRecivida = "ENVIADOS";
        } else {
            switch (opcion) {
                case 2:
                    nombre = "BOLETA";
                    ruta = System.getProperty("user.dir") + "\\BOLETARECIBIDA\\";
                    break;
                case 5:
                    nombre = "FACTURA";
                    ruta = System.getProperty("user.dir") + "\\FACTURARECIBIDA\\";
                    break;
                case 8:
                    nombre = "NOTACREDITO";
                    ruta = System.getProperty("user.dir") + "\\NCRECIBIDA\\";
                    break;
                case 9:
                    nombre = "NOTADEBITO";
                    ruta = System.getProperty("user.dir") + "\\DNRECIBIDA\\";
                    break;
                default:
                    break;
            }
            envidaRecivida = "RECIBIDA";
        }
        try {
            File file = new File(ruta + "CLIENTE" + nombre + ".xlsx");
            FileOutputStream out = new FileOutputStream(file);
            workbook.write(out);
            out.close();

//            Logica.enviarCorreo("CLIENTE" + nombre + envidaRecivida + ".xlsx", ruta + "CLIENTE" + nombre + ".xlsx");
//            Desktop desktop = Desktop.getDesktop();
//            if (file.exists()) {
//                desktop.open(file);
//            }
        } catch (Exception ex) {
//            JOptionPane.showMessageDialog(null, "( ˘︹˘ ) Algo salio mal!", "Oh, no!", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void crearExcel(String enviadosRecibidos, ArrayList<Documento> arrDocumento, int opcion, String tipo, int num2) throws FileNotFoundException, IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet("Facturas");

        XSSFFont headerFont = workbook.createFont();
        headerFont.setColor(IndexedColors.WHITE.index);
        CellStyle headerCellStyle = spreadsheet.getWorkbook().createCellStyle();
        // fill foreground color ...
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
        // and solid fill pattern produces solid grey cell fill
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setFont(headerFont);

//        Informacion.jProgressBar1.setMaximum(arrDocumento.size());
        int sumaNeto = 0;
        int sumaIva = 0;
        int sumaExcento = 0;

        int filaInicio = 0;
        for (int f = 0; f < arrDocumento.size(); f++) {
//            Informacion.jProgressBar1.setValue(f + 1);

            Documento documento = arrDocumento.get(f);
            Row fila = spreadsheet.createRow(filaInicio);
            for (int c = 0; c < 45; c++) {
                switch (c) {
                    case 0: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("001");
                        break;
                    }
                    case 1: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna0());
                        break;
                    }
                    case 2: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna1()));
                        break;
                    }
                    case 3: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna2()));
                        break;
                    }
                    case 4: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna3());
                        break;
                    }
                    case 5: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna4()));
                        break;
                    }
                    case 6: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna5()));
                        break;
                    }
                    case 7: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna6()));
                        break;
                    }
                    case 8: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 9: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 10: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 11: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 12: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 13: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 14: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 15: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 16: {
                        Cell celda = fila.createCell(c);
                        try {
                            celda.setCellValue(Integer.parseInt(documento.getColumna7()));
                        } catch (Exception ex) {
                            celda.setCellValue("");
                        }
                        break;
                    }
                    case 17: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 18: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 19: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna8()));
                        break;
                    }
                    case 20: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna9());
                        break;
                    }
                    case 21: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna10()));
                        break;
                    }
                    case 22: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna11());
                        break;
                    }
                    case 23: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna12());
                        break;
                    }
                    case 24: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna13());
                        break;
                    }
                    case 25: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna14()));
                        break;
                    }
                    case 26: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 27: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna15()));
                        sumaNeto = sumaNeto + Integer.parseInt(documento.getColumna15());
                        break;
                    }
                    case 28: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna16()));
                        sumaExcento = sumaExcento + Integer.parseInt(documento.getColumna16());
                        break;
                    }
                    case 29: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna17()));
                        sumaIva = sumaIva + Integer.parseInt(documento.getColumna17());
                        break;
                    }
                    case 30: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 31: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 32: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 33: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 34: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 35: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 36: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(Integer.parseInt(documento.getColumna18()));
                        break;
                    }
                    case 37: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 38: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 39: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 40: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("");
                        break;
                    }
                    case 41: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna19());
                        break;
                    }
                    case 42: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue(documento.getColumna20());
                        break;
                    }
                    case 43: {
                        Cell celda = fila.createCell(c);
                        celda.setCellValue("T");
                        break;
                    }
                    case 44: {
                        if (filaInicio == 0) {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue(tipo);
                            break;
                        }
                    }
                    default:
                        break;
                }
            }
            filaInicio++;
        }

        if (enviadosRecibidos.equals("ENVIADOS")) {
            for (int f = 0; f < 2; f++) {
                Row fila = spreadsheet.createRow(filaInicio);
                for (int c = 0; c < 45; c++) {
                    switch (c) {
                        case 0: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("001");
                            break;
                        }
                        case 1: {
                            Cell celda = fila.createCell(c);
                            if (f == 0) {
                                celda.setCellValue("5-1-10-01");
                            } else {
                                celda.setCellValue("2-1-40-03");
                            }
                            break;
                        }
                        case 2: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                            }

                            break;
                        }
                        case 3: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                            }
                            break;
                        }
                        case 4: {
                            Cell celda = fila.createCell(c);
                            switch (opcion) {
                                case 1:
                                    celda.setCellValue("BOLETA");
                                    break;
                                case 4:
                                    celda.setCellValue("FACTURA");
                                    break;
                                case 8:
                                    celda.setCellValue("NOTACREDITO");
                                    break;
                                case 9:
                                    celda.setCellValue("NOTADEBITO");
                                    break;
                                default:
                                    break;
                            }
                            break;
                        }
                        case 5: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue(Integer.parseInt("1"));
                            break;
                        }
                        case 6: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                            }

                            break;
                        }
                        case 7: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                            }
                            break;
                        }
                        case 8: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 9: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 10: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 11: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 12: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 13: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 14: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 15: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("");
                            break;
                        }
                        case 16: {
                            Cell celda = fila.createCell(c);
                            if (f == 0) {
                                celda.setCellValue(Integer.parseInt("300"));
                            }
                            break;
                        }
                    }
                }
                filaInicio++;
            }
        } else {
            for (int f = 0; f < 3; f++) {
                Row fila = spreadsheet.createRow(filaInicio);
                for (int c = 0; c < 45; c++) {
                    switch (c) {
                        case 0: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("001");
                            break;
                        }
                        case 1: {
                            Cell celda = fila.createCell(c);
                            if (f == 0) {
                                celda.setCellValue("1-1-80-01");
                            } else if (f == 1) {
                                celda.setCellValue("1-1-80-01");
                            } else if (f == 2) {
                                celda.setCellValue("1-1-60-02");
                            }
                            break;
                        }
                        case 2: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else if (f == 1) {
                                        celda.setCellValue(sumaExcento);
                                    } else if (f == 2) {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                            }
                            break;
                        }
                        case 3: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else if (f == 1) {
                                        celda.setCellValue(sumaExcento);
                                    } else if (f == 2) {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                            }

                            break;
                        }
                        case 4: {
                            Cell celda = fila.createCell(c);
                            switch (opcion) {
                                case 1:
                                    celda.setCellValue("BOLETA");
                                    break;
                                case 4:
                                    celda.setCellValue("FACTURA");
                                    break;
                                case 8:
                                    celda.setCellValue("NOTACREDITO");
                                    break;
                                case 9:
                                    celda.setCellValue("NOTADEBITO");
                                    break;
                                default:
                                    break;
                            }
                            break;
                        }
                        case 5: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue(Integer.parseInt("1"));
                            break;
                        }
                        case 6: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else if (f == 1) {
                                        celda.setCellValue(sumaExcento);
                                    } else if (f == 2) {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                            }
                            break;
                        }
                        case 7: {
                            switch (opcion) {
                                case 8:
                                    Cell celda = fila.createCell(c);
                                    if (f == 0) {
                                        celda.setCellValue(sumaNeto);
                                    } else if (f == 1) {
                                        celda.setCellValue(sumaExcento);
                                    } else if (f == 2) {
                                        celda.setCellValue(sumaIva);
                                    }
                                    break;
                                default:
                                    celda = fila.createCell(c);
                                    celda.setCellValue(Integer.parseInt("0"));
                                    break;
                            }

                            break;
                        }
                        case 43: {
                            Cell celda = fila.createCell(c);
                            celda.setCellValue("T");
                            break;
                        }
                    }
                }
                filaInicio++;
            }
        }

        String nombre = "";
        String ruta = "";
        if (enviadosRecibidos.equals("ENVIADOS")) {
            switch (opcion) {
                case 2:
                    nombre = "BOLETA";
                    ruta = System.getProperty("user.dir") + "\\BOLETAENVIADA\\";
                    break;
                case 5:
                    nombre = "FACTURA";
                    ruta = System.getProperty("user.dir") + "\\FACTURAENVIADA\\";
                    break;
                case 8:
                    nombre = "NOTACREDITO";
                    ruta = System.getProperty("user.dir") + "\\NCENVIADA\\";
                    break;
                case 9:
                    nombre = "NOTADEBITO";
                    ruta = System.getProperty("user.dir") + "\\DNENVIADA\\";
                    break;
                default:
                    break;
            }
        } else {
            switch (opcion) {
                case 2:
                    nombre = "BOLETA";
                    ruta = System.getProperty("user.dir") + "\\BOLETARECIBIDA\\";
                    break;
                case 5:
                    nombre = "FACTURA";
                    ruta = System.getProperty("user.dir") + "\\FACTURARECIBIDA\\";
                    break;
                case 8:
                    nombre = "NOTACREDITO";
                    ruta = System.getProperty("user.dir") + "\\NCRECIBIDA\\";
                    break;
                case 9:
                    nombre = "NOTADEBITO";
                    ruta = System.getProperty("user.dir") + "\\DNRECIBIDA\\";
                    break;
                default:
                    break;
            }
        }

        if (num2 != 999) {
            try {
                File file = new File(ruta + nombre + num2 + ".xlsx");
                FileOutputStream out = new FileOutputStream(file);
                workbook.write(out);
                out.close();

//            Desktop desktop = Desktop.getDesktop();
//            if (file.exists()) {
//                desktop.open(file);
//            }
            } catch (Exception ex) {
//                JOptionPane.showMessageDialog(null, "( ˘︹˘ ) Algo salio mal!", "Oh, no!", JOptionPane.ERROR_MESSAGE);
            }
        } else {
            try {
                File file = new File(ruta + nombre + ".xlsx");
                FileOutputStream out = new FileOutputStream(file);
                workbook.write(out);
                out.close();

//            Desktop desktop = Desktop.getDesktop();
//            if (file.exists()) {
//                desktop.open(file);
//            }
            } catch (Exception ex) {
//                JOptionPane.showMessageDialog(null, "( ˘︹˘ ) Algo salio mal!", "Oh, no!", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    public static Document convertStringToDocument(String xmlStr) {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder;

            builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new InputSource(new StringReader(xmlStr)));
            return doc;
        } catch (Exception e) {
            return null;
        }
    }

    public static void writeXml(Document doc, OutputStream output) throws TransformerException {
        TransformerFactory transformerFactory = TransformerFactory.newInstance();
        Transformer transformer = transformerFactory.newTransformer();
        DOMSource source = new DOMSource((Node) doc);
        StreamResult result = new StreamResult(output);

        transformer.transform(source, result);
    }

    public static void enviarCorreo(String nombre, String rutaArchivo) {
        try {
            String dir = System.getProperty("user.dir");
            Properties props = new Properties();
            props.put("mail.smtp.host", "smtppro.zoho.com");
            props.put("mail.smtp.port", "587");
            props.put("mail.smtp.auth", "true");
            props.put("mail.smtp.starttls.enable", "true");
            Session session = Session.getInstance(props, (javax.mail.Authenticator) null);
            session.setDebug(false);

            BodyPart texto = new MimeBodyPart();
            texto.setText("");

            MimeMultipart multiParte = new MimeMultipart();
            multiParte.addBodyPart(texto);

            MimeBodyPart adjunto = new MimeBodyPart();
            adjunto.setDataHandler(new DataHandler(new FileDataSource(rutaArchivo)));
            adjunto.setFileName(nombre);
            multiParte.addBodyPart(adjunto);

            MimeMessage message = new MimeMessage(session);
            message.setContent(multiParte);
            message.setFrom(new InternetAddress("soporte@ftoso.cl"));
            message.addRecipient(Message.RecipientType.TO, new InternetAddress("contabilidad@ftoso.cl"));
            message.setSubject(nombre);

            javax.mail.Transport t = session.getTransport("smtp");
            t.connect("soporte@ftoso.cl", "qweASDzxc123*");
            t.sendMessage(message, message.getAllRecipients());
            t.close();
        } catch (Exception var16) {
            System.out.println(var16);
        }
    }

    public static void procedimientoPrincipal(int num, int num2) throws InterruptedException, IOException {
        for (int i = 0; i < 4; i++) {

            Date dateBefore30Days = DateUtils.addDays(new Date(), num);
            DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
            String strDate = dateFormat.format(dateBefore30Days);

            System.out.println("strDate " + strDate);

            switch (i) {
                case 0:
                    System.out.println("Descarga BOLETA ENVIADOS");

                    ArrayList<Documento> arrDocumento = new ArrayList<>();
                    ArrayList<String> arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("ENVIADOS", strDate, strDate, 1, "CENTRALIZA BOLETA DE VENTA", arrDocumento, arrRuts, 0, num2);
                    Logica.descargarBoletas("ENVIADOS", strDate, strDate, 2, "CENTRALIZA BOLETA DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                case 1:
                    System.out.println("Descarga FACTURAS ENVIADOS");

                    arrDocumento = new ArrayList<>();
                    arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("ENVIADOS", strDate, strDate, 4, "CENTRALIZA FACTURAS DE VENTA", arrDocumento, arrRuts, 0, num2);
                    Logica.descargarBoletas("ENVIADOS", strDate, strDate, 5, "CENTRALIZA FACTURAS DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                case 2:
                    System.out.println("Descarga NOTA DE CREDITO ENVIADOS");

                    arrDocumento = new ArrayList<>();
                    arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("ENVIADOS", strDate, strDate, 8, "CENTRALIZA NOTA DE CREDITO DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                case 3:
                    System.out.println("Descarga NOTA DE DEBITO ENVIADOS");

                    arrDocumento = new ArrayList<>();
                    arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("ENVIADOS", strDate, strDate, 9, "CENTRALIZA NOTA DE DEBITO DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                default:
                    break;
            }
        }
        for (int i = 0; i < 4; i++) {
            Date dateBefore30Days = DateUtils.addDays(new Date(), -0);
            DateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
            String strDate = dateFormat.format(dateBefore30Days);

            System.out.println("strDate " + strDate);

            switch (i) {
                case 0:
                    ArrayList<Documento> arrDocumento = new ArrayList<>();
                    ArrayList<String> arrRuts = new ArrayList<>();
//                    Logica.descargarBoletas("RECIBIDOS", strDate, strDate, 1, "CENTRALIZA BOLETA DE VENTA", arrDocumento, arrRuts, 0);
//                    Logica.descargarBoletas("RECIBIDOS", strDate, strDate, 2, "CENTRALIZA BOLETA DE VENTA", arrDocumento, arrRuts, 1);
                    break;
                case 1:
                    System.out.println("Descarga BOLETA RECIBIDOS");

                    arrDocumento = new ArrayList<>();
                    arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("RECIBIDOS", strDate, strDate, 4, "CENTRALIZA FACTURAS DE VENTA", arrDocumento, arrRuts, 0, num2);
                    Logica.descargarBoletas("RECIBIDOS", strDate, strDate, 5, "CENTRALIZA FACTURAS DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                case 2:
                    System.out.println("Descarga NOTA DE CREDITO RECIBIDOS");

                    arrDocumento = new ArrayList<>();
                    arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("RECIBIDOS", strDate, strDate, 8, "CENTRALIZA NOTA DE CREDITO DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                case 3:
                    System.out.println("Descarga  NOTA DE DEBITO RECIBIDOS");

                    arrDocumento = new ArrayList<>();
                    arrRuts = new ArrayList<>();
                    Logica.descargarBoletas("RECIBIDOS", strDate, strDate, 9, "CENTRALIZA NOTA DE DEBITO DE VENTA", arrDocumento, arrRuts, 1, num2);
                    break;
                default:
                    break;
            }
        }
    }
}
