/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package Ventanas;

import Logica.Logica;
import java.awt.AWTException;
import java.awt.Image;
import java.awt.MenuItem;
import java.awt.PopupMenu;
import java.awt.SystemTray;
import java.awt.TrayIcon;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.imageio.ImageIO;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author sopor
 */
public class PrincipalNuevo {

    public static void main(String[] args) throws InterruptedException, IOException, AWTException {
        if (SystemTray.isSupported()) {
            try {
                SystemTray tray = SystemTray.getSystemTray();
                Image icon = ImageIO.read(new File("image.png"));
                Image resizedImage = icon.getScaledInstance(16, 16, Image.SCALE_SMOOTH);
                TrayIcon trayIcon = new TrayIcon(resizedImage, "Contabilidad");
                tray.add(trayIcon);

                MenuItem restartItem = new MenuItem("Reiniciar");
                restartItem.addActionListener(new ActionListener() {
                    public void actionPerformed(ActionEvent e) {
                        String javaCmd = System.getProperty("java.home") + "/bin/java";
                        String classPath = System.getProperty("java.class.path");
                        String mainClass = System.getProperty("sun.java.command");

                        List<String> command = new ArrayList<>();
                        command.add(javaCmd);
                        command.add("-cp");
                        command.add(classPath);
                        command.add(mainClass);

                        try {
                            ProcessBuilder builder = new ProcessBuilder(command);
                            builder.start();
                            System.exit(0);
                        } catch (IOException ex) {
                            ex.printStackTrace();
                        }
                    }
                });
                PopupMenu popup = new PopupMenu();

                MenuItem cerrartItem = new MenuItem("Cerrar");
                cerrartItem.addActionListener(new ActionListener() {
                    public void actionPerformed(ActionEvent e) {
                        System.exit(0);
                    }
                });
                popup.add(restartItem);
                popup.add(cerrartItem);

                // Asignar el menú de contexto al icono de bandeja
                trayIcon.setPopupMenu(popup);
            } catch (AWTException e) {
                System.err.println("No se pudo agregar el icono de bandeja.");
            }
        }

        Thread thread = new Thread(() -> {
            while (true) {
                try {
                    LocalTime horaActual = LocalTime.now();

                    File directory0 = new File("C:\\SOFTLAND\\DATOS\\CWIN");
                    File directory1 = new File(System.getProperty("user.dir") + "\\BOLETAENVIADA");
                    File directory2 = new File(System.getProperty("user.dir") + "\\FACTURAENVIADA");
                    File directory3 = new File(System.getProperty("user.dir") + "\\NCENVIADA");
                    File directory4 = new File(System.getProperty("user.dir") + "\\DNENVIADA");
                    File directory5 = new File(System.getProperty("user.dir") + "\\BOLETARECIBIDA");
                    File directory6 = new File(System.getProperty("user.dir") + "\\FACTURARECIBIDA");
                    File directory7 = new File(System.getProperty("user.dir") + "\\NCRECIBIDA");
                    File directory8 = new File(System.getProperty("user.dir") + "\\DNRECIBIDA");

                    // Agrega aquí el bloque de código que deseas que se ejecute si la hora actual es 1:00:00 y es lunes.
                    if (LocalDate.now().getDayOfWeek() != DayOfWeek.SATURDAY
                            && LocalDate.now().getDayOfWeek() != DayOfWeek.SUNDAY) {

                        System.out.println("horaActual.getHour() " + horaActual.getHour());
                        System.out.println("oraActual.getMinute() " + horaActual.getMinute());
                        System.out.println("horaActual.getSecond() " + horaActual.getSecond());

                        if (horaActual.getHour() == 15 && horaActual.getMinute() == 51 && horaActual.getSecond() == 0) {
                            eliminar(directory0);
                            eliminar(directory1);
                            eliminar(directory2);
                            eliminar(directory3);
                            eliminar(directory4);
                            eliminar(directory5);
                            eliminar(directory6);
                            eliminar(directory7);
                            eliminar(directory8);

                            if (LocalDate.now().getDayOfWeek() == DayOfWeek.MONDAY) {
                                Logica.procedimientoPrincipal(-1, 1);
                                Logica.procedimientoPrincipal(-2, 2);
                                Logica.procedimientoPrincipal(-3, 3);

                                pegarExcels(directory1, System.getProperty("user.dir") + "\\BOLETAENVIADA\\BOLETA.xlsx");
                                pegarExcels(directory2, System.getProperty("user.dir") + "\\FACTURAENVIADA\\FACTURA.xlsx");
                                pegarExcels(directory3, System.getProperty("user.dir") + "\\NCENVIADA\\NOTACREDITO.xlsx");
                                pegarExcels(directory4, System.getProperty("user.dir") + "\\DNENVIADA\\NOTADEBITO.xlsx");
                                pegarExcels(directory5, System.getProperty("user.dir") + "\\BOLETARECIBIDA\\BOLETA.xlsx");
                                pegarExcels(directory6, System.getProperty("user.dir") + "\\FACTURARECIBIDA\\FACTURA.xlsx");
                                pegarExcels(directory7, System.getProperty("user.dir") + "\\NCRECIBIDA\\NOTACREDITO.xlsx");
                                pegarExcels(directory8, System.getProperty("user.dir") + "\\DNRECIBIDA\\NOTADEBITO.xlsx");
                            } else {
                                Logica.procedimientoPrincipal(-1, 999);
                            }
                            LocalDate fecha = LocalDate.now();
                            DateTimeFormatter formato = DateTimeFormatter.ofPattern("ddMMyyyy");
                            String fechaFormateada = fecha.format(formato);
//                        pegarExcelsFinal(System.getProperty("user.dir") + "\\TOSO" + fechaFormateada + ".xlsx");
                            pegarExcelsFinal("C:\\SOFTLAND\\DATOS\\CWIN\\TOSO" + fechaFormateada + ".xlsx");

                        }
                    }
                    Thread.sleep(1000);
                } catch (InterruptedException | IOException ex) {
                    Logger.getLogger(PrincipalNuevo.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
        );
        thread.start();
    }

    public static void pegarExcels(File carpeta, String rutaNombre) throws FileNotFoundException, IOException {
        ArrayList<XSSFSheet> arrXSSFSheet = new ArrayList<>();

        File[] archivos = carpeta.listFiles();
        if (archivos != null) {
            for (File archivo : archivos) {
                if (archivo.isFile()) {
                    if (!archivo.getName().contains("CLIENTE")) {
                        FileInputStream archivo1 = new FileInputStream(archivo);
                        XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
                        XSSFSheet hoja1 = libro1.getSheetAt(0);

                        arrXSSFSheet.add(hoja1);
                    }
                }
            }
        }

        if (!arrXSSFSheet.isEmpty()) {

            System.out.println("arrXSSFSheet.size() " + arrXSSFSheet.size());

            XSSFWorkbook nuevoLibro = new XSSFWorkbook();
            XSSFSheet nuevaHoja = nuevoLibro.createSheet("Hoja1");

            int filaActual = 0;
            for (int i = 0; i < arrXSSFSheet.size(); i++) {
                XSSFSheet get = arrXSSFSheet.get(i);

                for (int x = 0; x <= get.getLastRowNum(); x++) {
                    Row fila = get.getRow(x);
                    XSSFRow nuevaFila = nuevaHoja.createRow(filaActual++);

                    for (int j = 0; j < fila.getLastCellNum(); j++) {
                        Cell celda = fila.getCell(j);
                        Cell nuevaCelda = nuevaFila.createCell(j);

                        if (celda.getCellType() == CellType.NUMERIC) {
                            nuevaCelda.setCellValue(celda.getNumericCellValue());
                        } else if (celda.getCellType() == CellType.STRING) {
                            nuevaCelda.setCellValue(celda.getStringCellValue());
                        } else if (celda.getCellType() == CellType.BOOLEAN) {
                            nuevaCelda.setCellValue(celda.getBooleanCellValue());
                        }
                    }
                }
            }
            FileOutputStream archivoNuevo = new FileOutputStream(rutaNombre);
            nuevoLibro.write(archivoNuevo);
            archivoNuevo.close();
            nuevoLibro.close();
        }
    }

    public static void pegarExcelsFinal(String rutaNombre) throws FileNotFoundException, IOException {
        ArrayList<XSSFSheet> arrXSSFSheet = new ArrayList<>();

        try {
            File file1 = new File(System.getProperty("user.dir") + "\\BOLETAENVIADA\\BOLETA.xlsx");
            FileInputStream archivo1 = new FileInputStream(file1);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {

            File file2 = new File(System.getProperty("user.dir") + "\\FACTURAENVIADA\\FACTURA.xlsx");
            FileInputStream archivo1 = new FileInputStream(file2);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {
            File file3 = new File(System.getProperty("user.dir") + "\\NCENVIADA\\NOTACREDITO.xlsx");
            FileInputStream archivo1 = new FileInputStream(file3);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {
            File file4 = new File(System.getProperty("user.dir") + "\\DNENVIADA\\NOTADEBITO.xlsx");
            FileInputStream archivo1 = new FileInputStream(file4);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {
            File file5 = new File(System.getProperty("user.dir") + "\\BOLETARECIBIDA\\BOLETA.xlsx");
            FileInputStream archivo1 = new FileInputStream(file5);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {
            File file6 = new File(System.getProperty("user.dir") + "\\FACTURARECIBIDA\\FACTURA.xlsx");
            FileInputStream archivo1 = new FileInputStream(file6);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {
            File file7 = new File(System.getProperty("user.dir") + "\\NCRECIBIDA\\NOTACREDITO.xlsx");
            FileInputStream archivo1 = new FileInputStream(file7);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }
        try {
            File file8 = new File(System.getProperty("user.dir") + "\\DNRECIBIDA\\NOTADEBITO.xlsx");
            FileInputStream archivo1 = new FileInputStream(file8);
            XSSFWorkbook libro1 = new XSSFWorkbook(archivo1);
            XSSFSheet hoja1 = libro1.getSheetAt(0);
            arrXSSFSheet.add(hoja1);
        } catch (Exception ex) {

        }

        if (!arrXSSFSheet.isEmpty()) {

            System.out.println("arrXSSFSheet.size() " + arrXSSFSheet.size());

            XSSFWorkbook nuevoLibro = new XSSFWorkbook();
            XSSFSheet nuevaHoja = nuevoLibro.createSheet("Hoja1");

            int filaActual = 0;
            for (int i = 0; i < arrXSSFSheet.size(); i++) {
                XSSFSheet get = arrXSSFSheet.get(i);

                for (int x = 0; x <= get.getLastRowNum(); x++) {
                    Row fila = get.getRow(x);
                    XSSFRow nuevaFila = nuevaHoja.createRow(filaActual++);

                    for (int j = 0; j < fila.getLastCellNum(); j++) {
                        Cell celda = fila.getCell(j);
                        Cell nuevaCelda = nuevaFila.createCell(j);

                        if (celda.getCellType() == CellType.NUMERIC) {
                            nuevaCelda.setCellValue(celda.getNumericCellValue());
                        } else if (celda.getCellType() == CellType.STRING) {
                            nuevaCelda.setCellValue(celda.getStringCellValue());
                        } else if (celda.getCellType() == CellType.BOOLEAN) {
                            nuevaCelda.setCellValue(celda.getBooleanCellValue());
                        }
                    }
                }
            }
            FileOutputStream archivoNuevo = new FileOutputStream(rutaNombre);
            nuevoLibro.write(archivoNuevo);
            archivoNuevo.close();
            nuevoLibro.close();
        }
    }

    public static void eliminar(File directory) {
        if (directory.exists() && directory.isDirectory()) {
            File[] files = directory.listFiles();
            for (File file : files) {
                if (!file.delete()) {
                    System.err.println("No se pudo eliminar el archivo " + file.getName());
                }
            }
        }
    }

    public static void moverFile(String rutaOrigen, String rutaDestino) {
        File archivoOrigen = new File(rutaOrigen);
        File archivoDestino = new File(rutaDestino);

        if (!archivoOrigen.exists()) {
            System.out.println("El archivo de origen no existe.");
        } else if (archivoDestino.exists()) {
            System.out.println("El archivo de destino ya existe.");
        } else {
            boolean resultado = archivoOrigen.renameTo(archivoDestino);
            if (resultado) {
                System.out.println("Archivo movido exitosamente.");
            } else {
                System.out.println("No se pudo mover el archivo.");
            }
        }
    }
}
