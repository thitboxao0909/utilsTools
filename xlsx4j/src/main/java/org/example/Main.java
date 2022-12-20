package org.example;

import com.aspose.cells.*;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Collection;
import java.util.List;
import java.util.ArrayList;
import java.util.Random;

public class Main {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        String inputFile = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\Cats.xlsx";
        String outputPath = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\CatsOut.xlsx";

        Human human = new Human("Hoang", "thitboxao0909@gmail.com");

        // Meo
        final String[] catName = {"Tom", "Jerry", "Fish", "Muop", "MeoBeo", "Den", "Trang", "Tam the", "Catty Mc CatFace", "C","Doremon","Nobita", "Map"};
        List<Meo> Meo = new ArrayList<Meo>();
        Random random = new Random();
        for (int i = 0; i < 300; i++) {
            int index = random.nextInt(catName.length);
            Meo m = new Meo(catName[index], random.nextInt(10));
            Meo.add(m);
        }
        Meo.add(new Meo("Tom", 1));
        Meo.add(new Meo("Jerry", 3));
        Meo.add(new Meo("Muop", 2));

        Human boss = new Human("Hoang", "thitboxao0909@gmail.com");


        //Workbook newBook = new Workbook();
        Workbook workbook = new Workbook(inputFile);
        Worksheet worksheet = workbook.getWorksheets().get(0);

// Exporting the array of names to first row and first column vertically
        Cells cells = worksheet.getCells();
        Cell nameCell = cells.get("B1");
        nameCell.setValue(boss.getName());

        Cell emailCell = cells.get("B2");
        emailCell.setValue(boss.getEmail());


        worksheet.getCells().importCustomObjects((Collection) Meo,
                                                new String[] { "Name", "Age" },
                                false, 4, 0,
                                                Meo.size(), true,
                                   null, false);
// Saving the Excel file
        workbook.save(outputPath);
    }

    public void ExcelFileOut(List<Object> list,
                             Integer firstRow,
                             Integer firstColumn,
                             Integer sheetNum,
                             String pathTemplate,
                             String pathOut) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        worksheet.getCells().importCustomObjects((Collection) list,
                new String[] {},
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);
        // Saving the Excel file
        workbook.save(pathOut);

    }
    public void ExcelFileOutWithPassword(List<Object> list,
                                         Integer firstRow,
                                         Integer firstColumn,
                                         Integer sheetNum,
                                         String pathTemplate,
                                         String pathOut,
                                         String passWord) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        worksheet.getCells().importCustomObjects((Collection) list,
                new String[] {},
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);

        // Password protect the file.
        workbook.getSettings().setPassword(passWord);
        // Specify XOR encrption type.
        workbook.setEncryptionOptions(EncryptionType.XOR, 40);
        // Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
        workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);
        // Saving the Excel file
        workbook.save(pathOut);
    }
    public OutputStream ExcelOutputStreamOut(List<Object> list,
                                             Integer firstRow,
                                             Integer firstColumn,
                                             Integer sheetNum,
                                             String pathTemplate,
                                             String pathOut) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        worksheet.getCells().importCustomObjects((Collection) list,
                new String[] {},
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);
        // Saving the Excel file
        workbook.save(outputStream, new XpsSaveOptions(FileFormatType.XLSX));
        return outputStream;
    }

    public OutputStream ExcelOutputStreamOutWithPassword (List<Object> list,
                                                          Integer firstRow,
                                                          Integer firstColumn,
                                                          Integer sheetNum,
                                                          String pathTemplate,
                                                          String pathOut,
                                                          String passWord) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        worksheet.getCells().importCustomObjects((Collection) list,
                new String[] {},
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);
        // Password protect the file.
        workbook.getSettings().setPassword(passWord);
        // Specify XOR encrption type.
        workbook.setEncryptionOptions(EncryptionType.XOR, 40);
        // Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
        workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);
        // Saving the Excel file
        workbook.save(outputStream, new XpsSaveOptions(FileFormatType.XLSX));
        return outputStream;
    }

}