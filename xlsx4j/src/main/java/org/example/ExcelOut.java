package org.example;

import com.aspose.cells.*;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.*;

public class ExcelOut {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        String inputFile = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\Cats.xlsx";
        String outputPath = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\CatsOut4.xlsx";

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

        Field[] listField =  Meo.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for(Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for(int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }

        Human boss = new Human("Hoang", "thitboxao0909@gmail.com");

        /*Workbook workbook = new Workbook(inputFile);
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Exporting the array of names to first row and first column vertically
        Cells cells = worksheet.getCells();
        Cell nameCell = cells.get("B1");
        nameCell.setValue(boss.getName());

        Cell emailCell = cells.get("B2");
        emailCell.setValue(boss.getEmail());


        worksheet.getCells().importCustomObjects((Collection) Meo,
                                                fieldArray,
                                false, 4, 0,
                                                Meo.size(), true,
                                   null, false);
        //worksheet.getCells().importCustomObjects(Meo, 4, 0, null);
        // Saving the Excel file
        workbook.save(outputPath);*/

        ExcelFileOut(Meo, 4, 0, 0, inputFile, outputPath);
    }

    public static void ExcelFileOut(List<?> list,
                             Integer firstRow,
                             Integer firstColumn,
                             Integer sheetNum,
                             String pathTemplate,
                             String pathOut) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        Field[] listField =  list.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for(Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for(int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }

        worksheet.getCells().importCustomObjects(list,
                fieldArray,
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);
        // Saving the Excel file
        workbook.save(pathOut);

    }
    public void ExcelFileOutWithPassword(List<?> list,
                                         Integer firstRow,
                                         Integer firstColumn,
                                         Integer sheetNum,
                                         String pathTemplate,
                                         String pathOut,
                                         String passWord) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        Field[] listField =  list.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for(Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for(int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }


        worksheet.getCells().importCustomObjects(list,
                fieldArray,
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
    public OutputStream ExcelOutputStreamOut(List<?> list,
                                             Integer firstRow,
                                             Integer firstColumn,
                                             Integer sheetNum,
                                             String pathTemplate,
                                             String pathOut) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        Field[] listField =  list.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for(Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for(int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }


        worksheet.getCells().importCustomObjects(list,
                fieldArray,
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);
        // Saving the Excel file
        workbook.save(outputStream, new XpsSaveOptions(FileFormatType.XLSX));
        return outputStream;
    }

    public OutputStream ExcelOutputStreamOutWithPassword (List<?> list,
                                                          Integer firstRow,
                                                          Integer firstColumn,
                                                          Integer sheetNum,
                                                          String pathTemplate,
                                                          String pathOut,
                                                          String passWord) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);
        Field[] listField =  list.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for(Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for(int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }


        worksheet.getCells().importCustomObjects(list,
                fieldArray,
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