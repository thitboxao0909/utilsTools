package org.example;

import com.aspose.cells.*;
import org.apache.commons.lang3.StringUtils;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.Random;

public class ExcelOut {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        String inputFile = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\Cats.xlsx";
        String outputPath = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\CatsOut5.xlsx";

        Human human = new Human("Hoang", "thitboxao0909@gmail.com");

        // Meo
        final String[] catName = {"Tom", "Jerry", "Fish", "Muop", "MeoBeo", "Den", "Trang", "Tam the", "Catty Mc CatFace", "C", "Doremon", "Nobita", "Map"};
        final String[] catBreeds = {"Ta", "Sphynx", "EnglishShorthair", "Siamese", "Burmese", "Doremon", "Persian", "Birman", "JapaneseBobTail"};
        List<Meo> Meo = new ArrayList<Meo>();
        Random random = new Random();
        for (int i = 0; i < 300; i++) {
            int indexName = random.nextInt(catName.length);
            int indexBreed = random.nextInt(catBreeds.length);
            Meo m = new Meo(catName[indexName], random.nextInt(10), catBreeds[indexBreed]);
            Meo.add(m);
        }
        Meo.add(new Meo("Tom", 1, "Ta"));
        Meo.add(new Meo("Jerry", 3, "Chuot"));
        Meo.add(new Meo("Muop", 2, "Doremon"));

        Field[] listField = Meo.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for (Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for (int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }

        Human boss = new Human("Hoang", "thitboxao0909@gmail.com");


        Workbook workbook = new Workbook(inputFile);
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Exporting the array of names to first row and first column vertically
        Cells cells = worksheet.getCells();
        Cell nameCell = cells.get("B1");
        nameCell.setValue(boss.getName());

        Cell emailCell = cells.get("B2");
        emailCell.setValue(boss.getEmail());


        worksheet.getCells().importCustomObjects((Collection) Meo,
                new String[]{"Age", "Name", "Breeds"},
                false, 4, 0,
                Meo.size(), true,
                null, false);
        //worksheet.getCells().importCustomObjects(Meo, 4, 0, null);
        // Saving the Excel file
        workbook.save(outputPath);


        ResultSet rs = null;
        int size = 0;
        while (rs.next()) {
            size++;
        }
        String[][] rsArray = new String[size][rs.getMetaData().getColumnCount()];
        worksheet.getCells().importArray(rsArray, 0, 0);
        //ExcelFileOut(Meo, 4, 0, 0, inputFile, outputPath);

        //reflect get all fields
        /* Field[] listField =  list.get(0).getClass().getDeclaredFields();
        List<String> fieldNames = new ArrayList<>();
        for(Field f : listField) {
            fieldNames.add(f.getName().toUpperCase());
        }

        String[] fieldArray = new String[fieldNames.size()];
        for(int i = 0; i < fieldNames.size(); i++) {
            fieldArray[i] = StringUtils.capitalize(fieldNames.get(i).toLowerCase());
        }*/
    }

    public static void ExcelFileOut(List<?> list,
                                    String[] config,
                                    Integer firstRow,
                                    Integer firstColumn,
                                    Integer sheetNum,
                                    String pathTemplate,
                                    String pathOut) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        worksheet.getCells().importCustomObjects(list,
                config,
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);

        // Saving the Excel file
        workbook.save(pathOut);

    }

    public void ExcelFileOutWithPassword(List<?> list,
                                         String[] config,
                                         Integer firstRow,
                                         Integer firstColumn,
                                         Integer sheetNum,
                                         String pathTemplate,
                                         String pathOut,
                                         String passWord) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        worksheet.getCells().importCustomObjects(list,
                config,
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
                                             String[] config,
                                             Integer firstRow,
                                             Integer firstColumn,
                                             Integer sheetNum,
                                             String pathTemplate,
                                             String pathOut) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        worksheet.getCells().importCustomObjects(list,
                config,
                false, firstRow, firstColumn,
                list.size(), true,
                null, false);
        // Saving the Excel file
        workbook.save(outputStream, new XpsSaveOptions(FileFormatType.XLSX));
        return outputStream;
    }

    public OutputStream ExcelOutputStreamOutWithPassword(List<?> list,
                                                         String[] config,
                                                         Integer firstRow,
                                                         Integer firstColumn,
                                                         Integer sheetNum,
                                                         String pathTemplate,
                                                         String pathOut,
                                                         String passWord) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        worksheet.getCells().importCustomObjects(list,
                config,
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
    public void ExcelOutResultSet(ResultSet rs,
                                  Integer firstRow,
                                  Integer firstColumn,
                                  Integer sheetNum,
                                  String pathTemplate,
                                  String pathOut) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        int size = 0;
        while (rs.next()) {
            size++;
        }
        String[][] rsArray = new String[size][rs.getMetaData().getColumnCount()];

        int i=0;
        while (rs.next()) {

            for (int j = 0; j < rs.getMetaData().getColumnCount(); j++) {
                rsArray[i][j] = rs.getString(j+1);
            }
            i++;
        }

        worksheet.getCells().importArray(rsArray, 0, 0);
        workbook.save(pathOut);
    }
    public void ExcelOutResultSetWithPassword(ResultSet rs,
                                  Integer firstRow,
                                  Integer firstColumn,
                                  Integer sheetNum,
                                  String pathTemplate,
                                  String pathOut,
                                  String passWord) throws Exception {
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        int size = 0;
        while (rs.next()) {
            size++;
        }
        String[][] rsArray = new String[size][rs.getMetaData().getColumnCount()];

        int i=0;
        while (rs.next()) {

            for (int j = 0; j < rs.getMetaData().getColumnCount(); j++) {
                rsArray[i][j] = rs.getString(j+1);
            }
            i++;
        }

        worksheet.getCells().importArray(rsArray, 0, 0);
        // Password protect the file.
        workbook.getSettings().setPassword(passWord);
        // Specify XOR encrption type.
        workbook.setEncryptionOptions(EncryptionType.XOR, 40);
        // Specify Strong Encryption type (RC4,Microsoft Strong Cryptographic Provider).
        workbook.setEncryptionOptions(EncryptionType.STRONG_CRYPTOGRAPHIC_PROVIDER, 128);
        // Saving the Excel file
        workbook.save(pathOut);
    }
    public OutputStream ExcelOutputStreamOutResultSet(ResultSet rs,
                                             Integer firstRow,
                                             Integer firstColumn,
                                             Integer sheetNum,
                                             String pathTemplate,
                                             String pathOut) throws Exception {
        FileOutputStream outputStream = new FileOutputStream(pathOut);
        Workbook workbook = new Workbook(pathTemplate);
        Worksheet worksheet = workbook.getWorksheets().get(sheetNum == null ? 0 : sheetNum);

        int size = 0;
        while (rs.next()) {
            size++;
        }
        String[][] rsArray = new String[size][rs.getMetaData().getColumnCount()];

        int i=0;
        while (rs.next()) {

            for (int j = 0; j < rs.getMetaData().getColumnCount(); j++) {
                rsArray[i][j] = rs.getString(j+1);
            }
            i++;
        }
        worksheet.getCells().importArray(rsArray, 0, 0);
        // Saving the Excel file
        workbook.save(outputStream, new XpsSaveOptions(FileFormatType.XLSX));
        return outputStream;
    }
}