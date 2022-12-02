package org.example;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.xlsx4j.exceptions.Xlsx4jException;

import java.io.File;

public class Main {
    public static void main(String[] args) throws Docx4JException {
        System.out.println("Hello world!");

        String inputFile = "C:\\Users\\thitb\\Desktop\\Common\\xlsx4j\\src\\main\\java\\org\\example\\Cats.xlsx";

        SpreadsheetMLPackage spreadsheetMLPackage = SpreadsheetMLPackage.load(new File(inputFile));



    }
}