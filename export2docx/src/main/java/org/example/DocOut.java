package org.example;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class DocOut {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        Customer cus = new Customer("Hoang", "Hanoi", "tbx@gmail.com");

        meo Tom = new meo("Tom", 2);
        meo Tom2 = new meo("Jerry", 1);
        meo Muop = new meo("Muop", 3);
        meo Den = new meo("Den", 1);
        List<meo> lst = new ArrayList<>();
        lst.add(Tom);
        lst.add(Tom2);
        lst.add(Muop);
        lst.add(Den);
        cus.setLstMeo(lst);

        String fileTemplatePath = "C:\\Users\\thitb\\Desktop\\Common\\export2docx\\src\\main\\java\\org\\example\\Name.docx";
        String fileOutName = "C:\\Users\\thitb\\Desktop\\Common\\export2docx\\src\\main\\java\\org\\example\\NameOut.docx";

        documentObjectOutput docOut = new documentObjectOutput();
        genWord(docOut, fileTemplatePath, fileOutName);
    }



    public static void genWord(Object o,String pathTem, String pathOut) throws Exception{
        String xml = fwdObjectToXML( o);
        System.out.println("xml==============================" + xml);

        InputStream is = Files.newInputStream(Paths.get(pathTem));
        WordprocessingMLPackage wordPackage = Docx4J.load(is);
        is.close();

        Docx4J.bind(wordPackage, xml,
                Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML | Docx4J.FLAG_BIND_REMOVE_SDT);

        ByteArrayOutputStream baosWord = new ByteArrayOutputStream();
        Docx4J.save(wordPackage, baosWord);
        byte[] bWord = baosWord.toByteArray();
        baosWord.close();

        Docx4J.save(wordPackage, new File(pathOut), Docx4J.FLAG_NONE);
        System.out.println("Saved: " + pathOut);
    }
    public static void genWordWithPassword(Object o,String pathTem, String pathOut, String password) throws Exception{
        String xml = fwdObjectToXML( o);
        System.out.println("xml==============================" + xml);

        InputStream is = Files.newInputStream(Paths.get(pathTem));
        WordprocessingMLPackage wordPackage = Docx4J.load(is);
        is.close();

        Docx4J.bind(wordPackage, xml,
                Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML | Docx4J.FLAG_BIND_REMOVE_SDT);

        ByteArrayOutputStream baosWord = new ByteArrayOutputStream();
        Docx4J.save(wordPackage, baosWord);
        byte[] bWord = baosWord.toByteArray();
        baosWord.close();

        OutputStream outStream = null;
        Docx4J.save(wordPackage, new File(pathOut), Docx4J.FLAG_NONE, password);
    }
    public static OutputStream genWord(Object o,String pathTem) throws Exception{
        String xml = fwdObjectToXML( o);
        System.out.println("xml==============================" + xml);

        InputStream is = Files.newInputStream(Paths.get(pathTem));
        WordprocessingMLPackage wordPackage = Docx4J.load(is);
        is.close();

        Docx4J.bind(wordPackage, xml,
                Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML | Docx4J.FLAG_BIND_REMOVE_SDT);

        ByteArrayOutputStream baosWord = new ByteArrayOutputStream();
        Docx4J.save(wordPackage, baosWord);
        byte[] bWord = baosWord.toByteArray();
        baosWord.close();

        OutputStream outStream = null;
        Docx4J.save(wordPackage, outStream, Docx4J.FLAG_NONE);
        //System.out.println("Saved: " + pathOut);
        return outStream;
    }

    public static OutputStream genWordWithPassword(Object o,String pathTem, String password) throws Exception{
        String xml = fwdObjectToXML( o);
        System.out.println("xml==============================" + xml);

        InputStream is = Files.newInputStream(Paths.get(pathTem));
        WordprocessingMLPackage wordPackage = Docx4J.load(is);
        is.close();

        Docx4J.bind(wordPackage, xml,
                Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML | Docx4J.FLAG_BIND_REMOVE_SDT);

        ByteArrayOutputStream baosWord = new ByteArrayOutputStream();
        Docx4J.save(wordPackage, baosWord);
        byte[] bWord = baosWord.toByteArray();
        baosWord.close();

        OutputStream outStream = null;
        Docx4J.save(wordPackage, outStream, Docx4J.FLAG_NONE, password);
        //System.out.println("Saved: " + pathOut);
        return outStream;
    }
    private static String fwdObjectToXML(Object model) throws Exception {
        String xmlContent = "";
        try {
            // Create JAXB Context
            JAXBContext jaxbContext = JAXBContext.newInstance(documentObjectOutput.class);

            // Create Marshaller
            Marshaller jaxbMarshaller = jaxbContext.createMarshaller();

            // Required formatting??
            jaxbMarshaller.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, Boolean.TRUE);

            // Print XML String to Console
            StringWriter sw = new StringWriter();

            // Write XML to StringWriter
            jaxbMarshaller.marshal(model, sw);

            // Verify XML Content
            xmlContent = sw.toString();

        } catch (JAXBException e) {
            e.printStackTrace();
        }

        return xmlContent;
    }
}