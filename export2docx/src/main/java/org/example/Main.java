package org.example;

import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.TableFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.table.TblFactory;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Marshaller;
import javax.xml.namespace.QName;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");

        Customer cus = new Customer("Hoang", "Hanoi", "tbx@gmail.com");

        meo Tom = new meo("Tom", 2);
        meo Tom2 = new meo("Jerry", 1);
        List<meo> lst = new ArrayList<>();
        lst.add(Tom);
        lst.add(Tom2);
        cus.setLstMeo(lst);

        documentObjectOutput docOut = new documentObjectOutput();

        docOut.setName(cus.getName());
        docOut.setAddress(cus.getAddress());
        docOut.setEmail(cus.getEmail());
        docOut.setFileName("C:\\Users\\thitb\\Desktop\\Common\\export2docx\\src\\main\\java\\org\\example\\NameOut.docx");
        docOut.setFileTemplatePath("C:\\Users\\thitb\\Desktop\\Common\\export2docx\\src\\main\\java\\org\\example\\Name.docx");
        docOut.setMeo(cus.getLstMeo());

        String fileTemplatePath = docOut.getFileTemplatePath();
        String fileName = docOut.getFileName();

        String xml = fwdObjectToXML(docOut);
        System.out.println("xml==============================" + xml);
        InputStream is = Files.newInputStream(Paths.get(fileTemplatePath));
        WordprocessingMLPackage wordPackage = Docx4J.load(is);
        is.close();

        Docx4J.bind(wordPackage, xml,
                Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML | Docx4J.FLAG_BIND_REMOVE_SDT);

        ByteArrayOutputStream baosWord = new ByteArrayOutputStream();
        Docx4J.save(wordPackage, baosWord);
        byte[] bWord = baosWord.toByteArray();
        baosWord.close();

        MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();

        mainDocumentPart.addParagraphOfText("Danh sach cac con meo");

        int writableWidthTwips = wordPackage.getDocumentModel()
                .getSections().get(0).getPageDimensions().getWritableWidthTwips();
        int columnNumber = 3;
        int rowsNumber = lst.size() + 1;
        Tbl tbl = TblFactory.createTable(rowsNumber, 3, writableWidthTwips/columnNumber);
        //List<Object> rows = tbl.getContent();

        ObjectFactory factory = Context.getWmlObjectFactory();

        TableFinder finder = new TableFinder();
        new TraversalUtil(mainDocumentPart.getContent(), finder);

        System.out.println("Found " + finder.tblList.size()  + "tables");

        Tbl searched = (Tbl) XmlUtils.unwrap(finder.tblList.get(0));
        System.out.println(searched.toString());
        List<Object> rows = searched.getContent();


        /*int i = 0;
        for (meo meo1 : lst) {
            Object row1 = rows.get(i);
            i++;
            Tr tr1 = (Tr) row1;
            Tc cell1 = (Tc) tr1.getContent().get(0);
            P cellPara = (P) cell1.getContent().get(0);

            Text tx = factory.createText();
            R run = factory.createR();
            String name = meo1.catName;
            tx.setValue(name);
            run.getContent().add(tx);
            cellPara.getContent().add(run);
        }*/

        mainDocumentPart.addObject(tbl);

        Docx4J.save(wordPackage, new File(fileName), Docx4J.FLAG_NONE);
        System.out.println("Saved: " + fileName);
    }

    private static String fwdObjectToXML(documentObjectOutput model) throws Exception {
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