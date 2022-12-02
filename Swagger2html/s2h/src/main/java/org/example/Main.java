package org.example;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Alert;
import javafx.scene.control.TextField;
import javafx.scene.image.Image;
import javafx.scene.image.ImageView;
import javafx.scene.layout.FlowPane;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.StackPane;
import org.swagger2html.Swagger2Html;
import javafx.scene.control.Label;

import java.awt.Desktop;
import java.io.*;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.commons.io.FilenameUtils;

public class Main extends Application {
    private File selected;
    private String htmlName;

    public static void main(String[] args) throws Exception {
        System.out.println("Hello world!");
        launch(args);
    }
    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Swagger to HTML");

        GridPane root = new GridPane();
        root.setPadding(new Insets(10, 10, 10, 10));
        root.setHgap(10);
        root.setVgap(10);

        try {
            FileInputStream inputstream = new FileInputStream("./src/main/resources/Dancing.gif");
            Image image = new Image(inputstream, 200, 200, true, true);
            ImageView imageView = new ImageView(image);
            //root.setAlignment(Pos.CENTER);
            root.add(imageView, 1, 1);
        } catch (Exception ex) {
            System.out.println(ex);
        }

        FileChooser fileChooser = new FileChooser();
        TextField field = new TextField();
        field.setText("index.html");

        Button button = new Button("Select json File");
        root.add(button, 0, 1);
        Button convert = new Button("Convert");

        button.setOnAction(e -> {
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                System.out.println("Absolute Path >> " + selectedFile.getAbsolutePath());
                //System.out.println("Name >> " + selectedFile.getParent());
                if (validateFile(selectedFile)) {
                    Label label = new Label("Output HTML name:");
                    root.add(label,0,2);
                    root.add(field,1,2);
                    root.add(convert,0,3);
                    this.selected = selectedFile;
                } else {
                    this.showAlert();
                }
            }
        });

        convert.setOnAction( e -> {
            this.htmlName = field.getText();
            if (selected != null) {
                try {
                    if (convertJson(selected, htmlName)) {
                        Label success = new Label("Convert Success!!");
                        root.add(success, 0, 5);
                    }
                } catch (Exception ex) {
                    Label failure = new Label("Convert Failed >> " + ex);
                    root.add(failure, 0 , 5);
                    System.out.println(ex);
                }
            }
        });

        Scene scene = new Scene(root, 400, 350);

        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private boolean validateFile(File file) {
        if (FilenameUtils.getExtension(file.getAbsolutePath()).equals("json")) return true; else return false;
    }
    private boolean convertJson(File file ,String htmlName) throws Exception {
        Swagger2Html s2h = new Swagger2Html();
        //String filePath = args[0];
        //String jsonName = args[1];
        //String htmlName = args[2];
        String jsonPath = file.getAbsolutePath();
        String htmlPath = file.getParent() + "\\" + htmlName;
        FileOutputStream fileOutputStream = new FileOutputStream(htmlPath);
        Writer writer = new java.io.OutputStreamWriter(fileOutputStream, "utf8");
        BufferedWriter bw = new BufferedWriter(writer);

        //Writer writer = new FileWriter(htmlPath,);

        s2h.toHtml(jsonPath, bw);
        Desktop desktop = Desktop.getDesktop();
        File output = new File(htmlPath);
        desktop.open(output);
        return true;
    }

    private void showAlert() {
        Alert alert = new Alert(Alert.AlertType.WARNING);
        alert.setTitle("Warning alert");
        alert.setHeaderText("Wrong file type:");
        alert.setContentText("JSON format only!");
        alert.showAndWait();
    }
}