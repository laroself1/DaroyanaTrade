/**
 * Created by LaroSelf on 20.05.2016.
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import Beans.NumberOfSummaries;
import Beans.Truck;
import com.sun.xml.internal.ws.policy.privateutil.PolicyUtils;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.TabPane;
import javafx.scene.image.Image;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.text.Text;
import javafx.stage.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import java.time.LocalDate;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.time.ZoneId;
import java.util.*;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.HPos;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.layout.*;
import javafx.scene.text.Font;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.control.Button;

public class Main extends Application{
    @Override
    public void start(Stage primaryStage) throws Exception {
        VBox summaryVBox1 = new VBox(6);
        VBox registersVBox = new VBox(6);
        HBox doRegisters = new HBox();
        VBox consignor = new VBox(10);
        consignor.setStyle("-fx-padding: 8;");

        summaryVBox1.setStyle("-fx-padding: 10;");
        VBox summaryVBoxButton = new VBox();
        summaryVBoxButton.setStyle("-fx-padding: 8;");
        VBox registerVBoxButton = new VBox();
        registerVBoxButton.setStyle("-fx-padding: 8;");
        consignor.setAlignment(Pos.TOP_CENTER);
        registerVBoxButton.setAlignment(Pos.CENTER_RIGHT);
        summaryVBoxButton.setAlignment(Pos.TOP_CENTER);
        doRegisters.setAlignment(Pos.CENTER);

        registersVBox.setStyle("-fx-padding: 10;");
        HBox consi = new HBox(12);
        consi.setAlignment(Pos.CENTER);
        HBox manage = new HBox(12);
        manage.setAlignment(Pos.CENTER);
        HBox summaryMainTFandFileChoose = new HBox(8);
        HBox summaryInputTFandFileChoose = new HBox(8);
        HBox sumFromFileChoose = new HBox(8);
        HBox regToFileChoose = new HBox(8);

        StackPane root = new StackPane();
        TabPane tabPane = new TabPane();
        tabPane.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);
        Tab summaryTab = new Tab("Сводка");
        Tab registersTab = new Tab("Реестр");

        Label fromReport = new Label("From");
        Label empty = new Label(" ");
        Label empty2 = new Label(" ");
        Label empty3 =new Label("               ");
        empty.setFont(new Font("Arial", 2));
        empty2.setFont(new Font("Arial", 2));
        Label summaryTF1Label = new Label(" Входящий отчет");
        summaryTF1Label.setFont(new Font("Arial", 12));
        Label registerTF1Label = new Label("Реестр");
        registerTF1Label.setFont(new Font("Arial", 12));
        Label summaryMainLabel = new Label(" Общая сводка");
        summaryMainLabel.setFont(new Font("Arial", 12));
        Label summaryMainFromLabel = new Label("Общая сводка");
        summaryMainFromLabel.setFont(new Font("Arial", 12));

        TextField summaryInputTF = new TextField();
        HBox.setHgrow(summaryInputTF, Priority.ALWAYS);
        TextField summaryMainTF = new TextField();
        HBox.setHgrow(summaryMainTF, Priority.ALWAYS);
        TextField registerTF1 = new TextField();
        HBox.setHgrow(registerTF1, Priority.ALWAYS);
        TextField summaryMainFromTF = new TextField();
        HBox.setHgrow(summaryMainFromTF, Priority.ALWAYS);

        Button doMainSummary = new Button("Внести данные");
        doMainSummary.setFont(new Font("Arial", 12));
        Button doRegistersBut = new Button("Внести данные");
        Button doRegistersBut2 = new Button("Внести данные");
        Button readMainSum = new Button("Считать поставщиков");
        readMainSum.setFont(new Font("Arial", 12));
        doRegistersBut.setFont(new Font("Arial", 12));
        doRegistersBut2.setFont(new Font("Arial", 12));
       // doRegistersBut.setAlignment(Pos.CENTER);
        Button fileChInputSummary = new Button("...");
        Button fileChOutMainSummary = new Button("...");
        Button fileRegisterOut = new Button("...");
        Button mainSummaryFrom = new Button("...");
        Button readManagersBut = new Button("Считать менеджеров");
        readManagersBut.setFont(new Font("Arial", 12));

        ComboBox consignorCo = new ComboBox();
        ComboBox managerCo = new ComboBox();
        HBox tg = new HBox(2);
        tg.setAlignment(Pos.CENTER);
        ToggleGroup reportTG = new ToggleGroup();
        RadioButton a = new RadioButton("Ума Агро");
        RadioButton b = new RadioButton("Global Commodities");
        a.setToggleGroup(reportTG);
        b.setToggleGroup(reportTG);
        tg.getChildren().addAll(a,b)
       ;
readManagersBut.setOnAction(new EventHandler<ActionEvent>() {
    @Override
    public void handle(ActionEvent event) {
    try{
        try{
            String str = summaryMainFromTF.getText();
            List<String> managList = FileOperations.getManagersNames(str);
        String chosen = "HA&CO LP";

//        String fO= managerCo.getSelectionModel().getSelectedItem().toString();
    //FileOperations.getManagersTrucks(fO,chosen);

            managerCo.getItems().setAll(managList);
            manage.getChildren().add(managerCo);
            manage.getChildren().add(doRegistersBut2);




    }catch (OpenXML4JException e){e.printStackTrace();}
    }catch (IOException a){a.printStackTrace();}
    }

});
        doRegistersBut2.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                try{
                    try{


                        String chosen = null;

                       chosen =managerCo.getSelectionModel().getSelectedItem().toString();
                        String fileout = registerTF1.getText();
                        System.out.println(chosen);

                        List<Truck> mainSum = FileOperations.getManagersTrucks(summaryMainFromTF.getText(), chosen);

                        List <Truck>reg = FileOperations.getRegister(registerTF1.getText());

                        List <Truck> res =FileOperations.compareRandS(mainSum,reg);

                        FileOperations.saveNewRegister2(fileout,res);

                        StackPane secondaryLayout = new StackPane();

                        Scene secondScene = new Scene(secondaryLayout, 250, 100);
                        Stage secondStage = new Stage();
                        Button ok = new Button("Ok");
                        ok.setMinWidth(40);
                        Label did =new Label("Данные успешно внесены");
                        Label empty =new Label();
                        empty.setFont(new Font("Arial", 3));
                        VBox vbox = new VBox(2);
                        vbox.getChildren().addAll(did,empty,ok);
                        secondaryLayout.getChildren().addAll(vbox);
                        vbox.setAlignment(Pos.CENTER);
                        ok.setOnAction(new EventHandler<ActionEvent>() {
                            @Override
                            public void handle(ActionEvent event) {
                                secondStage.close();
                                registerTF1.clear();
                                consignorCo.setValue(null);

                            }
                        });

                        secondStage.initModality(Modality.APPLICATION_MODAL);
                        secondStage.setTitle("Сообщение");
                        secondStage.setScene(secondScene);
                        secondStage.show();




                    }catch (IOException e){e.printStackTrace();}
                }catch (OpenXML4JException a){a.printStackTrace();}
            }
        });

        readMainSum.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                try{
                    try{
                        List arrCons = new ArrayList<String>();


                              List<String> res= null;
                        res =   FileOperations.getMainSummarryCons(
                                FileOperations.getMainSummarry(summaryMainFromTF.getText()));

                        java.util.Collections.sort(res);

                        consignorCo.getItems().setAll(res);
                        consi.getChildren().add(consignorCo);
                        consi.getChildren().add(doRegistersBut);


                    }catch (IOException e){e.printStackTrace();}
                }catch (OpenXML4JException a){a.printStackTrace();}
                //  arrCons.add(FileOperations.getMainSummarryCons(summaryMainFromTF.getText()));
                ;
            }
        });
        fileRegisterOut.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                registerTF1.setText(FileChoose.chooseFile());
            }
        });
        mainSummaryFrom.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                summaryMainFromTF.setText(FileChoose.chooseFile());

            }
        });

        fileChInputSummary.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent arg0) {
                summaryInputTF.setText(FileChoose.chooseFile());
            }
        });
        fileChOutMainSummary.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent arg0) {
                summaryMainTF.setText(FileChoose.chooseFile());
            }
        });
        doRegistersBut.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
               try{
                   try{


                       String cons = null;
                    cons =consignorCo.getSelectionModel().getSelectedItem().toString();
                      System.out.println( cons);
                       List<Truck> mainSum = FileOperations.getMainSummarry(summaryMainFromTF.getText());
                       List <Truck>reg = FileOperations.getRegister(registerTF1.getText());

                       List <Truck> res =FileOperations.compareRandS(mainSum,reg);

                       String fileout = registerTF1.getText();
                       FileOperations.saveNewRegister(fileout,res,cons);

                       StackPane secondaryLayout = new StackPane();

                       Scene secondScene = new Scene(secondaryLayout, 250, 100);
                       Stage secondStage = new Stage();
                       Button ok = new Button("Ok");
                       ok.setMinWidth(40);
                       Label did =new Label("Данные успешно внесены");
                       Label empty =new Label();
                       empty.setFont(new Font("Arial", 3));
                       VBox vbox = new VBox(2);
                       vbox.getChildren().addAll(did,empty,ok);
                       secondaryLayout.getChildren().addAll(vbox);
                       vbox.setAlignment(Pos.CENTER);
                       ok.setOnAction(new EventHandler<ActionEvent>() {
                           @Override
                           public void handle(ActionEvent event) {
                               secondStage.close();
                               registerTF1.clear();
                                consignorCo.setValue(null);

                           }
                       });

                       secondStage.initModality(Modality.APPLICATION_MODAL);
                       secondStage.setTitle("Сообщение");
                       secondStage.setScene(secondScene);
                       secondStage.show();




            }catch (IOException e){e.printStackTrace();}
        }catch (OpenXML4JException a){a.printStackTrace();}
            }
        });
        doMainSummary.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                String in =summaryInputTF.getText().toString();
                String out = summaryMainTF.getText().toString();
                try{  try {

                   try{
                       List<Truck> inp = null;
                       List<Truck> outp = FileOperations.getMainSummarry(out);
                       String rep = null;
                       if (a.isSelected()) {
                           rep = "Ума Агро";
                          inp= FileOperations.getUmaAgro(in);
                       }
                       if (b.isSelected()) {
                          rep = "Global Commodities";
                           inp = FileOperations.getGlobal(in);
                       }
                     ;
                       System.out.println(rep);
     List<Truck> result=FileOperations.compareRandS(inp,outp);
                       FileOperations.saveNewSummary(out,result);



                       StackPane secondaryLayout = new StackPane();

                       Scene secondScene = new Scene(secondaryLayout, 250, 100);
                       Stage secondStage = new Stage();
                       Button ok = new Button("Ok");
                       ok.setMinWidth(40);
                       Label did =new Label("Данные успешно внесены");
                       Label empty =new Label();
                       empty.setFont(new Font("Arial", 3));
                       VBox vbox = new VBox(2);
                       vbox.getChildren().addAll(did,empty,ok);
                       secondaryLayout.getChildren().addAll(vbox);
                       vbox.setAlignment(Pos.CENTER);
                       ok.setOnAction(new EventHandler<ActionEvent>() {
                           @Override
                           public void handle(ActionEvent event) {
                               secondStage.close();
                               summaryInputTF.clear();
                               reportTG.selectToggle(null);

                           }
                       });

                       secondStage.initModality(Modality.APPLICATION_MODAL);
                       secondStage.setTitle("Сообщение");
                       secondStage.setScene(secondScene);
                       secondStage.show();

                    }catch (InvalidFormatException i){i.printStackTrace();}
                }catch (IOException e){e.printStackTrace();}
            }catch (OpenXML4JException a){a.printStackTrace();}
            }
        });

        consi.getChildren().addAll(readMainSum );
        manage.getChildren().addAll(readManagersBut);
        doRegisters.getChildren().addAll();
        registerVBoxButton.getChildren().addAll();
        consignor.getChildren().addAll(consi,manage);
        summaryTab.setContent(summaryVBox1);
        registersTab.setContent(registersVBox);
        tabPane.getTabs().addAll(summaryTab,registersTab);
        root.getChildren().addAll(tabPane);
        ComboBox numberOfSummaries = new ComboBox();
        numberOfSummaries.setItems(FXCollections.observableArrayList(NumberOfSummaries.values()));


        summaryInputTFandFileChoose.getChildren().addAll(summaryInputTF,fileChInputSummary);
        summaryMainTFandFileChoose.getChildren().addAll(summaryMainTF,fileChOutMainSummary);
        sumFromFileChoose.getChildren().addAll(summaryMainFromTF,mainSummaryFrom);
        regToFileChoose.getChildren().addAll(registerTF1,fileRegisterOut);
        // chooseReport.getChildren().addAll(fromReport,numberOfSummaries);

        summaryVBoxButton.getChildren().addAll(doMainSummary);
        summaryVBox1.getChildren().addAll(empty,summaryTF1Label, summaryInputTFandFileChoose,empty3,tg,summaryMainLabel,
                summaryMainTFandFileChoose,summaryVBoxButton);

        registersVBox.getChildren().addAll(empty2,summaryMainFromLabel,sumFromFileChoose,registerTF1Label,
                regToFileChoose,consignor);

        Scene scene = new Scene(root, 800, 246);
    //   primaryStage.getIcons().add(new Image("C:\\Users\\LaroSelf\\IdeaProjects\\bnnbn\\src\\no-copy-icon.png"));
        primaryStage.setScene(scene);
        primaryStage.setTitle("Помощник");
        primaryStage.show();
    }}