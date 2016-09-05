import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.StackPane;
import javafx.scene.layout.VBox;
import javafx.scene.layout.VBoxBuilder;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;

/**
 * Created by LaroSelf on 20.05.2016.
 */
public class FileChoose {
    public static String chooseFile() {
        Stage stage = new Stage();
        FileChooser fileChooser = new FileChooser();
        FileChooser.ExtensionFilter textFilter = new FileChooser.ExtensionFilter("Excel files", "*.xls","*.xlsx","*.OOXML ");
        fileChooser.getExtensionFilters().add(textFilter);
        File file = fileChooser.showOpenDialog(stage);
        System.out.println(file);
        return file.getPath().toString();
    }
}

