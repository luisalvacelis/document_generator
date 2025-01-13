package dg.natos.mp.controller;

import dg.natos.mp.ExtraMethods;
import dg.natos.mp.FileClass;
import dg.natos.mp.view.MainView;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import javax.swing.JFileChooser;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.stereotype.Component;

@Component
public class MainController extends MainView implements ActionListener {

    private FileClass fileClass;
    private ExecuteProcess executeProcess = null;
    private List<Map<String, String>> listDataMap = null;
    private int fileIndex = 0;
    private int totalFilesToCreate = 0;

    public MainController(FileClass fileClass) {
        super();
        this.fileClass = fileClass;
        this.registerObjects();
        this.registerEvents();
    }

    public void setFileClass(FileClass fileClass) {
        this.fileClass = fileClass;
    }

    private void registerEvents() {
        this.jbtExamineTemplate.addActionListener(this);
        this.jbtRegisterTemplate.addActionListener(this);
        this.jbtDeleteTemplate.addActionListener(this);
        this.jbtExamineDataExcel.addActionListener(this);
        this.jbtDownloadExcel.addActionListener(this);
        this.jbtStartProcess.addActionListener(this);
        this.jbtCancelProcess.addActionListener(this);

        this.jcbSelectTemplate.addItemListener((ItemEvent e) -> {
            if (e.getStateChange() == ItemEvent.SELECTED) {
                if (jcbSelectTemplate.getSelectedIndex() != 0) {
                    this.changeEnable(true);
                    String nameTemplate = jcbSelectTemplate.getSelectedItem().toString() + ".docx";
                    this.jtfCountVariables.setText(String.valueOf(fileClass.readAndCountVariablesTemplate(nameTemplate)));
                } else {
                    this.changeEnable(false);
                    this.clearInputs();
                }
            }
        });
    }

    private void registerObjects() {
        ExtraMethods.loadListCB(jcbSelectTemplate, fileClass.getNamesTemplates());
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (jbtExamineTemplate == e.getSource()) {
            this.examineTemplate();
        } else if (jbtRegisterTemplate == e.getSource()) {
            this.registerTemplate();
        } else if (jbtDeleteTemplate == e.getSource()) {
            this.deleteTemplate();
        } else if (jbtExamineDataExcel == e.getSource()) {
            this.examineDataExcel();
        } else if (jbtDownloadExcel == e.getSource()) {
            this.downloadExampleData();
        } else if (jbtStartProcess == e.getSource()) {
            this.startProcess();
        } else if (jbtCancelProcess == e.getSource()) {
            this.cancelProcess();
        }
    }

    private void examineTemplate() {
        JFileChooser jfc = new JFileChooser();
        if (jfc.showDialog(this, "Seleccione archivo (.docx)") == JFileChooser.APPROVE_OPTION) {
            String path = jfc.getSelectedFile().getAbsolutePath();
            if (path.endsWith("docx")) {
                this.jtfNameTemplate.setText(path);
            } else {
                ExtraMethods.sendErrorMessage("Error: Archivo no admitido, vuelva a intentarlo.");
            }
        }
    }

    private void registerTemplate() {
        String pathTemplate = jtfNameTemplate.getText();
        if (!pathTemplate.isEmpty()) {
            this.fileClass.saveTemplate(pathTemplate, fileClass.readAndCountVariablesTemplate(pathTemplate));
            this.clearInputs();
        } else {
            ExtraMethods.sendErrorMessage("Error: El archivo no seleccionado, vuelva a intentarlo.");
        }
    }

    private void deleteTemplate() {
        this.fileClass.removeTemplate(jcbSelectTemplate.getSelectedItem().toString());
        this.clearInputs();
    }
    
    
    
    private void examineDataExcel() {
        JFileChooser jfc = new JFileChooser();
        if (jfc.showDialog(this, "Seleccione archivo (.xlsx)") == JFileChooser.APPROVE_OPTION) {
            String path = jfc.getSelectedFile().getAbsolutePath();
            if (path.endsWith("xlsx")) {
                if (fileClass.validateColumnsData(path, Integer.parseInt(jtfCountVariables.getText()))) {
                    ExtraMethods.sendSuccessfullyMessage("Collumnas validadas.\nSe cargará el preview.");
                    this.jtfInputData.setText(path);
                    this.jbtStartProcess.setEnabled(true);
                    this.loadPreview();
                } else {
                    ExtraMethods.sendErrorMessage("Error: El archivo contiene errores en las columnas.");
                }
            } else {
                ExtraMethods.sendErrorMessage("Error: Archivo no admitido, vuelva a intentarlo.");
            }
        }
    }

    private void downloadExampleData() {
        JFileChooser folderChooser = new JFileChooser();
        folderChooser.setDialogTitle("Seleccione la ubicación para guardar el archivo.");
        folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int userSelection = folderChooser.showSaveDialog(null);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            File selectedFolder = folderChooser.getSelectedFile();
            String filePath = selectedFolder.getAbsolutePath() + File.separator + "example_data.xlsx";
            if (fileClass.createExcelFile(filePath)) {
                ExtraMethods.sendSuccessfullyMessage("Ha sido creado el archivo 'example_data.xlsx' en la ruta: " + filePath);
            } else {
                ExtraMethods.sendErrorMessage("Error: No se pudo crear el archivo 'example_data.xlsx'");
            }
        }
    }
    
    

    private void loadPreview() {
        String nameTemplate = jcbSelectTemplate.getSelectedItem().toString() + ".docx";
        String pathData = jtfInputData.getText();
        
        String content = fileClass.generatePreviewContent(nameTemplate, fileClass.readFirstRowFromData(pathData));
        this.jepResult.setText(content);
    }

    private void startProcess() {
        JFileChooser folderChooser = new JFileChooser();
        folderChooser.setDialogTitle("Seleccione la ubicación para guardar los archivos.");
        folderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int userSelection = folderChooser.showSaveDialog(null);
        if (userSelection == JFileChooser.APPROVE_OPTION) {
            this.fileIndex = 0;
            this.jbtCancelProcess.setEnabled(true);
            this.jtfInputData.setName(folderChooser.getSelectedFile().getAbsolutePath() + File.separator);
            
            this.listDataMap = fileClass.readRowsFromData(new File(jtfInputData.getText()));
            this.totalFilesToCreate = listDataMap.size();
            
            this.jepResult.setText("Empezando proceso\n    Archivo actual: " + fileIndex + " de " + totalFilesToCreate);
            this.executeProcess = new ExecuteProcess(this);
            this.executeProcess.start();
        }
    }

    private void cancelProcess() {
        this.clearInputs();
        this.changeEnable(false);
        this.jbtStartProcess.setEnabled(false);
        this.jbtCancelProcess.setEnabled(false);
        this.executeProcess.cancel();
        try {
            this.executeProcess.join();
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        this.executeProcess = null;
    }

    private void clearInputs() {
        this.jtfNameTemplate.setText(null);
        this.jtfInputData.setText(null);
        this.jtfCountVariables.setText(null);
        this.jbtStartProcess.setEnabled(false);
        ExtraMethods.loadListCB(jcbSelectTemplate, fileClass.getNamesTemplates());
    }

    private void changeEnable(boolean status) {
        this.jbtExamineDataExcel.setEnabled(status);
        this.jbtRegisterTemplate.setEnabled(!status);
        this.jbtExamineTemplate.setEnabled(!status);
        this.jbtDeleteTemplate.setEnabled(status);
    }

    public void executeProcess() {
        if (!listDataMap.isEmpty()) {
            File templateDocx = fileClass.getTemplate(jcbSelectTemplate.getSelectedItem().toString() + ".docx");
            String outputDirectory = jtfInputData.getName();
            try {
                Map<String, String> dataMap = listDataMap.get(0);
                XWPFDocument newDocument = fileClass.generateDocument(templateDocx, dataMap);
                String outputFilePath = outputDirectory + jcbSelectTemplate.getSelectedItem().toString() + "_" + (fileIndex + 1) + ".docx";
                try (FileOutputStream out = new FileOutputStream(outputFilePath)) {
                    newDocument.write(out);
                }
                fileIndex++;
                this.listDataMap.remove(0);
                this.jepResult.setText("Empezando proceso\n    Archivo actual: " + fileIndex + " de " + totalFilesToCreate);
                Thread.sleep(3000);
            } catch (IOException | InterruptedException e) {
                e.printStackTrace();
            }
        } else {
            ExtraMethods.sendSuccessfullyMessage("Todos los archivos fueron generados en la ruta: " + jtfInputData.getName());
            this.cancelProcess();
        }
    }
}
