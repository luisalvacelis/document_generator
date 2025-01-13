package dg.natos.mp;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.swing.JEditorPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.bukkit.configuration.file.YamlConfiguration;
import org.springframework.stereotype.Component;

@Component
public class FileClass {

    private final File directory;
    private final File file;
    private final File directoryTemplates;
    private final YamlConfiguration configuration;

    public FileClass() {
        this.directory = new File("C:\\Document Generator");
        this.directoryTemplates = new File(directory, "templates");
        this.file = new File(directory, "config.yml");
        this.configuration = YamlConfiguration.loadConfiguration(file);
        this.createFile();
    }

    private void createFile() {
        try {
            if (!file.exists()) {
                this.directory.mkdir();
                this.directoryTemplates.mkdir();
                this.file.createNewFile();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void saveFile() {
        try {
            this.configuration.save(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void addTemplate(String templateName, int n_variables) {
        List<String> templates = configuration.getStringList("Templates");
        if (!templates.contains(templateName)) {
            templates.add(templateName);
            configuration.set("Templates", templates);
            configuration.set("Template." + templateName.replaceAll(".docx", "") + ".n_variables", n_variables);
            this.saveFile();
        } else {
            configuration.set("Template." + templateName.replaceAll(".docx", "") + ".n_variables", n_variables);
            this.saveFile();
        }
    }

    public List<String> getTemplate() {
        return configuration.getStringList("Templates");
    }

    public int getNumberOfVariables(String templateName) {
        return configuration.getInt("Config.Template." + templateName + ".n_variables", 0);
    }

    public void removeTemplate(String templateName) {
        String templateNameToRemove = templateName + ".docx";
        List<String> templates = configuration.getStringList("Templates");
        templates.remove(templateNameToRemove);
        configuration.set("Templates", templates);
        configuration.set("Template." + templateName, null);
        this.saveFile();
        File fileToRemove = new File(directoryTemplates, templateNameToRemove);
        if (fileToRemove.delete()) {
            ExtraMethods.sendSuccessfullyMessage("El archivo '" + templateNameToRemove + "' se ha eliminado.");
        } else {
            ExtraMethods.sendErrorMessage("Error: El archivo '" + templateNameToRemove + "' no se pudo eliminar.");
        }
    }

    public List<File> listTemplatesFile() {
        List<File> docxFiles = new ArrayList<>();
        File[] files = directoryTemplates.listFiles((dir, name) -> name.toLowerCase().endsWith(".docx"));

        if (files != null) {
            docxFiles.addAll(Arrays.asList(files));
        }
        return docxFiles;
    }

    public List<String> getNamesTemplates() {
        List<String> names = new ArrayList<>();
        List<File> templates = listTemplatesFile();
        for (File template : templates) {
            names.add(template.getName().replaceAll(".docx", ""));
        }
        return names;
    }

    public File getTemplate(String nameTemplate) {
        return new File(directoryTemplates, nameTemplate);
    }

    public void saveTemplate(String pathTemplate, int n_variables) {
        try {
            File template = new File(pathTemplate);
            Path sourcePath = template.toPath();
            File destinationFile = new File(directoryTemplates, template.getName());
            Path destinationPath = destinationFile.toPath();

            Files.copy(sourcePath, destinationPath, StandardCopyOption.REPLACE_EXISTING);

            this.addTemplate(template.getName(), n_variables);
            ExtraMethods.sendSuccessfullyMessage("Plantilla guardada en la ruta: " + template.getAbsolutePath());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public int readAndCountVariablesTemplate(String nameTemplate) {
        int count = 0;
        Pattern variablePattern = Pattern.compile("%VAR\\d+%");
        File template = new File(directoryTemplates, nameTemplate);
        try (FileInputStream fis = new FileInputStream(template); XWPFDocument document = new XWPFDocument(fis)) {
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                count += countVariablesInParagraph(paragraph, variablePattern);
            }
            ExtraMethods.sendSuccessfullyMessage("Total variables encontradas: " + count);
            document.close();
            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return count;
    }

    public int countVariablesInParagraph(XWPFParagraph paragraph, Pattern variablePattern) {
        int count = 0;
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs != null) {
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                if (text != null) {
                    Matcher matcher = variablePattern.matcher(text);
                    while (matcher.find()) {
                        count++;
                    }
                }
            }
        }
        return count;
    }

    public boolean validateColumnsData(String path, int totalColumns) {
        boolean status = false;
        try (FileInputStream fis = new FileInputStream(path); Workbook workbook = new XSSFWorkbook(fis)) {
            if (workbook.getNumberOfSheets() == 1) {
                Sheet sheet = workbook.getSheetAt(0);
                Row firstRow = sheet.getRow(0);
                if (firstRow != null) {
                    int columnCount = firstRow.getPhysicalNumberOfCells();
                    ExtraMethods.sendMessage("Number columns with data: " + columnCount);
                    if (columnCount == totalColumns) {
                        boolean isValid = validateColumnHeaders(firstRow, columnCount);
                        if (isValid) {
                            ExtraMethods.sendMessage("Headers are valid.");
                            status = true;
                        } else {
                            ExtraMethods.sendErrorMessage("Error: Los encabezados no son válidos.");
                        }
                    } else {
                        ExtraMethods.sendErrorMessage("Error: El total de columnas actuales no es igual al total de variables de la plantilla.");
                    }
                } else {
                    ExtraMethods.sendErrorMessage("Error: La primera fila está vacía.");
                }
            } else {
                ExtraMethods.sendErrorMessage("Error: el archivo debe contener exactamente una hoja.");
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return status;
    }

    public boolean validateColumnHeaders(Row row, int columnCount) {
        for (int i = 0; i < columnCount; i++) {
            Cell cell = row.getCell(i);
            String expectedHeader = "VAR" + (i + 1);

            if (cell == null || cell.getCellType() != CellType.STRING) {
                ExtraMethods.sendMessage("Column " + (i + 1) + " does not contain a valid header.");
                return false;
            }

            String currentHeader = cell.getStringCellValue().trim();
            if (!expectedHeader.equals(currentHeader)) {
                ExtraMethods.sendMessage("Expected '" + expectedHeader + "' but found '" + currentHeader + "'.");
                return false;
            }
        }
        return true;
    }

    public boolean createExcelFile(String filePath) {
        boolean status = false;
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");

            Row headerRow = sheet.createRow(0);

            String[] headers = {"VAR1", "VAR2", "VAR3", "VAR4", "VAR5"};

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            for (int i = 0; i < headers.length; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
            status = true;
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return status;
    }

    public Map<String, String> readFirstRowFromData(String pathExcel) {
        Map<String, String> dataMap = new HashMap<>();
        try (FileInputStream fis = new FileInputStream(pathExcel)) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            Row dataRow = sheet.getRow(1);
            if (headerRow == null || dataRow == null) {
                throw new IllegalArgumentException("File excel not contains new rows.");
            }
            for (Cell headerCell : headerRow) {
                String variableName = headerCell.getStringCellValue();
                Cell dataCell = dataRow.getCell(headerCell.getColumnIndex());
                if (dataCell != null) {
                    String cellValue = dataCell.toString();
                    dataMap.put("%" + variableName + "%", cellValue);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return dataMap;
    }

    public List<Map<String, String>> readRowsFromData(File data) {
        List<Map<String, String>> listDataMap = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(data)) {
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();
            Iterator<Row> rowIterator = sheet.iterator();
            rowIterator.next();
            while (rowIterator.hasNext()) {
                Row currentRow = rowIterator.next();
                Map<String, String> dataMap = new HashMap<>();
                for (int colIndex = 0; colIndex < columnCount; colIndex++) {
                    String columnName = headerRow.getCell(colIndex).getStringCellValue();
                    String cellValue = ExtraMethods.getCellValueAsString(currentRow.getCell(colIndex));
                    dataMap.put("%" + columnName + "%", cellValue);
                }
                listDataMap.add(dataMap);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            System.gc();
        }
        return listDataMap;
    }

    public String generatePreviewContent(String pathtemplate, Map<String, String> dataMap) {
        StringBuilder contentBuilder = new StringBuilder();
        try (FileInputStream fis = new FileInputStream(getTemplate(pathtemplate)); XWPFDocument document = new XWPFDocument(fis)) {

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                String text = paragraph.getText();
                for (Map.Entry<String, String> entry : dataMap.entrySet()) {
                    text = text.replace(entry.getKey(), entry.getValue());
                }
                contentBuilder.append(text).append("\n");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return contentBuilder.toString();
    }

    public XWPFDocument generateDocument(File fileTemplate, Map<String, String> dataMap) {
        XWPFDocument document = null;
        try {
            FileInputStream fis = new FileInputStream(fileTemplate);
            XWPFDocument template = new XWPFDocument(fis);
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            template.write(baos);
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
            document = new XWPFDocument(bais);

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        for (Map.Entry<String, String> entry : dataMap.entrySet()) {
                            text = text.replace(entry.getKey(), entry.getValue());
                        }
                        run.setText(text, 0);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return document;
    }
}
