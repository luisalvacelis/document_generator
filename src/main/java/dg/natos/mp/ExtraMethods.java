package dg.natos.mp;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;
import javax.swing.DefaultComboBoxModel;
import javax.swing.ImageIcon;
import javax.swing.JComboBox;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;

public class ExtraMethods {

    public static void sendMessage(String msg) {
        System.out.println("[" + getCurrentDateFormat("dd-MM-yyyy hh:mm:ss aa") + "] " + msg);
    }

    public static String getCurrentDateFormat(String format) {
        SimpleDateFormat sdf = new SimpleDateFormat(format);
        return sdf.format(Calendar.getInstance().getTime());
    }

    public static void loadListCB(JComboBox<String> cb, List<String> list) {
        DefaultComboBoxModel<String> model = new DefaultComboBoxModel<>();
        model.addElement("---- Seleccionar ----");

        if (!list.isEmpty()) {
            for (String item : list) {
                model.addElement(item);
            }
        }

        cb.setModel(model);
    }

    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING -> {
                return cell.getStringCellValue();
            }
            case NUMERIC -> {
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            }
            case BOOLEAN -> {
                return String.valueOf(cell.getBooleanCellValue());
            }
            case FORMULA -> {
                return cell.getCellFormula();
            }
            default -> {
                return "";
            }
        }
    }

    public static void sendErrorMessage(String text) {
        JOptionPane.showMessageDialog(null, text, "Error", 0, new ImageIcon(new ExtraMethods().getClass().getResource("/images/error32x32.png")));
    }

    public static void sendSuccessfullyMessage(String text) {
        JOptionPane.showMessageDialog(null, text, "Éxito", 0, new ImageIcon(new ExtraMethods().getClass().getResource("/images/exito32x32.png")));
    }

    public static int sendConfirmMessage(String msg) {
        return JOptionPane.showConfirmDialog(null, msg, "Confirmar", 0, 0, new ImageIcon(new ExtraMethods().getClass().getResource("/images/archivos32x32.png")));
    }

    public static int sendYesOrNotMessage(String msg) {
        return JOptionPane.showConfirmDialog(null, msg, "Confirmar", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, new ImageIcon(new ExtraMethods().getClass().getResource("/images/archivos32x32.png")));
    }

    public static void sendLogMessage(String text) {
        System.out.println("[" + getCurrentDateFormat("dd/MM/yyyy hh:mm:ss aa") + "] " + text);
    }
}
