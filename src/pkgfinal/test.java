/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pkgfinal;

import com.aspose.cells.Button;
import com.aspose.cells.Color;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.PlacementType;
import com.aspose.cells.VbaModule;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

/**
 *
 * @author Admin
 */
public class test {

    public static void main(String[] args) throws Exception {
        String dataDir = "C:\\Users\\Admin\\OneDrive\\Desktop\\ATPM/";

        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.getWorksheets().get(0);

        int moduleIdx = workbook.getVbaProject().getModules().add(sheet);
        VbaModule module = workbook.getVbaProject().getModules().get(moduleIdx);
        module.setCodes("Sub ShowMessage()" + "\r\n"
                + "    MsgBox \"Welcome to Aspose!\"" + "\r\n"
                + "End Sub");

        Button button = (Button) sheet.getShapes().addShape(MsoDrawingType.BUTTON, 2, 0, 2, 0, 28, 80);
        button.setPlacement(PlacementType.FREE_FLOATING);
        button.getFont().setName("Tahoma");
        button.getFont().setBold(true);
        button.getFont().setColor(Color.getBlue());
        button.setText("Aspose");

        button.setMacroName(sheet.getName() + ".ShowMessage");

        workbook.save(dataDir + "Output.xlsm");

        System.out.println("File saved");
    }

}
