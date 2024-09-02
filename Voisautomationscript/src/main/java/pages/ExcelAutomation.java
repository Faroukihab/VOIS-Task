package pages;
import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;
import org.sikuli.script.App;

public class ExcelAutomation {
    public static void main(String[] args) {
        try {
            Screen screen = new Screen();

            App excel = App.open("excel.exe");

            Pattern excelIcon = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\src\\main\\java\\Images\\excel.png");
            screen.wait(excelIcon, 10);

            
            Pattern fileMenu = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\src\\main\\java\\Images\\file.png");
            screen.click(fileMenu);
            screen.type("O", org.sikuli.script.Key.ALT); // Access the Open dialog
            screen.type("C:\\Users\\fihab\\eclipse-workspace\\updatethedate\\src\\data\\TaskData.xlsx" + org.sikuli.script.Key.ENTER); // Update with your file path

          
            screen.wait(2.0);

            // Sort by Join Date
     
            Pattern cellA1 = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\src\\main\\java\\Images\\a1.png");
            screen.click(cellA1);

            screen.type(org.sikuli.script.Key.F5);
            screen.type("A1" + org.sikuli.script.Key.ENTER);
            screen.type(org.sikuli.script.Key.ESC);  // Exit Go To dialog

            // Go to the Data tab (capture the Data tab screenshot)
            Pattern dataTab = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\\\src\\\\main\\\\java\\data_tab.png");
            screen.click(dataTab);

            // Click on the Sort icon (capture the Sort icon screenshot)
            Pattern sortIcon = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\updatethedate\\src\\Images\\sort.png");
            screen.click(sortIcon);

            // Wait for the Sort dialog to appear (capture the Sort dialog screenshot)
            Pattern sortDialog = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\\\src\\\\main\\\\java\\sort_dialog_box.png");
            screen.wait(sortDialog, 10);

            // Click on the dropdown to select the column (capture the dropdown screenshot)
            Pattern sortByDropdown = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\updatethedate\\src\\Images\\col1_joindate.png");
            screen.click(sortByDropdown);
        
            Pattern des = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\src\\main\\java\\des.png");
            screen.click(des);
            // Type "Join Date" to select it
            screen.type("Join Date" + org.sikuli.script.Key.ENTER);

            // Confirm the sort (capture the OK button screenshot)
            Pattern okButton = new Pattern("C:\\Users\\fihab\\eclipse-workspace\\Voisautomationscript\\src\\main\\java\\ok_button.png");
            screen.click(okButton);

            //  Remove Duplicates by Name
            // Click on the Remove Duplicates icon (capture the Remove Duplicates icon screenshot)
           Pattern removeDuplicates = new Pattern("path_to_screenshot/remove_duplicates.png");
            screen.click(removeDuplicates);

            // Wait for the Remove Duplicates dialog to appear (capture the dialog screenshot)
            Pattern removeDuplicatesDialog = new Pattern("path_to_screenshot/remove_duplicates_dialog.png");
            screen.wait(removeDuplicatesDialog, 10);

           
            // Assuming you have a screenshot for the "Name" column checkbox
            Pattern nameColumn = new Pattern("path_to_screenshot/name_column.png");
            screen.click(nameColumn);

       
            screen.click(okButton);

            // 7. Save the changes
            screen.type("S", org.sikuli.script.Key.CTRL);

            // 8. Close Excel
            excel.close();

            System.out.println("Excel file processed successfully!");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
