
package podL6;

import java.io.File;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import jxl.Workbook;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class XLSWriter {
    
    private File file;
    private WritableWorkbook workbook;
    private WritableSheet sheet;
    
    
    
    // Konstruktor otwierający okno wyboru pliku
    public XLSWriter(){
        file = chooseFileToSave("XLS");
    }
    
    // Konstruktor z podaną ścieżką
    public XLSWriter(String path){
        try {file = new File(path);}
        catch(Exception ex) {ex.printStackTrace(); System.out.println("Błąd w tworzeniu pliku");}
    }
    
    
    
    
    public void writeLabel(int columnIndex, int rowIndex, String label) throws WriteException{
        jxl.write.Label lab = new jxl.write.Label(columnIndex,rowIndex,label);
        sheet.addCell(lab);
    }
    
    public void writeNumber(int columnIndex, int rowIndex, int label) throws WriteException{
        jxl.write.Number lab = new jxl.write.Number(columnIndex,rowIndex,label);
        sheet.addCell(lab);
    }
    
    
    
    
    
    // Metody obsługujące plik
    public void openFile() throws IOException{
        if(file != null) workbook = Workbook.createWorkbook(file);
        else throw new IOException("Plik nie został utworzony");
    }
    
    public void openSheet(String sheetName) throws IOException{
        if(workbook != null) sheet = workbook.createSheet(sheetName, 0);
        else throw new IOException("Nie otworzono pliku");
    }
    
    public void closeFile() throws IOException, WriteException{
        if(file != null) {
            workbook.write();
            workbook.close();
        }
        else throw new IOException("Plik nie został utworzony");
    }
    
    
    
    
    
    /**
     * Metoda otwiera okno z wyborem pliku
     * @param fileExtension Domyślne rozszerzenie pliku (bez kropki)
     * @return Wybrany plik lub NULL w przypadku anulowania
     */
    private static File chooseFileToSave(String fileExtension){
        JFileChooser jfc = new JFileChooser();
        int answer = jfc.showOpenDialog(null);        
        if (answer != jfc.APPROVE_OPTION) return null;
        
        File file = jfc.getSelectedFile();
        String chosenFileName = file.getName();
        
        if(!checkFileExtension(chosenFileName, fileExtension)) file = getFileWithExpectedExtension(file, fileExtension);
        
        if(file.exists() && !confirmFileOverride(file)){
            return null;
        }
        
        return file;
    }
    
    /**
     * Metoda sprawdza, czy rozszerzenie pliku zawarte w jego nazwie zgadza się z porządanym
     * @param fileName Nazwa pliku
     * @param extension 
     * @return 
     */
    private static boolean checkFileExtension(String fileName, String extension){
        if(fileName.contains(".")) return false;        
        String nameAndExtension[] = fileName.split(".");
        if(nameAndExtension.length != 2) return false;
        return extension.equalsIgnoreCase(nameAndExtension[1]);
    }
    
    /**
     * Ustawia rozszerzenie pliku
     * @param file Plik, dla którego mo zostać ustawione rozszerzenie
     * @param extension Rozszerzenie (bez kropki)
     */
    private static File getFileWithExpectedExtension(File file, String extension){
        String filePath = file.toString();
        if(!filePath.contains(".")) filePath += "." + extension.toLowerCase();
        else {
            StringBuilder newFilePath = new StringBuilder();
            int index = 0;
            while(filePath.charAt(index) != '.') newFilePath.append(filePath.charAt(index++));
            filePath = newFilePath.append(".").append(extension.toLowerCase()).toString();
        }
        
        return new File(filePath);
    }
    
    private static boolean confirmFileOverride(File file){
        int response = JOptionPane.showConfirmDialog(null,"Wybrano istniejący plik, czy chcesz go nadpisać?","Plik istnieje", JOptionPane.YES_NO_OPTION,JOptionPane.QUESTION_MESSAGE);
        return response == JOptionPane.YES_OPTION;
    }
}
