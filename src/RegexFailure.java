import java.io.*;
import java.nio.channels.FileChannel;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class RegexFailure
{
   private static final String CONFIGFILE = "E:/SimpleIndexBackup/SimpleIndexSchedulerConfig.properties";


   public static void main(String[] args) throws Exception 
   {
      // TODO Auto-generated method stub
         List<String> l = readInputFile();
         //processExcelFile(l);
      //findDuplicates();
      final File folder = new File("E:/esg/ptpotc/k2m_import/lot");
      //listFilesForFolder(folder);
      comparefolderfileswithexcelentries(l);
   }
   
   public static void comparefolderfileswithexcelentries(List<String> l) throws Exception{
      Workbook wb = null;
      FileInputStream fi = null;
      File file = null;
      try 
      {
         file = new File("E:/dev/sanjeev/ScanOutput6-11_8-2.xlsx");
          fi = new FileInputStream(file);
         // long starttime;
          wb = WorkbookFactory.create(fi);
         System.out.println("Begin processing for loop");
         int counter=0;
         for (Sheet sheet : wb)
         {
            Row row = null;
            String newCellValue = null;
            Cell newFileName = null;
            Cell dupFlag = null;
            row = sheet.getRow(0);
            newCellValue = "uniqueFiles";
            newFileName = row.createCell(9);
            newFileName.setCellValue(newCellValue);
            for(String s : l){
               for(int k=1; k<=2485;k++){
                  row = sheet.getRow(k);
                  if(row!=null){
                     if (row.getCell(7) != null && row.getCell(7).getStringCellValue().equalsIgnoreCase(s))
                     {
                        if(row.getCell(9) == null){
                           System.out.println(s);
                           System.out.println(++counter);
                           newFileName = row.createCell(9);
                           newFileName.setCellValue(2);
                           break;
                        }
                     }
               }
            }
            }
         }
         System.out.println(counter);
      }catch (Exception e) {
         e.printStackTrace();
      }finally{
         FileOutputStream foo = null;
         try{
            foo = new FileOutputStream(file);
            wb.write(foo);
         }catch(Exception e){
            
         }finally{
            foo.close();
         
            if(wb != null)
               wb.close();
            if(fi !=null){
               fi.close();
         }
         }
      }
   }
   
   public static void listFilesForFolder(final File folder) throws FileNotFoundException {
      //File file = new File("E:/esg/ptpotc/filelist.txt");
      //FileOutputStream f = new FileOutputStream(file);
      try (PrintStream out = new PrintStream(new FileOutputStream("E:/SimpleIndexBackup/programtolistfilesinlot/filelist.txt"))) {
         for (final File fileEntry : folder.listFiles()) {
            System.out.println(fileEntry.getName());
            out.println(fileEntry.getName());
         }
     }catch (Exception e) {
      // TODO: handle exception
   }
     
  }

   public static void findDuplicates() throws IOException{
      
      Workbook wb = null;
      FileInputStream fi = null;
      File file = null;
      try 
      {
         file = new File("E:/dev/sanjeev/ScanOutput6-11_8-2.xlsx");
          fi = new FileInputStream(file);
         // long starttime;
          wb = WorkbookFactory.create(fi);
         System.out.println("Begin processing for loop");
int counter=0;
         for (Sheet sheet : wb)
         {
            Row row = null;
            String newCellValue = null;
            Cell newFileName = null;
            Cell dupFlag = null;
            row = sheet.getRow(0);
            newCellValue = "Duplicates";
            newFileName = row.createCell(8);
            newFileName.setCellValue(newCellValue);
            for(int i=1; i<1440;i++){
               Row row1 = sheet.getRow(i);
               for(int k=i+1; k<=1440;k++){
                  row = sheet.getRow(k);
                  if(row!=null){
                     if (row1.getCell(7).getStringCellValue().equalsIgnoreCase(row.getCell(7).getStringCellValue()))
                     {
                        if(row.getCell(9) == null){
                           System.out.println(row.getCell(7).getStringCellValue());
                           System.out.println(++counter);
                           newFileName = row.createCell(8);
                           newFileName.setCellValue(1);
                           //newFileName = row1.createCell(8);
                           //newFileName.setCellValue(1);
                        }
                     }
               }
            }
            }
         }
         
      }catch (Exception e) {
         e.printStackTrace();
      }finally{
         FileOutputStream foo = null;
         try{
            foo = new FileOutputStream(file);
            wb.write(foo);
         }catch(Exception e){
            
         }finally{
            foo.close();
         
            if(wb != null)
               wb.close();
            if(fi !=null){
               fi.close();
         }
         }
      }
}
   public static void processExcelFile(List<String> l) throws IOException{
      
         Workbook wb = null;
         FileInputStream fi = null;
         File file = null;
         try 
         {
             file = new File("E:/SimpleIndexBackup/ScanOpsthru060818.xlsx");
             fi = new FileInputStream(file);
            // long starttime;
             wb = WorkbookFactory.create(fi);
            System.out.println("Begin processing for loop");
            int counter=0;
            for (Sheet sheet : wb)
            {
               Row row = null;
               String newCellValue = null;
               Cell newFileName = null;
               Cell dupFlag = null;
               row = sheet.getRow(0);
               newCellValue = "Regex Failures";
               newFileName = row.createCell(9);
               newFileName.setCellValue(newCellValue);
               for(String s : l){
                  for(int k=1; k<=131232;k++){
                     row = sheet.getRow(k);
                     if(row!=null){
                        if (row.getCell(7) != null && row.getCell(7).getStringCellValue().equalsIgnoreCase(s))
                        {
                           if(row.getCell(9) == null){
                              System.out.println(s);
                              System.out.println(++counter);
                              newFileName = row.createCell(9);
                              newFileName.setCellValue(2);
                           }
                        }
                  }
               }
               }
            }
            
         }catch (Exception e) {
            e.printStackTrace();
         }finally{
            FileOutputStream foo = null;
            try{
               foo = new FileOutputStream(file);
               wb.write(foo);
            }catch(Exception e){
               
            }finally{
               foo.close();
            
               if(wb != null)
                  wb.close();
               if(fi !=null){
                  fi.close();
            }
            }
         }
   }
   public static List readInputFile() throws IOException{
      String s = null;
      List l = new ArrayList();
      try(BufferedReader br = new BufferedReader(new FileReader("E:/filelist.txt"))){
      
      while((s = br.readLine())!=null){
         //l.add(s.substring(0, s.lastIndexOf(".pdf")+4));
         l.add(s);
      }
      }catch(Exception e){
         e.printStackTrace();
      }
      System.out.println(l.size()+"  " +l.get(l.size()-1));
      return l;
   }
}
