import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class BatchProcess implements Runnable
{
   private static final String CONFIGFILE = "E:/SimpleIndexBackup/SimpleIndexSchedulerConfig.properties";

   public static void main(String[] args)
   {
      // processAllFiles();
      BatchProcess sc = new BatchProcess();
      Thread t = new Thread(sc, "Scheduler");
      t.start();
   }

   @Override
   public void run()
   {
      Properties p = new Properties();

      try (FileInputStream f = new FileInputStream(new File(CONFIGFILE));
            BufferedInputStream b = new BufferedInputStream(f))
      {
         p.load(b);

         String interval = p.getProperty("timeintervalinseconds");
         String runonce = p.getProperty("runonce");
         b.close();
         f.close();
         if(Boolean.valueOf(runonce)){
            processSpecificFiles();
         }else{
            for (;;)
            {
               System.out.println("Scheduler thread begin to execute");
               try
               {
                  //cleanupTargetFolder();
                  processSpecificFiles();

                  Thread.sleep(Integer.valueOf(interval) * 1000);
               } catch (InterruptedException e)
               {
                  // TODO Auto-generated catch block
                  e.printStackTrace();
               }

            }
         }
      } catch (IOException e1)
      {
         // TODO Auto-generated catch block
         e1.printStackTrace();
      }
   }

   @SuppressWarnings("resource")
   private void processSpecificFiles()
   {
      Properties p = new Properties();

      try (FileInputStream f = new FileInputStream(new File(CONFIGFILE));
            BufferedInputStream b = new BufferedInputStream(f))
      {
         p.load(b);
         int beginLineNumber = Integer.parseInt(p.getProperty("beginLineNumber"));
         int numberOfLines = Integer.parseInt(p.getProperty("numberOfLines"));
         String inputExcelFilePath = p.getProperty("inputExcelFilePath");
         String originalFilePath = p.getProperty("originalFilePath");
         String newFilePath = p.getProperty("newFilePath");
         String missingFilePath = p.getProperty("missingFilePath");

         File file = new File(inputExcelFilePath);
         FileInputStream fi = new FileInputStream(file);
         FileOutputStream fo = null;
         // long starttime;
         FileChannel in = null, out = null;
         List<String> missingFiles = new ArrayList();
         Workbook wb = WorkbookFactory.create(fi);
         System.out.println("Begin processing for loop");
int counter=0;
         for (Sheet sheet : wb)
         {
            int j = beginLineNumber;
            if (beginLineNumber <= sheet.getPhysicalNumberOfRows())
            {
               try
               {
                  Row row = null;
                  String newCellValue = null;
                  Cell newFileName = null;
                  Cell dupFlag = null;
                  if (beginLineNumber <= 1)
                  {
                     row = sheet.getRow(0);
                     newCellValue = "New Filename";
                     newFileName = row.createCell(7);
                     newFileName.setCellValue(newCellValue);
                     newFileName = row.createCell(8);
                     newFileName.setCellValue("Duplicate?");
                  }
                  
                  CellStyle style = wb.createCellStyle();
                  style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
                  style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//                  Font font = wb.createFont();
//                       font.setColor(IndexedColors.RED.getIndex());
//                       style.setFont(font);

                  // starttime = System.nanoTime();
                  System.out.println(
                        "Lines processed from :" + beginLineNumber + " to " + (beginLineNumber + numberOfLines));
                  // System.out.println("Row count time:
                  // "+(starttime-System.nanoTime()));
                  //Set s = new HashSet();
                  Map<String, Integer> m = new HashMap();
                  for (; j < beginLineNumber + numberOfLines; j++)
                  {
                     row = sheet.getRow(j);
                     if(row!=null){
                        if (row.getCell(7) == null)
                        {
                           // long starttime2 = System.nanoTime();
                           System.out.println(j);
                           String catalogNum = row.getCell(2).getStringCellValue();
                           String lotNum = String.valueOf(row.getCell(3).getNumericCellValue());
                           lotNum=lotNum.substring(0, lotNum.length()-2);
                           String boxNum = row.getCell(4).getStringCellValue();
                           // Lot_TTJ_Cat_101-10001_Box_11-005_ScanU_$.docx
                           newCellValue = "Lot_" + lotNum + "_Cat_" + catalogNum + "_Box_" + boxNum + "_ScanU_EIP.pdf";
                           // System.out.println(newCellValue);
                          // s.add(newCellValue);
                           System.out.println(newCellValue);
                           ++counter;
                           newFileName = row.createCell(7);
                           newFileName.setCellValue(newCellValue);
//                           if(m.containsKey(newCellValue)){
//                              //System.out.println("Duplicate :"+newCellValue);
//                              row.getCell(0).setCellStyle(style);
//                              row.getCell(1).setCellStyle(style);
//                              row.getCell(2).setCellStyle(style);
//                              row.getCell(3).setCellStyle(style);
//                              row.getCell(4).setCellStyle(style);
//                              row.getCell(5).setCellStyle(style);
//                              row.getCell(6).setCellStyle(style);
//                              row.getCell(7).setCellStyle(style);
//                              dupFlag = row.createCell(8);
//                              dupFlag.setCellValue(1);
//                           }else{
//                              m.put(newCellValue, j);
//                              newFileName = row.createCell(8);
//                              newFileName.setCellValue(0);
//                           }
//                           System.out.println("Writing file: " + newCellValue + " index: " + j);
//                           try
//                           {
//                              in = new RandomAccessFile(
//                                    originalFilePath + "/" + row.getCell(6).getStringCellValue().substring(1), "r")
//                                          .getChannel();
//                              out = new RandomAccessFile(newFilePath + "/" + newCellValue, "rw").getChannel();
//                              out.transferFrom(in, 0, Long.MAX_VALUE);
//   
//                           } catch (FileNotFoundException e)
//                           {
//                              System.out.println(
//                                    "File not found: " + row.getCell(4).getStringCellValue() + "--->" + newCellValue);
//                              missingFiles.add(row.getCell(4).getStringCellValue() + "--->" + newCellValue + "\n");
//                           }
                        } 
//                        else
//                        {
//                           break;
//                        }
                     } else
                     {
                        System.out.println("No further rows to process in the sheet");
                        break;
                     }
                  }
                  System.out.println("Number of records processed :"+counter);
                  System.out.println("Number of unique rows "+m.size());
                  System.out.println("Successfully completed batch process");
                  m.clear();
                  m=null;
               } catch (Exception e)
               {
                  System.out.println("Exception during batch process");
                  e.printStackTrace();
               } finally
               {
                  if (fi != null)
                  {
                     fi.close();
                  }
                  if (fo != null)
                  {
                     fo.close();
                  }
                  if (in != null)
                  {
                     in.close();
                  }
                  if (out != null)
                  {
                     out.close();
                  }
                  FileOutputStream o = new FileOutputStream(new File(CONFIGFILE));
                  BufferedOutputStream ob = new BufferedOutputStream(o);
                  p.setProperty("beginLineNumber", String.valueOf(j));
                  p.store(ob, "");
                  ob.close();
                  o.close();
                  try(FileOutputStream foo = new FileOutputStream(file);
                        FileOutputStream mfo = new FileOutputStream(new File(missingFilePath + beginLineNumber+".txt"));){
                  
                  wb.write(foo);
                  
                  System.out.println(
                        "Number of Missing files: " + missingFiles.size() + " in " + missingFilePath + beginLineNumber+".txt");
                  mfo.write(("Number of Missing files: " + missingFiles.size() + "\n").getBytes());
                  mfo.write(("Below is the list of files missing" + "\n").getBytes());
                  mfo.write(("Original Filename ----> New Filename" + "\n").getBytes());
                  mfo.write(("-------------------------------------------------------------" + "\n").getBytes());

                  for (String mf : missingFiles)
                  {
                     mfo.write(mf.getBytes());
                  }
                  }catch(Exception e){
                     e.printStackTrace();
                  }finally{
                     if(wb != null)
                        wb.close();
                  }
               }
            } else
            {
               System.out.println("No rows to process");
            }
         }
      } catch (Exception e)
      {
         System.out.println("Exception during batch process");
         e.printStackTrace();
      }
   }

   private void cleanupTargetFolder(){
      Properties p = new Properties();

      try (FileInputStream f = new FileInputStream(new File(CONFIGFILE));
            BufferedInputStream b = new BufferedInputStream(f))
      {
         p.load(b);
         String newFilePath = p.getProperty("newFilePath");
         File dir = new File(newFilePath);
         if(dir.isDirectory()){
            File[] files = dir.listFiles();
            for(File fi: files){
               if(fi.exists()){
                  System.out.println("Deleting file "+fi.getName());
                  fi.delete();
               }
            }
         }
      }catch (Exception e)
      {
         System.out.println("Exception during batch process");
         e.printStackTrace();
      } 
      
   }
   private static void processAllFiles()
   {
      try
      {
         File f = new File("E:/SimpleIndexBackup/ScanOutputOps_113017.xlsx");
         FileInputStream fi = new FileInputStream(f);
         FileOutputStream fo = null;
         // long starttime;
         FileChannel in = null, out = null;
         List<String> missingFiles = new ArrayList();
         Workbook wb = WorkbookFactory.create(fi);
         String originalFilePath = "E:/SimpleIndexBackup/pdffiles";
         String newFilePath = "E:/esg/ptpotc/k2m_import/lot";
         System.out.println("Begin processing for loop");
         for (Sheet sheet : wb)
         {
            try
            {
               Row row = sheet.getRow(0);
               String newCellValue = "New Filename";
               Cell newFileName = row.createCell(7);
               newFileName.setCellValue(newCellValue);
               // starttime = System.nanoTime();
               System.out.println("Number of rows :" + sheet.getPhysicalNumberOfRows());
               // System.out.println("Row count time:
               // "+(starttime-System.nanoTime()));
               for (int j = 1; j < sheet.getPhysicalNumberOfRows(); j++)
               {
                  row = sheet.getRow(j);
                  if (row.getCell(2) != null)
                  {
                     // long starttime2 = System.nanoTime();
                     String catalogNum = row.getCell(2).getStringCellValue();
                     String lotNum = row.getCell(3).getStringCellValue();
                     String boxNum = row.getCell(4).getStringCellValue();
                     // Lot_TTJ_Cat_101-10001_Box_11-005_ScanU_$.docx
                     newCellValue = "Lot_" + lotNum + "_Cat_" + catalogNum + "_Box_" + boxNum + "_ScanU_EIP.pdf";
                     // System.out.println(newCellValue);

                     newFileName = row.createCell(7);
                     newFileName.setCellValue(newCellValue);
                     // System.out.println("Time taken to create cell :
                     // "+newCellValue+(starttime2-System.nanoTime()));
                     // long starttime3 = System.nanoTime();
                     // fi = new FileInputStream(new
                     // File(originalFilePath+"/"+row.getCell(6).getStringCellValue().substring(1)));
                     // fo = new FileOutputStream(new
                     // File(newFilePath+"/"+newCellValue));
                     System.out.println("Writing file: " + newCellValue + " index: " + j);
                     try
                     {
                        in = new RandomAccessFile(
                              originalFilePath + "/" + row.getCell(6).getStringCellValue().substring(1), "r")
                                    .getChannel();
                        out = new RandomAccessFile(newFilePath + "/" + newCellValue, "rw").getChannel();
                        out.transferFrom(in, 0, Long.MAX_VALUE);

                     } catch (FileNotFoundException e)
                     {
                        System.out.println(
                              "File not found: " + row.getCell(4).getStringCellValue() + "--->" + newCellValue);
                        missingFiles.add(row.getCell(4).getStringCellValue() + "--->" + newCellValue + "\n");
                     }

                     /*
                      * int c; while ((c = fi.read()) != -1) { fo.write(c); }
                      */
                     // System.out.println("Time taken to write :
                     // "+newCellValue+" "+(starttime3-System.nanoTime()));
                  } else
                  {
                     break;
                  }
               }
               fo = new FileOutputStream(f);
               wb.write(fo);
               fo = new FileOutputStream(new File("E:/SimpleIndexBackup/missingFileList.txt"));
               System.out.println("Number of Missing files: " + missingFiles.size());
               fo.write(("Number of Missing files: " + missingFiles.size() + "\n").getBytes());
               fo.write(("Below is the list of files missing" + "\n").getBytes());
               fo.write(("Original Filename ----> New Filename" + "\n").getBytes());
               fo.write(("-------------------------------------------------------------" + "\n").getBytes());

               for (String file : missingFiles)
               {
                  fo.write(file.getBytes());
               }
               System.out.println("Successfully completed batch process");
            } catch (Exception e)
            {
               System.out.println("Exception during batch process");
               e.printStackTrace();
            } finally
            {
               if (fi != null)
               {
                  fi.close();
               }
               if (wb != null)
                  wb.close();
            }
            if (fo != null)
            {
               fo.close();
            }
            if (in != null)
               in.close();
         }
         if (out != null)
         {
            out.close();
         }

      } catch (Exception e)
      {
         System.out.println("Exception during batch process");
         e.printStackTrace();
      }
   }
}
