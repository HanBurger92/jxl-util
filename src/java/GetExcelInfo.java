import java.io.*;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;

public class GetExcelInfo {

    public static void main(String[] args){
        // 创建此类的对象，以使用此类的方法
        // create the object of the class to use its method
        GetExcelInfo obj = new GetExcelInfo();
s
        // 利用此功能来创建一个窗口，给用户提供选择文件的GUI视图
        // provide GUI View to let user select file
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel File", "xls");
        fileChooser.setFileFilter(filter);
        int returnVal = fileChooser.showOpenDialog(null);
        if(returnVal == JFileChooser.APPROVE_OPTION){
            System.out.println("你选择了打开文件：" + fileChooser.getSelectedFile().getAbsolutePath());
        }

        // 读取文件
        // read file
        File xlsFile = new File(fileChooser.getSelectedFile().getAbsolutePath());
        obj. readExcel(xlsFile);

    }

    public void readExcel(File file){
        try {
            // 创建输入流来读取文件
            // create stream to read file
            FileInputStream fis = new FileInputStream(file);
            String cellInfo = "";

            // jxl提供的workbook类
            // workbook class provided by jxl
            Workbook workbook = Workbook.getWorkbook(fis);
            // 拿到Excel的页签数量
            // get the sheet number of excel
            int sheetSize = workbook.getNumberOfSheets();
            for (int i = 0; i<sheetSize; i++){
                Sheet sheet = workbook.getSheet(i);

                // 每一行
                for(int j=0; j<sheet.getRows(); j++){
                    // 每一列
                    for (int z=0; z<sheet.getColumns(); z++){
                        cellInfo = sheet.getCell(z,j).getContents();
                        System.out.println(cellInfo);
                    }
                }
            }

        }catch (Exception e){
            e.printStackTrace();
        }
    }

    public void writeFile(File file){
        try {
            FileOutputStream fop = new FileOutputStream(file);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
