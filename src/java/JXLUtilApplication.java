import jxl.CellView;
import jxl.Workbook;
import jxl.demo.Write;
import jxl.write.*;
import jxl.write.biff.RowsExceededException;

import java.io.File;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Random;

public class JXLUtilApplication {

    // todo: 增加计算温度最大值和最小值
    public static void main(String[] args) throws IOException, RowsExceededException, WriteException {
        //todo: 此为测试版 正式版要改参数
        //todo: 写个GUI，指引导出Excel到指定地址
        for (int i = 0; i < 2; i++) {
            // 创建Excel
            String separator = File.separator;
            String directory = "D:" + separator + "ExcelFile";
            String name = "test" + i + ".xls";
            File file = new File(directory, name);
            file.createNewFile();

            // 创建工作本
            WritableWorkbook workbook = Workbook.createWorkbook(file);

            // 创建工作表
            WritableSheet writableSheet = workbook.createSheet("温度表", 0);

            // 设置表头名
            String[] titles = {"时间", "温度"};

            // 设置单元格格式
            WritableFont wf = new WritableFont(WritableFont.ARIAL,12, WritableFont.NO_BOLD);
            WritableCellFormat wcfF = new WritableCellFormat(
                    NumberFormats.TEXT); //定义一个单元格样式
            wcfF.setFont(wf);
            CellView cv = new CellView(); //定义一个列显示样式
            cv.setFormat(wcfF);//把定义的单元格格式初始化进去
            cv.setSize(20*265);//设置列宽度（不设置的话是0，不会显示）
            writableSheet.setColumnView(0, cv);//设置工作表中第n列的样式
            writableSheet.setColumnView(1, cv);//设置工作表中第n列的样式

            // 创建单元格
            Label label = null;
            label = new Label(0, 0, "记录报告");
            writableSheet.addCell(label);
            for (int j = 0; j < titles.length; j++) {
                label = new Label(j, 1, titles[j]);
                writableSheet.addCell(label);
            }

            // 拿取随机时间
            //todo: 此为测试版 正式版要改参数
            //todo: 要注意30号和31号的时候第二天月份的改变
            String year = "2017/";
            String month = 1 + "/";
            String day = i + 1 + "";
            Date startDate = randomDate(year + month + day + " 17:00", year + month + day + " 18:00");
            day = i + 2 + "";
            Date endDate = randomDate(year + month + day + " 9:00", year + month + day + " 10:00");
            int z = 2;
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd HH:mm");
            double previous = 4.0;

            // 插入日期和温度
            while (true) {
                if (endDate.after(startDate)) {
                    // 插入日期
                    try {
                        label = new Label(0, z, sdf.format(startDate));
                        writableSheet.addCell(label);
                        double random = randomDouble(previous);
                        previous = random;
                        jxl.write.Number labelN = new jxl.write.Number(1, z, random);
                        writableSheet.addCell(labelN);
                        z++;

                        // 每五分钟一次
                        Calendar calendar = Calendar.getInstance();
                        calendar.setTime(startDate);
                        calendar.add(Calendar.MINUTE, 5);
                        startDate = calendar.getTime();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                } else {
                    break;
                }
            }
            workbook.write();
            workbook.close();
        }

    }


    private static Date randomDate(String begin, String end) {
        try {
            SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd HH:mm");
            Date startDate = format.parse(begin);
            Date endDate = format.parse(end);

            if (startDate.getTime() > endDate.getTime()) return null;

            long date = random(startDate.getTime(), endDate.getTime());

            return new Date(date);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private static long random(long begin, long end) {
        long rtn = begin + (long) (Math.random() * (end - begin));
        if (rtn == begin || rtn == end) {
            return random(begin, end);
        }
        return rtn;
    }

    private static double randomDouble(double previous){

        double temperature = 3.0;
        if(previous > 6.5){
            temperature = Math.random() + 5.5;
            temperature = (double)Math.round(temperature * 10)/10;
            return temperature;
        } else if(previous < 3){
            temperature = Math.random() + 3;
            temperature = (double)Math.round(temperature * 10)/10;
            return temperature;
        }else{
            temperature = Math.random() + 4;
            temperature = (double)Math.round(temperature * 10)/10;
            return temperature;
        }

    }
}

