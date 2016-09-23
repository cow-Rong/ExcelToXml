

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.output.Format;
import org.jdom.output.XMLOutputter;

/**
 * 操作Excel表格的功能类
 * 说明：
 * excel文件为 2003-2007 xls文件
 */
public class ExcelReaderXMLWriter2 {

    public static void main(String[] args) {
        try {
            InputStream stream = new FileInputStream("d:\\android_apk\\testdata.xls");
            File f = new File("d:\\android_apk\\testdata.xml");// 新建个file对象把解析之后得到的xml存入改文件中
            writerXML(stream, f);// 将数据以xml形式写入文本
        } catch (FileNotFoundException e) {
            System.out.println("未找到指定路径的文件!");
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

    private static void writerXML(InputStream stream, File f)
            throws IOException {
        FileOutputStream fo = new FileOutputStream(f);// 得到输入流
        Document doc = readExcell(stream);// 读取EXCEL函数
        Format format = Format.getCompactFormat().setEncoding("gb2312")
                .setIndent("");
        //GBK ; gb2312
        XMLOutputter XMLOut = new XMLOutputter(format);// 在元素后换行，每一层元素缩排四格
        XMLOut.output(doc, fo);
        fo.close();
    }

    private static Document readExcell(InputStream stream) {
    	System.out.println("Document");
        Element root = new Element("contacts");
        Document doc = new Document(root);
        try {
            HSSFWorkbook wb = new HSSFWorkbook(stream);
            System.out.println("temp");
            int WbLength = wb.getNumberOfSheets();
            System.out.println("WbLength" + WbLength);
            for (int i = 0; i < WbLength; i++) {
                HSSFSheet shee = wb.getSheetAt(i);
                int length = shee.getLastRowNum();
                System.out.println("行数：" + length);
                for (int j = 1; j <= length; j++) {
                    HSSFRow row = shee.getRow(j);
                    if (row == null) {
                        continue;
                    }
                    int cellNum = row.getPhysicalNumberOfCells();// 获取一行中最后一个单元格的位置
                    System.out.println("列数cellNum：" + cellNum);
                    // int cellNum =16;
                    Element e = null;
                    e = new Element("contact");
                    // Element[] es = new Element[16];
                    for (int k = 0; k < cellNum; k++) {
                        HSSFCell cell = row.getCell((short) k);
                        String temp = get(k);
                        System.out.print(k+" "+temp+":");
                        Element item = new Element(temp);
                        if (cell == null) {
                            item.setText("");
                            e.addContent(item);
                            cellNum++;//如果存在空列，那么cellNum增加1，这一步很重要。
                            continue;
                        }

                        else {
                            String cellvalue = "";
                            switch (cell.getCellType()) {
                            // 如果当前Cell的Type为NUMERIC
                            case HSSFCell.CELL_TYPE_NUMERIC:
                            case HSSFCell.CELL_TYPE_FORMULA: {
                                // 判断当前的cell是否为Date
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    // 如果是Date类型则，转化为Data格式

                                    // 方法1：这样子的data格式是带时分秒的：2011-10-12 0:00:00
                                    // cellvalue =
                                    cell.getDateCellValue().toLocaleString();

                                    // 方法2：这样子的data格式是不带带时分秒的：2011-10-12
                                    Date date = cell.getDateCellValue();
                                    SimpleDateFormat sdf = new SimpleDateFormat(
                                            "yyyy-MM-dd");
                                    cellvalue = sdf.format(date);
                                    item.setText(cellvalue);

                                }
                                // 如果是纯数字
                                else {
                                    // 取得当前Cell的数值
                                    cellvalue = String.valueOf((int)cell.getNumericCellValue());
                                    item.setText(cellvalue);
                                }
                                break;
                            }
                            // 如果当前Cell的Type为STRIN
                            case HSSFCell.CELL_TYPE_STRING:
                                // 取得当前的Cell字符串
                                cellvalue = cell.getRichStringCellValue()
                                        .getString();
                                item.setText(cellvalue);
                                break;
                            // 默认的Cell值
                            default:
                                cellvalue = " ";
                                item.setText(cellvalue);
                            }
                            e.addContent(item);
                            System.out.println(cellvalue);
                        }
                    }
                    root.addContent(e);

                }

            }
        } catch (Exception e) {
        	e.printStackTrace();
        }
        try {
            stream.close();
        } catch (IOException e) {

        }
        return doc;
    }

    /**
     * 按照自身要求添加节点
     * @param k
     * @return
     */
    private static String get(int k) {
        String test = "";
        switch (k) {
        case 0:
            test = "id";
            break;
        case 1:
            test = "name";
            break;
        case 2:
            test = "phone";
            break;
        case 3:
        	test = "email";
        	break;
        default:
        }
        return test;

    }

}