import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.POIXMLDocument;
import org.apache.xmlbeans.XmlCursor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

/**
 * Created by huanyue6660 on 18/02/09.
 */
public class POITest {
    private static final String docxReadPath = "F:\\h\\poitest\\poitest.docx";
    private static final String docxWritePath = "F:\\h\\poitest\\poiout.docx";
    /**
     * 读取文件
     * @param srcPath
     * @return XWPFDocument
     */
    private static XWPFDocument read_file(String srcPath)
    {
        String[] sp = srcPath.split("\\.");
        if ((sp.length > 0) && sp[sp.length - 1].equalsIgnoreCase("docx"))
        {
            try {
                OPCPackage pack = POIXMLDocument.openPackage(srcPath);
                XWPFDocument doc = new XWPFDocument(pack);
                return doc;
            } catch (IOException e) {
                System.out.println("读取文件出错！");
                e.printStackTrace();
                return null;
            }
        }
        return null;
    }

    /**
     * 插入文字与表格
     * @param document
     * @return XWPFDocument
     */
    private static XWPFDocument insertParagraph(XWPFDocument document){
        Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
        int ind = 1;
        //获取段落位置
        while (itPara.hasNext()) {
            XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
            XWPFRun xrun = paragraph.createRun();
            xrun.setText("这是第"+ ind++ +"个段落！");

            //插入表格
            XmlCursor cursor = paragraph.getCTP().newCursor();
            XWPFTable tb = document.insertNewTbl(cursor);
            //行
            XWPFTableRow row = tb.getRow(0);
            row.addNewTableCell();
            row.getCell(0).setText("0");
            row.getCell(1).setText("1");
        }
        return document;
    }

    /**
     * 写入文件到磁盘
     * @param document
     * @param path
     */
    private static void writeDoc(XWPFDocument document,String path){
        FileOutputStream fOut = null;
        try {
            fOut = new FileOutputStream(path);
            document.write(fOut);
            fOut.close();
            fOut = null;
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 遍历段落内容
     * @param document
     * @return XWPFDocument
     */
    private static XWPFDocument readPar(XWPFDocument document){
        Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
        while (itPara.hasNext()) {
            XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
            //run表示相同区域属性相同的字符，结果以‘，’分隔；
            List<XWPFRun> runs = paragraph.getRuns();
            for (int i = 0; i < runs.size(); i++)
            {
                String oneparaString = runs.get(i).getText(runs.get(i).getTextPosition());
                System.out.println(oneparaString);
            }
        }
        return document;
    }

    /**
     * 遍历所有表格的内容
     * @param document
     */
    private static void readTableContent(XWPFDocument document){
        Iterator<XWPFTable> itTable = document.getTablesIterator();
        int ind = 0;
        while (itTable.hasNext()){
            ind++;
            XWPFTable table = (XWPFTable) itTable.next();
            //行
            int rcount = table.getNumberOfRows();
            for (int i = 0; i < rcount; i++){
                XWPFTableRow row = table.getRow(i);
                //列
                List<XWPFTableCell> cells = row.getTableCells();
                int len = cells.size();
                for(int j = 0;j < len;j++){
                    XWPFTableCell xc = cells.get(j);
                    String sc = xc.getText();
                    System.out.println("第"+ ind +"个表格，第"+ (i+1) +"行，第"+ (j+1) +"列：" +sc);
                }
            }
        }
    }

    public static void main(String[] args){
        XWPFDocument document = read_file(docxReadPath);
        document = insertParagraph(document);
        readPar(document);
        //重新写入，否则表格内容读取不到。
        writeDoc(document, docxWritePath);
        document = read_file(docxWritePath);
        readTableContent(document);
    }
}
















