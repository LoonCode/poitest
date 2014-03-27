package com.poitest;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Calendar;
import java.util.Date;


public class PoiHelloWorld {
    /**
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {

        String excel6 = "D:\\LGQ-JAVA\\OfficeParse\\WebRoot\\files\\CreateTest.xlsx";
        String excel7 = "D:\\LGQ-JAVA\\OfficeParse\\WebRoot\\files\\CreateTest.xls";

        //readExcel(excel4);
        createExcel(excel6);
        //modifyExcel(excel2);
    }


    static void createExcel(String filePath) throws Exception {

        FileOutputStream fos = new FileOutputStream(new File(filePath));
        Workbook wb = null;
        if (filePath.endsWith("xls"))
            wb = new HSSFWorkbook();
        else
//            OPCPackage pkg = OPCPackage.open(path);
            wb = new XSSFWorkbook();

        Sheet st1 = wb.createSheet("测试Sheet");
        // 注意工作表Sheet名称不能超过 31 个字符
        // 同时也不能包含如下字符:
        // 0x0000
        // 0x0003
        // colon (:)
        // backslash (\)
        // asterisk (*)
        // question mark (?)
        // forward slash (/)
        // opening square bracket ([)
        // closing square bracket (])
        // 可以使用 org.apache.poi.ss.util.WorkbookUtil#createSafeSheetName(String nameProposal)}
        // 以安全的方式创建工作表名称, 非法的字符将被替换成空格 (' ')，例如：
        //String safeName = WorkbookUtil.createSafeSheetName("[O'Brien's sales*?]"); // 将返回 " O'Brien's sales   "
        String safeName = WorkbookUtil.createSafeSheetName("[:测试/安全?Sheet*名\\]");
        Sheet st2 = wb.createSheet(safeName);
        createCell(st2, 15, 5);
        /**
         * 设置工作表打印格式(测试未发现明显变化，具体作用不时)
         */
        Sheet st3 = wb.createSheet("Sheet打印格式测试");
        //设置Excel的打印区域，二种方式，1：使用表达式，2：使用索引值，设置完成后打开Excel按Ctrl+P，打印预览可见只有选定单元格区域将被打印
        wb.setPrintArea(2, "$A$1:$G$10");
        //wb.setPrintArea(2, 0, 6, 0, 9);
        PrintSetup ps = st3.getPrintSetup();
        createCell(st3, 50, 20);
        st3.setAutobreaks(true);
        //设置工作表被选择，打开Excel后会发现该Sheet处于高亮选中状态，打印时默认打印所有选中的工作表Sheet
        st3.setSelected(true);
        ps.setFitHeight((short) 1);
        ps.setFitWidth((short) 1);

        /**
         * 为工作表打印设置页楣和页脚，在页楣和页脚中可以设置左、中、右三处的文本内容，同时在文本内容中支持使用特殊表达式动态插入值，如&P代表当前页数
         * &B xxxxx &B代表中间字符加粗等，还能设置字体信息。HeaderFooter类是org.apache.poi.hssf.usermodel.HeaderFooter抽象类，静态方法
         * 返回的表达式xls和xlsx通用，所以即使是XSSF也可以使用这个类来生成页楣页脚表达式
         * 可以在Excel中按Ctrl+P打印预览中查看页楣页脚效果
         */
        Header header = st3.getHeader();
        Footer footer = st3.getFooter();
        //页楣左边显示时间并添加下划线
        header.setLeft("Hearder_Left " + HeaderFooter.startUnderline() + "Time=" + HeaderFooter.time() + HeaderFooter.endUnderline());
        //页楣中间显示日期并且添加双下划线
        header.setCenter("Hearder_Center " + HeaderFooter.startDoubleUnderline() + "Date=" + HeaderFooter.date() + HeaderFooter.endDoubleUnderline());
        //页楣右边显示当前文件名并加粗
        header.setRight("Hearder_Right " + HeaderFooter.startBold() + "File=" + HeaderFooter.file() + HeaderFooter.endBold());
        //页脚显示当前Sheet工作表名
        footer.setLeft("Footer_Left tab=" + HeaderFooter.tab());
        //页脚中间显示指定字体、字号
        footer.setCenter("Footer_Center" + HeaderFooter.font("黑体", "Italic") +
                HeaderFooter.fontSize((short) 8) + "黑体斜体8号字");
        //页脚显示当前页码和总页数
        footer.setRight("Footer_Right Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        System.out.println("【HeaderFooter】: date=" + HeaderFooter.date() + ", file=" + HeaderFooter.file() + ", endBold=" +
                HeaderFooter.endBold() + ", endDoubleUnderline=" + HeaderFooter.endDoubleUnderline() + ", endUnderline=" +
                HeaderFooter.endUnderline() + ", numPages" + HeaderFooter.numPages() + ", page=" + HeaderFooter.page() + ", startBold=" +
                HeaderFooter.startBold() + ", startDoubleUnderline=" + HeaderFooter.startDoubleUnderline() + ", startUnderline=" +
                HeaderFooter.startUnderline() + ", tab=" + HeaderFooter.tab() + ", time=" + HeaderFooter.time() + ", font=" +
                HeaderFooter.font("黑体", "Italic") + ", fontSize=" + HeaderFooter.fontSize((short) 8));

        /**
         * 在打印中重复行或列(实测未发现明显变化)
         * Repeating rows and columns
         * It's possible to set up repeating rows and columns in your printouts by using the setRepeatingRowsAndColumns()
         * function in the HSSFWorkbook class.
         * This function Contains 5 parameters. The first parameter is the index to the sheet (0 = first sheet). The second
         * and third parameters specify the range for the columns to repreat. To stop the columns from repeating pass in -1
         * as the start and end column. The fourth and fifth parameters specify the range for the rows to repeat. To stop
         * the columns from repeating pass in -1 as the start and end rows.
         */
        //Set the columns to repeat from column 0 to 2 on the first sheet
        wb.setRepeatingRowsAndColumns(0, 0, 2, -1, -1);
        // Set the the repeating rows and columns on the second sheet.
        wb.setRepeatingRowsAndColumns(1, 4, 5, 1, 2);

        /**
         * 在指定索引位置上创建新行，索引由0开始
         */
        Row row = st1.createRow(1);
        //可以设置行高，单位缇(twips)，即20分之1点(1 twips = 1 point / 20)
        row.setHeight((short) 1000);
        //在该行中创建单元格，参数为指定的索引位置由0开始，新创建出来的单元格全部是Cell.CELL_TYPE_BLANK类型的，在调用setCellValue后会根据设置
        //的数据格式变化单元格类型
        Cell cell = row.createCell(1);
        st1.setColumnWidth(1, 5120);
        //CELL_TYPE_STRING:1
        cell.setCellValue("测试单元格内容!(@#&%^)");
        cell.setAsActiveCell();

        //CELL_TYPE_NUMERIC:0
        row.createCell(2).setCellValue(1.2);

        CreationHelper createHelper = wb.getCreationHelper();
        //CELL_TYPE_STRING:1
        row.createCell(3).setCellValue(
                createHelper.createRichTextString("This is a string")
        );

        //CELL_TYPE_BOOLEAN:4
        row.createCell(4).setCellValue(true);

        cell = row.createCell(5);
        //CELL_TYPE_NUMERIC:0
        cell.setCellValue(new Date());
        CellStyle cs = wb.createCellStyle();
        cs.setDataFormat(createHelper.createDataFormat().getFormat("yyyy年MM月dd日 HH时mm分ss秒"));
        cell.setCellStyle(cs);
        //CELL_TYPE_NUMERIC:0
        row.createCell(6).setCellValue(Calendar.getInstance());

        //CELL_TYPE_ERROR:5
        row.createCell(7).setCellType(Cell.CELL_TYPE_ERROR);


        /**
         * 设置单元格文本对齐方式
         */
        Row row2 = st1.createRow(2);
        row2.setHeightInPoints(30);
        createCell(wb, row2, (short) 0, CellStyle.ALIGN_CENTER, CellStyle.VERTICAL_BOTTOM);
        createCell(wb, row2, (short) 1, CellStyle.ALIGN_CENTER_SELECTION, CellStyle.VERTICAL_BOTTOM);
        //ALIGN_FILL水平对齐填充方式会复制同样的单元格内容来横向填满单元格
        createCell(wb, row2, (short) 2, CellStyle.ALIGN_FILL, CellStyle.VERTICAL_CENTER);
        createCell(wb, row2, (short) 3, CellStyle.ALIGN_GENERAL, CellStyle.VERTICAL_CENTER);
        createCell(wb, row2, (short) 4, CellStyle.ALIGN_JUSTIFY, CellStyle.VERTICAL_JUSTIFY);
        createCell(wb, row2, (short) 5, CellStyle.ALIGN_LEFT, CellStyle.VERTICAL_TOP);
        createCell(wb, row2, (short) 6, CellStyle.ALIGN_RIGHT, CellStyle.VERTICAL_TOP);

        /**
         * 设置单元格边框
         */
        Row row3 = st1.createRow(3);
        cell = row3.createCell(1);
        cell.setCellValue("单元格边框设置");
        //为单元格上下左右的全部边框设置线宽，和线色.
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(CellStyle.BORDER_DASH_DOT);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THICK);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);

        /**
         * 设置单元格的前景和背景填充色
         */
        Row row4 = st1.createRow(4);
        //浅绿背景色
        style = wb.createCellStyle();
        //设置背景色
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        //设置填充图案
        style.setFillPattern(CellStyle.BIG_SPOTS);
        cell = row4.createCell(1);
        cell.setCellValue("背景色填充");
        cell.setCellStyle(style);
        //桔黄前景色, 填充前景色不是字体颜色
        style = wb.createCellStyle();
        //设置背景色
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        //设置填充图案
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        cell = row4.createCell(2);
        cell.setCellValue("前景色填充");
        cell.setCellStyle(style);

        /**
         * 单元格合并，与Excel一致，合并后只能保存左上角单元格的内容
         */
        Row row5 = st1.createRow(5);
        cell = row5.createCell(1);
        cell.setCellValue("被合并单元格1");
        cell = row5.createCell(2);
        cell.setCellValue("被合并单元格2");
        Row row6 = st1.createRow(6);
        cell = row6.createCell(1);
        cell.setCellValue("被合并单元格3");
        cell = row6.createCell(2);
        cell.setCellValue("被合并单元格4");

        //CellRangeAddress region = CellRangeAddress.valueOf("B2:E5"); 也可以这样获得CellRangeAddress实例
        st1.addMergedRegion(new CellRangeAddress(
                5, //first row (0-based)
                6, //last row  (0-based)
                1, //first column (0-based)
                2  //last column  (0-based)
        ));


        /**
         * 设置单元格字体
         */
        Row row7 = st1.createRow(7);
        //创建字体并设置相关字体属性
        Font font = wb.createFont();
        //设置字号大小
        font.setFontHeightInPoints((short) 24);
        //设置字体名称
        font.setFontName("宋体");
        //设置字形是否斜体
        font.setItalic(true);
        //设置是否加删除线
        font.setStrikeout(true);
        //设置字形是否粗体
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        //设置下划线类型
        font.setUnderline(Font.U_DOUBLE);
        //设置字体颜色，Font类中预设的颜色常量只有Normal和Red两种，要自定义的话可以使用org.apache.poi.ss.usermodel.IndexedColors枚举来指定
        font.setColor(Font.COLOR_RED);
        // Fonts are set into a style so create a new one to use.
        style = wb.createCellStyle();
        style.setFont(font);
        // Create a cell and put a value in it.
        cell = row7.createCell(1);
        cell.setCellValue("单元格字体测试1");
        cell.setCellStyle(style);

        font = wb.createFont();
        font.setColor(IndexedColors.PINK.getIndex());
        font.setFontHeightInPoints((short) 10);
        font.setFontName("华文琥珀");

        style = wb.createCellStyle();
        style.setFont(font);

        cell = row7.createCell(2);
        cell.setCellValue("单元格字体测试2");
        cell.setCellStyle(style);

        /**
         * 设置单元格内容自动换行和列宽自动适应
         */
        Row row8 = st1.createRow(8);
        cell = row8.createCell(2);
        cell.setCellValue("在文件中包含 \n 换行符同时设置CellStyle的WrapText=true以在单元格内换行.\n" +
                "通过Sheet的autoSizeColumn设置指定列的列宽是否自动适应单元格内容，并有可选的合并操作");
        cs = wb.createCellStyle();
        cs.setWrapText(true);
        cell.setCellStyle(cs);
        //设置行高为4倍默认行高，以查看单元格内容的换行效果
        row8.setHeightInPoints((4 * st1.getDefaultRowHeightInPoints()));


        /**
         * 设置指定列的列宽为自动适应，对于公式列只有当公式计算后结果被缓存时才会按结果进行列宽自适应。
         * 列宽为自动适应依赖于java2D类群，如果图形环境不可用需要手工指定JVM启动参数java.awt.headless=true
         */
        st1.autoSizeColumn((short) 2);
        /**
         * 设置单元格的数据格式化格式(DataFormat)
         */
        Row row9 = st1.createRow(9);
        DataFormat format = wb.createDataFormat();
        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("0.0"));
        cell = row9.createCell(0);
        cell.setCellValue(11111.2526457);
        cell.setCellStyle(style);

        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("#,##0.0000"));
        cell = row9.createCell(1);
        cell.setCellValue(11111.2526457);
        cell.setCellStyle(style);

        style = wb.createCellStyle();
        style.setDataFormat(format.getFormat("M/d/yy HH:mm"));
        cell = row9.createCell(2);
        cell.setCellValue(new Date());
        cell.setCellStyle(style);


        /**
         * 使用便捷的工具类，RegionUtil来为合并的单元格设置格式属性，使用CellUtil来创建单元格或设置单元格格式属性等设置单元格格式时如果新增格式后的
         * CellStyle存在就不会再新增而使用已经存在的
         */
        Row row10 = st1.createRow(10);
        Row row11 = st1.createRow(11);
        cell = row10.createCell(1);
        cell.setCellValue("使用RegionUtil设置合并单元的边框");
        CellRangeAddress region = CellRangeAddress.valueOf("B11:E14");
        st1.addMergedRegion(region);
        // Set the border and border colors.
        final short borderMediumDashed = CellStyle.BORDER_MEDIUM_DASHED;
        RegionUtil.setBorderBottom(borderMediumDashed, region, st1, wb);
        RegionUtil.setBorderTop(borderMediumDashed, region, st1, wb);
        RegionUtil.setBorderLeft(borderMediumDashed, region, st1, wb);
        RegionUtil.setBorderRight(borderMediumDashed, region, st1, wb);
        RegionUtil.setBottomBorderColor(IndexedColors.AQUA.getIndex(), region, st1, wb);
        RegionUtil.setTopBorderColor(IndexedColors.AQUA.getIndex(), region, st1, wb);
        RegionUtil.setLeftBorderColor(IndexedColors.AQUA.getIndex(), region, st1, wb);
        RegionUtil.setRightBorderColor(IndexedColors.AQUA.getIndex(), region, st1, wb);
        // Shows some usages of HSSFCellUtil
        style = wb.createCellStyle();
        style.setIndention((short) 4);
        CellUtil.createCell(row10, 8, "使用CellUtil工具类创建的单元格内容1", style);
        cell = CellUtil.createCell(row11, 8, "使用CellUtil工具类创建的单元格内容2");
        CellUtil.setAlignment(cell, wb, CellStyle.ALIGN_CENTER);

        /**
         * 将工作表Sheet中的指定范围内的行，向上或向下移动指定的行数，第1个参数是开始行索引，第2上参数是结束行索引，第3个参数是
         * 移动的行数，正数表示向下移动，负数表示向上移动(实测发现移动后xls中被移动的原行位置变空，原行替换目标行。xlsx移动后打开错误)
         */
        //st2.shiftRows(5, 5, 2, false, false);

        /**
         * 设置工作表Sheet的缩放比率，以分数来表示，第1个参数是分子，第2个参数是分母，如缩放至75%则setZoom(3, 4)或setZoom(75, 100)
         */
        st2.setZoom(3, 4);

        /**
         * 设置行列冻结和设置编辑界面拆分，一个Sheet中只能存在1个FreezPane或一个SplitPane，新增的任一个Pane都将替换之前存在的那一个
         * 如果前两个参数都为0，则将取消所有的Pane
         */
        //第1个参数列索引以该列左边为冻结起始位置(所以0表示不冻结)，第2个参数行索引以该行上边为冻结起始位置(所以0表示不冻结)，
        //第3个参数表示非列冻结窗口起始显示的列的索引，第4个参数表示非行冻结起始显示的行的索引。一般使用只有前2个参数的重载方法
        st1.createFreezePane(3, 2);
        // Freeze just one column
        st2.createFreezePane(1, 0, 1, 0);
        //即Excel[视图]中的[拆分]功能，以某个点为中心，将编辑界面拆分成4个界面，每个界面都能单独滚动但编辑的是同一个Sheet，有点类似html中的frame
        //参数1、参数2是缇，指定拆分中心点的横纵坐标，绝对定位与单元格无关。
        st3.createSplitPane(2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT);


        /**
         * 设置工作表Sheet中的图形、图片、文本框、单元格备注、图形组
         */
        //图形处理目前POI3.8还没能完全统一API调用，只有几个简单图形和属性能够使用统一API调用。
        if (st1 instanceof HSSFSheet) {

            HSSFSheet hst = (HSSFSheet) st1;
            HSSFPatriarch hp = hst.createDrawingPatriarch();
            //在指定的锚点位置上创建一个图形
            HSSFSimpleShape shape = hp.createSimpleShape(hp.createAnchor(0, 0, 0, 0, 1, 14, 2, 16));
            //设置图形的类型为线
            shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);
            //设置填充颜色
            shape.setFillColor(IndexedColors.BLUE_GREY.getIndex());
            //设置边框线型
            shape.setLineStyle(HSSFSimpleShape.LINESTYLE_LONGDASHDOTDOTGEL);
            //设置边框线颜色
            shape.setLineStyleColor(255, 0, 0);
            //设置边框线宽
            shape.setLineWidth(HSSFSimpleShape.LINEWIDTH_ONE_PT * 3);

            //在指定的锚点位置上创建一个文本框
            HSSFTextbox textbox = hp.createTextbox(new HSSFClientAnchor(250, 125, 750, 125, (short) 2, 14, (short) 2, 16));
            font = wb.createFont();
            font.setItalic(true);
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            font.setUnderline(Font.U_DOUBLE);
            HSSFRichTextString string = new HSSFRichTextString("富文本<br>测试富文本");
            //对富文本中的字符设置字体，参数1字符开始索引(包含)，参数2字符的结束索引(不包含)，参数3字体
            string.applyFont(0, 3, font);
            textbox.setString(string);
            textbox.setLineStyleColor(0, 0, 255);
            textbox.setFillColor(200, 200, 200);

            //在指定锚点上创建图形组
            HSSFShapeGroup shapeGroup = hp.createGroup(hp.createAnchor(0, 0, 1023, 255, 3, 14, 4, 15));
            //在图形组中指定锚点位置上创建图形(ChildAnchor与ClientAnchor不同没有col1、row1、col2、row2，因为其坐标是相对图形组内部来说的,
            //默认其横纵坐标与ClientAnchor一样都是0-1023、0-255，但可以通过调用)
            //shapeGroup.setCoordinates(int x1, int y1, int x2, int y2)来指定图形组内的相对坐标范围，所有组内图形必须使用指定的坐标范围定位
            shape = shapeGroup.createShape(new HSSFChildAnchor(0, 0, 500, 125));
            shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
            ////在图形组中指定锚点位置上创建文本框
            textbox = shapeGroup.createTextbox(new HSSFChildAnchor(500, 125, 1023, 255));

            //在指定锚点位置上创建图片(目前POI3.8版本中支持PNG、JPG、DIB格式图片)
            InputStream is = new FileInputStream("C:\\Users\\Public\\Pictures\\Sample Pictures\\543617.jpg");
            byte[] bytes = IOUtils.toByteArray(is);
            is.close();
            //先将图片统一保存在Workbook中并取得存储索引
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            //创建锚点信息
            ClientAnchor anchor = createHelper.createClientAnchor();
            anchor.setCol1(6);
            anchor.setRow1(14);
            anchor.setCol2(8);
            anchor.setRow2(22);
            Picture pict = hp.createPicture(anchor, pictureIdx);
            //调用下面2个方法会使得图片按原始大小被重置，目前只有PNG、JPG支持重置
            //pict.resize();
            pict.getPreferredSize();

            /**
             * 为工作表Sheet设置单元格备注(批注)，批注相当于是一个文本框，也需要由Patriarch创建出来，并绑定到某个单元格上
             */
            Row row17 = st1.createRow(17);
            cell = row17.createCell(1);
            cell.setCellValue("单元格备注测试");
            // When the comment box is visible, have it show in a 1x3 space
            anchor = createHelper.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex() + 1);
            anchor.setCol2(cell.getColumnIndex() + 2);
            anchor.setRow1(row17.getRowNum());
            anchor.setRow2(row17.getRowNum() + 3);
            // Create the comment and set the text+author
            Comment comment = hp.createCellComment(anchor);
            RichTextString str = createHelper.createRichTextString("单元格批注内容！");
            comment.setString(str);
            comment.setAuthor("Apache POI:李国庆测试");
            //为单元格绑定批注
            cell.setCellComment(comment);


            /**
             * 为单元格创建超链接
             */
            //cell style for hyperlinks
            //by default hypelrinks are blue and underlined
            CellStyle hlink_style = wb.createCellStyle();
            Font hlink_font = wb.createFont();
            hlink_font.setUnderline(Font.U_SINGLE);
            hlink_font.setColor(IndexedColors.BLUE.getIndex());
            hlink_style.setFont(hlink_font);
            Row row18 = st1.createRow(18);
            cell = row18.createCell((short) 0);
            cell.setCellValue("URL超链接");
            Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_URL);
            link.setAddress("http://poi.apache.org/");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            //link to a file in the current directory
            cell = row18.createCell((short) 1);
            cell.setCellValue("文件超链接");
            link = createHelper.createHyperlink(Hyperlink.LINK_FILE);
            link.setAddress("D:\\LGQ-JAVA\\OfficeParse\\WebRoot\\files\\TestTest.xls");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            //e-mail link
            cell = row18.createCell((short) 2);
            cell.setCellValue("EMAIL超链接");
            link = createHelper.createHyperlink(Hyperlink.LINK_EMAIL);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.setAddress("mailto:awm96@163.com");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            //link to a place in this workbook
            //create a target sheet and cell
            cell = row18.createCell((short) 3);
            cell.setCellValue("文档内超链接跳转");
            Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
            link2.setAddress("'Sheet打印格式测试'!A1");
            cell.setHyperlink(link2);
            cell.setCellStyle(hlink_style);

        } else if (st1 instanceof XSSFSheet) {

            XSSFSheet xst = (XSSFSheet) st1;
            XSSFDrawing xd = xst.createDrawingPatriarch();
            //在指定的锚点位置上创建一个图形，注意：XSSFClientAnchor的dx1、dy1、dx2、dy2定义与HSSF不同，不再是0-1023、0-255的相对比率，而是
            //一个固定单位的绝对定位，通常使用XSSFShape.EMU_PER_POINT(折合成1点)和XSSFShape.EMU_PER_PIXEL(折合成1像素)做为单位乘以1个值来定位
            XSSFSimpleShape shape = xd.createSimpleShape(xd.createAnchor(0, 0, 0, 0, 1, 14, 2, 16));
            //设置图形的类型为线
            shape.setShapeType(ShapeTypes.LINE);
            //设置填充颜色
            shape.setFillColor(0, 0, 250);
            //设置边框线型：solid=0、dot=1、dash=2、lgDash=3、dashDot=4、lgDashDot=5、lgDashDotDot=6、sysDash=7、sysDot=8、sysDashDot=9、sysDashDotDot=10
            shape.setLineStyle(7);
            //设置边框线颜色
            shape.setLineStyleColor(255, 0, 0);
            //设置边框线宽,单位Point
            shape.setLineWidth(2);

            //在指定的锚点位置上创建一个文本框
            XSSFTextBox textbox = xd.createTextbox(
                    xd.createAnchor(
                            XSSFShape.EMU_PER_POINT * 10, XSSFShape.EMU_PER_PIXEL * 5, XSSFShape.EMU_PER_POINT * 200,
                            (int) (XSSFShape.EMU_PER_POINT * (st1.getRow(16) == null ? st1.getDefaultRowHeightInPoints() : st1.getRow(16).getHeightInPoints())),
                            2, 14, 2, 16
                    )
            );
            font = wb.createFont();
            font.setItalic(true);
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            font.setUnderline(Font.U_DOUBLE);
            XSSFRichTextString string = new XSSFRichTextString("富文本<br>测试富文本");
            //对富文本中的字符设置字体，参数1字符开始索引(包含)，参数2字符的结束索引(不包含)，参数3字体
            string.applyFont(0, 3, font);
            textbox.setText(string);
            textbox.setLineStyleColor(0, 0, 255);
            textbox.setFillColor(200, 200, 200);

            //在指定锚点上创建图形组
            XSSFShapeGroup shapeGroup = xd.createGroup(xd.createAnchor(0, 0, 0, 0, 3, 14, 5, 16));
            //在图形组中指定锚点位置上创建图形,ChildAnchor与ClientAnchor不同没有col1、row1、col2、row2
            //注意：XSSFChildAnchor与HSSF不同，默认其横纵坐标都是相对于整个Sheet的左上角顶点的，需要调用
            //shapeGroup.setCoordinates(int x1, int y1, int x2, int y2)来指定图形组内的相对坐标范围，
            //并使得所有组内图形使用各自的ChildAnchor坐标范围来相对定位
            shapeGroup.setCoordinates(0, 0, 100, 100);
            //shape = shapeGroup.createSimpleShape(new XSSFChildAnchor(0,0,XSSFShape.EMU_PER_POINT*50,XSSFShape.EMU_PER_POINT*20));
            shape = shapeGroup.createSimpleShape(new XSSFChildAnchor(0, 0, 50, 50));
            shape.setShapeType(ShapeTypes.CUBE);
            ////在图形组中指定锚点位置上创建文本框
            //textbox = shapeGroup.createTextbox(new XSSFChildAnchor(XSSFShape.EMU_PER_POINT*50,XSSFShape.EMU_PER_POINT*20,XSSFShape.EMU_PER_POINT*150,XSSFShape.EMU_PER_POINT*50));
            textbox = shapeGroup.createTextbox(new XSSFChildAnchor(50, 50, 100, 100));

            //在指定锚点位置上创建图片(目前POI3.8版本中支持PNG、JPG、DIB格式图片)
            InputStream is = new FileInputStream("C:\\Users\\Public\\Pictures\\Sample Pictures\\543617.jpg");
            byte[] bytes = IOUtils.toByteArray(is);
            is.close();
            //先将图片统一保存在Workbook中并取得存储索引
            int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            //创建锚点信息
            ClientAnchor anchor = createHelper.createClientAnchor();
            anchor.setCol1(6);
            anchor.setRow1(14);
            anchor.setCol2(8);
            anchor.setRow2(22);
            Picture pict = xd.createPicture(anchor, pictureIdx);
            //调用下面2个方法会使得图片按原始大小被重置，目前只有PNG、JPG支持重置
            //pict.resize();
            pict.getPreferredSize();

            /**
             * 为工作表Sheet设置单元格备注(批注)，批注相当于是一个文本框，也需要由Patriarch创建出来，并绑定到某个单元格上
             */
            Row row17 = st1.createRow(17);
            cell = row17.createCell(1);
            cell.setCellValue("单元格备注测试");
            // When the comment box is visible, have it show in a 1x3 space
            anchor = createHelper.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex() + 1);
            anchor.setCol2(cell.getColumnIndex() + 2);
            anchor.setRow1(row17.getRowNum());
            anchor.setRow2(row17.getRowNum() + 3);
            // Create the comment and set the text+author
            Comment comment = xd.createCellComment(anchor);
            RichTextString str = createHelper.createRichTextString("单元格批注内容！");
            comment.setString(str);
            comment.setAuthor("Apache POI:李国庆测试");
            //为单元格绑定批注
            cell.setCellComment(comment);


            /**
             * 为单元格创建超链接
             */
            //cell style for hyperlinks
            //by default hypelrinks are blue and underlined
            CellStyle hlink_style = wb.createCellStyle();
            Font hlink_font = wb.createFont();
            hlink_font.setUnderline(Font.U_SINGLE);
            hlink_font.setColor(IndexedColors.BLUE.getIndex());
            hlink_style.setFont(hlink_font);
            Row row18 = st1.createRow(18);
            cell = row18.createCell((short) 0);
            cell.setCellValue("URL超链接");
            Hyperlink link = createHelper.createHyperlink(Hyperlink.LINK_URL);
            link.setAddress("http://poi.apache.org/");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            //link to a file in the current directory
            cell = row18.createCell((short) 1);
            cell.setCellValue("文件超链接");
            link = createHelper.createHyperlink(Hyperlink.LINK_FILE);
            link.setAddress("TestTest.xls");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            //e-mail link
            cell = row18.createCell((short) 2);
            cell.setCellValue("EMAIL超链接");
            link = createHelper.createHyperlink(Hyperlink.LINK_EMAIL);
            //note, if subject contains white spaces, make sure they are url-encoded
            link.setAddress("mailto:awm96@163.com");
            cell.setHyperlink(link);
            cell.setCellStyle(hlink_style);
            //link to a place in this workbook
            //create a target sheet and cell
            cell = row18.createCell((short) 3);
            cell.setCellValue("文档内超链接跳转");
            Hyperlink link2 = createHelper.createHyperlink(Hyperlink.LINK_DOCUMENT);
            link2.setAddress("'Sheet打印格式测试'!A1");
            cell.setHyperlink(link2);
            cell.setCellStyle(hlink_style);

        }
        /**
         * 为工作表Sheet的行和列设置分组，并可指定用Excel打开时分组是否折叠显示
         */
        st2.groupRow(5, 14);
        st2.groupRow(7, 14);
        st2.groupRow(16, 19);
        st2.groupColumn((short) 4, (short) 7);
        st2.groupColumn((short) 9, (short) 12);
        st2.groupColumn((short) 10, (short) 11);

        //设置Excel打开时分组是否折叠显示，参数1是分组的开始行或列索引，参数2是否折叠显示
        st2.setRowGroupCollapsed(7, true);
        st2.setColumnGroupCollapsed((short) 4, true);

        wb.write(fos);
        fos.close();
    }

    private static void createCell(Workbook wb, Row row, short column, short halign, short valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }

    private static void createCell(Sheet st, int rows, int cols) {

        for (int i = 0; i < rows; i++) {

            Row r = st.createRow(i);
            for (int j = 0; j < cols; j++) {

                Cell c = r.createCell(j);
                CellReference cellRef = new CellReference(i, j);
                c.setCellValue(cellRef.formatAsString());
            }
        }
    }




}
